"""GeoIntel Daily Dispatch Intelligence module.

Reads the mine's Daily Dispatch .xlsb workbook and turns the hourly dispatch
matrix into operational dashboards. The module is deliberately separate from
app.py so it can be developed without disturbing DCI and Red Rating.
"""
from __future__ import annotations

import io
import re
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image

try:
    from pyxlsb import open_workbook
except ImportError:  # shown clearly in the UI instead of crashing the whole app
    open_workbook = None

DISPATCH_HISTORY_FILE = Path("gdi_dispatch_history.csv")
DISPATCH_RING_FILE = Path("gdi_undercut_ring_positions.csv")
RED_RATING_FILE = Path("geointel_red_rating_tracker.csv")

PRODUCTIVE_ACTIVITIES = {
    "drilling", "supporting", "loading/lashing", "loading", "lashing",
    "cleaning", "barring/scaling", "scaling", "meshing", "charging",
    "blasting", "shotcrete", "shotcrete spraying", "face prep",
    "entry examination", "safe declaration", "longhole drilling",
}
DELAY_KEYWORDS = {
    "delay": "General delay",
    "equipment availability": "Equipment availability",
    "breakdown": "Equipment breakdown",
    "power supply": "Power supply",
    "water supply": "Water supply",
    "ventilation": "Ventilation",
    "idling face": "Idling face",
    "waiting": "Waiting",
    "seismic red": "Seismic red rating",
    "red rating": "Seismic red rating",
}
IGNORE_ACTIVITIES = {"", "none", "nan", "shift start", "shift end"}


def _load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.DataFrame()


def _excel_date(value: Any) -> pd.Timestamp | None:
    """Convert Excel serial/date-like value to Timestamp."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        try:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
        except Exception:
            return None
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    return None if pd.isna(parsed) else parsed


def _clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _delay_category(activity: str) -> str | None:
    text = _clean(activity).lower()
    for keyword, category in DELAY_KEYWORDS.items():
        if keyword in text:
            return category
    return None


def _is_productive(activity: str) -> bool:
    text = _clean(activity).lower()
    if text in IGNORE_ACTIVITIES or _delay_category(text):
        return False
    return any(term in text for term in PRODUCTIVE_ACTIVITIES)


def _sheet_rows(workbook_path: str, sheet_name: str) -> list[list[Any]]:
    rows: list[list[Any]] = []
    with open_workbook(workbook_path) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            for row in sheet.rows():
                rows.append([cell.v for cell in row])
    return rows


def _cell(rows: list[list[Any]], r: int, c: int, default: Any = None) -> Any:
    try:
        return rows[r][c]
    except (IndexError, TypeError):
        return default


def parse_dispatch_sheet(rows: list[list[Any]], sheet_name: str) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Parse one daily shift sheet into one row per area-hour."""
    date_value = _cell(rows, 5, 5)
    report_date = _excel_date(date_value)
    crew = _clean(_cell(rows, 5, 8))
    shift = _clean(_cell(rows, 5, 9)) or ("Day Shift" if sheet_name.upper().endswith("D") else "Night Shift")
    supervisor = _clean(_cell(rows, 5, 6))

    summary_headers = [_clean(v) for v in (rows[1] if len(rows) > 1 else [])]
    summary_values = rows[2] if len(rows) > 2 else []
    summary = {}
    for i, header in enumerate(summary_headers):
        if header:
            summary[header] = summary_values[i] if i < len(summary_values) else None

    heading_row = 7
    operator_row = 8
    machine_row = 9
    time_start_row = 10
    time_end_row = 21

    records: list[dict[str, Any]] = []
    max_cols = max((len(r) for r in rows[:23]), default=0)
    for col in range(2, max_cols):
        area = _clean(_cell(rows, heading_row, col))
        if not area or area.lower() in {"none", "nan"}:
            continue
        operator = _clean(_cell(rows, operator_row, col))
        machine = _clean(_cell(rows, machine_row, col))
        for r in range(time_start_row, min(time_end_row + 1, len(rows))):
            time_slot = _clean(_cell(rows, r, 1))
            activity = _clean(_cell(rows, r, col))
            if not time_slot or not activity:
                continue
            delay_category = _delay_category(activity)
            records.append({
                "Date": report_date.normalize() if report_date is not None else pd.NaT,
                "Sheet": sheet_name,
                "Crew": crew or "Unspecified",
                "Shift": shift or "Unspecified",
                "Supervisor": supervisor,
                "Area": area.upper(),
                "Operator": operator,
                "Machine": machine.upper(),
                "Time Slot": time_slot,
                "Activity": activity,
                "Productive": _is_productive(activity),
                "Delay": delay_category is not None,
                "Delay Category": delay_category or "",
                "Hours": 1.0,
            })

    meta = {
        "Date": report_date.normalize() if report_date is not None else pd.NaT,
        "Sheet": sheet_name,
        "Crew": crew or "Unspecified",
        "Shift": shift or "Unspecified",
        "Supervisor": supervisor,
        **summary,
    }
    return pd.DataFrame(records), meta


def parse_dispatch_workbook(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    if open_workbook is None:
        return pd.DataFrame(), pd.DataFrame(), ["pyxlsb is not installed. Run: python -m pip install pyxlsb"]

    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix != ".xlsb":
        return pd.DataFrame(), pd.DataFrame(), ["Please upload the original Daily Dispatch .xlsb workbook."]

    warnings: list[str] = []
    with tempfile.NamedTemporaryFile(suffix=".xlsb", delete=False) as temp:
        temp.write(uploaded_file.getbuffer())
        temp_path = temp.name

    all_activity: list[pd.DataFrame] = []
    all_meta: list[dict[str, Any]] = []
    try:
        with open_workbook(temp_path) as wb:
            sheets = list(wb.sheets)
        daily_sheets = [s for s in sheets if re.fullmatch(r"\d{1,2}[DN]", str(s).strip(), re.I)]
        if not daily_sheets:
            warnings.append("No daily sheets such as 18D, 18N, 19D or 19N were found.")
        for sheet_name in daily_sheets:
            try:
                activity, meta = parse_dispatch_sheet(_sheet_rows(temp_path, sheet_name), sheet_name)
                if not activity.empty:
                    all_activity.append(activity)
                all_meta.append(meta)
            except Exception as exc:
                warnings.append(f"{sheet_name}: could not be read ({exc}).")
    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except Exception:
            pass

    activity_df = pd.concat(all_activity, ignore_index=True) if all_activity else pd.DataFrame()
    meta_df = pd.DataFrame(all_meta)
    return activity_df, meta_df, warnings


def _merge_red_rating_delays(activity: pd.DataFrame) -> pd.DataFrame:
    """Append seismic red-rating restrictions as delay records."""
    red = _load_csv(RED_RATING_FILE)
    if red.empty or "Date" not in red.columns:
        return activity
    red["Date"] = pd.to_datetime(red["Date"], errors="coerce").dt.normalize()
    red["Duration Hours"] = pd.to_numeric(red.get("Duration Hours", 0), errors="coerce").fillna(0.0)
    rows = []
    for _, incident in red.dropna(subset=["Date"]).iterrows():
        locations = [x.strip().upper() for x in str(incident.get("Locations", "")).split(",") if x.strip()]
        if not locations:
            locations = ["UNSPECIFIED"]
        allocated = float(incident.get("Duration Hours", 0)) / len(locations)
        for location in locations:
            rows.append({
                "Date": incident["Date"], "Sheet": "Red Rating", "Crew": "All", "Shift": str(incident.get("Shift", "Unspecified")),
                "Supervisor": "", "Area": location, "Operator": "", "Machine": "", "Time Slot": "Red rating",
                "Activity": "Seismic Red Rating", "Productive": False, "Delay": True,
                "Delay Category": "Seismic red rating", "Hours": allocated,
            })
    if not rows:
        return activity
    red_rows = pd.DataFrame(rows)
    return pd.concat([activity, red_rows], ignore_index=True) if not activity.empty else red_rows


def _dominant_activity(group: pd.DataFrame) -> str:
    useful = group[~group["Activity"].str.lower().isin(IGNORE_ACTIVITIES)]
    if useful.empty:
        return "No recorded work"
    return str(useful.groupby("Activity")["Hours"].sum().sort_values(ascending=False).index[0])


def build_area_summary(activity: pd.DataFrame, selected_date: pd.Timestamp) -> pd.DataFrame:
    day = activity[activity["Date"] == selected_date].copy()
    if day.empty:
        return pd.DataFrame()
    rows = []
    for area, group in day.groupby("Area"):
        productive = float(group.loc[group["Productive"], "Hours"].sum())
        delay = float(group.loc[group["Delay"], "Hours"].sum())
        total = float(group["Hours"].sum())
        rows.append({
            "Area": area,
            "Crew": ", ".join(sorted(set(group["Crew"].dropna().astype(str))))[:80],
            "Shift": ", ".join(sorted(set(group["Shift"].dropna().astype(str))))[:80],
            "Machine": ", ".join(sorted(x for x in set(group["Machine"].astype(str)) if x))[:80],
            "Operator": ", ".join(sorted(x for x in set(group["Operator"].astype(str)) if x))[:80],
            "Main Work": _dominant_activity(group),
            "Productive Hours": productive,
            "Delay Hours": delay,
            "Recorded Hours": total,
            "Productivity %": round(100 * productive / total, 1) if total else 0.0,
            "Status": "Active" if productive > 0 else ("Delayed" if delay > 0 else "No productive work"),
        })
    return pd.DataFrame(rows).sort_values(["Status", "Area"])


def build_seven_day_tracker(activity: pd.DataFrame, selected_date: pd.Timestamp) -> pd.DataFrame:
    start = selected_date - pd.Timedelta(days=6)
    window = activity[(activity["Date"] >= start) & (activity["Date"] <= selected_date)].copy()
    if window.empty:
        return pd.DataFrame()
    all_areas = sorted(window["Area"].dropna().unique())
    rows = []
    for area in all_areas:
        area_data = window[window["Area"] == area]
        productive_days = sorted(area_data.loc[area_data["Productive"], "Date"].dropna().unique())
        last = max(productive_days) if productive_days else pd.NaT
        days_inactive = 7 if pd.isna(last) else int((selected_date - pd.Timestamp(last)).days)
        row = {
            "Area": area,
            "Days Worked": len(productive_days),
            "Last Worked": "Never in window" if pd.isna(last) else pd.Timestamp(last).strftime("%Y-%m-%d"),
            "Days Since Productive Work": days_inactive,
            "7-Day Productive Hours": float(area_data.loc[area_data["Productive"], "Hours"].sum()),
            "7-Day Delay Hours": float(area_data.loc[area_data["Delay"], "Hours"].sum()),
            "Activity Status": "Inactive 7+ days" if days_inactive >= 7 else "Inactive 4–6 days" if days_inactive >= 4 else "Inactive 1–3 days" if days_inactive >= 1 else "Worked today",
        }
        for offset in range(6, -1, -1):
            date = selected_date - pd.Timedelta(days=offset)
            key = date.strftime("%d %b")
            d = area_data[area_data["Date"] == date]
            row[key] = "●" if d["Productive"].any() else ("D" if d["Delay"].any() else "—")
        rows.append(row)
    return pd.DataFrame(rows).sort_values(["Days Since Productive Work", "Area"], ascending=[False, True])


def _read_ring_upload(uploaded) -> pd.DataFrame:
    if uploaded is None:
        return pd.DataFrame()
    name = uploaded.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(uploaded)
        return pd.read_excel(uploaded)
    except Exception as exc:
        st.error(f"Could not read ring-position file: {exc}")
        return pd.DataFrame()


def prepare_ring_table(df: pd.DataFrame, selected_date: pd.Timestamp) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Date", "Heading", "Ring"])
    work = df.copy()
    aliases = {c.lower().strip(): c for c in work.columns}
    heading_col = next((aliases[k] for k in aliases if k in {"heading", "area", "undercut heading"}), None)
    ring_col = next((aliases[k] for k in aliases if k in {"ring", "ring number", "current ring"}), None)
    date_col = next((aliases[k] for k in aliases if k == "date"), None)
    if heading_col is None or ring_col is None:
        return pd.DataFrame(columns=["Date", "Heading", "Ring"])
    out = pd.DataFrame({
        "Date": pd.to_datetime(work[date_col], errors="coerce").dt.normalize() if date_col else selected_date,
        "Heading": work[heading_col].astype(str).str.strip().str.upper(),
        "Ring": pd.to_numeric(work[ring_col], errors="coerce"),
    }).dropna(subset=["Heading", "Ring"])
    return out


def lead_lag_table(rings: pd.DataFrame, selected_date: pd.Timestamp, limit: int, warning: int) -> pd.DataFrame:
    if rings.empty:
        return pd.DataFrame()
    valid = rings[rings["Date"] <= selected_date].sort_values("Date").drop_duplicates("Heading", keep="last").copy()
    valid = valid[valid["Heading"].str.contains("UC", case=False, na=False)]
    if valid.empty:
        return pd.DataFrame()
    def heading_order(text: str) -> tuple:
        nums = re.findall(r"\d+", text)
        return (int(nums[0]) if nums else 9999, text)
    valid["_order"] = valid["Heading"].map(heading_order)
    valid = valid.sort_values("_order").reset_index(drop=True)
    valid["Adjacent Heading"] = valid["Heading"].shift(1)
    valid["Adjacent Ring"] = valid["Ring"].shift(1)
    valid["Lead/Lag Rings"] = valid["Ring"] - valid["Adjacent Ring"]
    valid["Status"] = valid["Lead/Lag Rings"].abs().map(
        lambda x: "No comparison" if pd.isna(x) else "Outside limit" if x > limit else "Approaching limit" if x >= warning else "Within limit"
    )
    valid["Direction"] = valid["Lead/Lag Rings"].map(lambda x: "—" if pd.isna(x) else "Lead" if x > 0 else "Lag" if x < 0 else "Aligned")
    return valid.drop(columns=["_order"])


def mine_health_index(activity: pd.DataFrame, area_summary: pd.DataFrame, tracker: pd.DataFrame, leadlag: pd.DataFrame, selected_date: pd.Timestamp) -> tuple[int, str, dict[str, float]]:
    day = activity[activity["Date"] == selected_date]
    productive = float(day.loc[day["Productive"], "Hours"].sum())
    delays = float(day.loc[day["Delay"], "Hours"].sum())
    denominator = productive + delays
    productivity_score = 100 * productive / denominator if denominator else 50
    active_score = 100 * (area_summary["Productive Hours"] > 0).mean() if not area_summary.empty else 50
    inactivity_score = 100 - min(100, 12 * int((tracker["Days Since Productive Work"] >= 4).sum())) if not tracker.empty else 50
    leadlag_score = 100 - min(100, 30 * int((leadlag["Status"] == "Outside limit").sum()) + 12 * int((leadlag["Status"] == "Approaching limit").sum())) if not leadlag.empty else 70
    components = {
        "Productive time": productivity_score,
        "Active areas": active_score,
        "7-day continuity": inactivity_score,
        "Undercut sequence": leadlag_score,
    }
    score = int(round(sum(components.values()) / len(components)))
    label = "Healthy" if score >= 80 else "Watch" if score >= 65 else "Concern" if score >= 45 else "Critical"
    return score, label, components


def dispatch_intelligence_text(activity: pd.DataFrame, area_summary: pd.DataFrame, tracker: pd.DataFrame, leadlag: pd.DataFrame, selected_date: pd.Timestamp) -> str:
    day = activity[activity["Date"] == selected_date]
    if day.empty:
        return "No dispatch activity is available for the selected date."
    productive_hours = float(day.loc[day["Productive"], "Hours"].sum())
    delay_hours = float(day.loc[day["Delay"], "Hours"].sum())
    dominant = "not identified"
    useful = day[day["Productive"]]
    if not useful.empty:
        dominant = useful.groupby("Activity")["Hours"].sum().sort_values(ascending=False).index[0]
    top_delay = "none recorded"
    delayed = day[day["Delay"]]
    if not delayed.empty:
        top_delay = delayed.groupby("Delay Category")["Hours"].sum().sort_values(ascending=False).index[0]
    idle7 = int((tracker["Days Since Productive Work"] >= 7).sum()) if not tracker.empty else 0
    outside = int((leadlag["Status"] == "Outside limit").sum()) if not leadlag.empty else 0
    worst_area = "not identified"
    if not area_summary.empty and area_summary["Delay Hours"].sum() > 0:
        worst_area = area_summary.sort_values("Delay Hours", ascending=False).iloc[0]["Area"]
    return (
        f"On {selected_date.strftime('%d %B %Y')}, {productive_hours:.1f} productive area-hours and {delay_hours:.1f} delay-hours were recorded. "
        f"The dominant work activity was {dominant}. The leading delay category was {top_delay}, with {worst_area} carrying the highest recorded delay burden. "
        f"The seven-day review identifies {idle7} area(s) with no productive work in the full window. "
        f"Undercut lead-and-lag monitoring identifies {outside} adjacent comparison(s) outside the configured limit. "
        f"Priority should be given to restoring inactive headings, resolving recurring delay causes—including seismic red-rating restrictions—and protecting the undercut sequence."
    )


def render_daily_dispatch(red_rating_file: Path | str = RED_RATING_FILE) -> None:
    """Render the complete Daily Dispatch Intelligence tab."""
    global RED_RATING_FILE
    RED_RATING_FILE = Path(red_rating_file)

    st.markdown(
        '<div style="background:linear-gradient(135deg,#123047,#176b87);color:white;border-radius:14px;padding:1rem 1.15rem;margin:.15rem 0 .75rem 0;">'
        '<div style="font-size:1.45rem;font-weight:900;">🚧 Daily Dispatch Intelligence</div>'
        '<div style="font-size:.86rem;opacity:.92;">What happened underground, where time was lost, and what needs attention next.</div></div>',
        unsafe_allow_html=True,
    )

    upload = st.file_uploader("Upload Daily Dispatch report (.xlsb)", type=["xlsb"], key="dispatch_xlsb")
    if upload is None:
        st.info("Upload the daily dispatch workbook to activate all dashboards. Seismic red-rating delays will be merged automatically from the Red Rating Tracker.")
        return

    with st.spinner("Reading daily dispatch workbook…"):
        activity, meta, warnings = parse_dispatch_workbook(upload)
    for warning in warnings:
        st.warning(warning)
    if activity.empty:
        st.error("No hourly dispatch records could be extracted from this workbook.")
        return

    activity = _merge_red_rating_delays(activity)
    activity["Date"] = pd.to_datetime(activity["Date"], errors="coerce").dt.normalize()
    available_dates = sorted(activity["Date"].dropna().unique(), reverse=True)
    selected_date = pd.Timestamp(st.selectbox("Reporting date", available_dates, format_func=lambda x: pd.Timestamp(x).strftime("%d %B %Y"), key="dispatch_date")).normalize()

    area_summary = build_area_summary(activity, selected_date)
    tracker = build_seven_day_tracker(activity, selected_date)

    # Ring/map setup is shared by the lead-lag dashboard and the health index.
    saved_rings = _load_csv(DISPATCH_RING_FILE)
    if not saved_rings.empty:
        saved_rings["Date"] = pd.to_datetime(saved_rings["Date"], errors="coerce").dt.normalize()
        saved_rings["Ring"] = pd.to_numeric(saved_rings["Ring"], errors="coerce")
    limit = 10
    warning_limit = 8
    leadlag = lead_lag_table(saved_rings, selected_date, limit, warning_limit)
    health_score, health_label, health_components = mine_health_index(activity, area_summary, tracker, leadlag, selected_date)

    day = activity[activity["Date"] == selected_date]
    active_areas = int((area_summary["Productive Hours"] > 0).sum()) if not area_summary.empty else 0
    inactive_areas = int((tracker["Days Since Productive Work"] >= 7).sum()) if not tracker.empty else 0
    delay_hours = float(day.loc[day["Delay"], "Hours"].sum())
    seismic_hours = float(day.loc[day["Delay Category"] == "Seismic red rating", "Hours"].sum())

    # 1. Daily summary + bonus health index
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Mine Health Index", f"{health_score}/100", health_label)
    k2.metric("Active areas", active_areas)
    k3.metric("Inactive 7+ days", inactive_areas)
    k4.metric("Productive hours", f"{day.loc[day['Productive'], 'Hours'].sum():.1f} h")
    k5.metric("Total delays", f"{delay_hours:.1f} h")
    k6.metric("Seismic red delays", f"{seismic_hours:.1f} h")

    tabs = st.tabs([
        "📋 Daily Overview", "🗺️ Areas Worked", "🔨 Work Types", "⚠️ Delays",
        "📅 7-Day Activity", "📏 Undercut Lead/Lag", "🚜 Equipment",
        "👷 Crew & Shift", "📈 Trends", "🧠 Daily Intelligence", "⭐ Health Index"
    ])

    # 1. Daily Operations Summary
    with tabs[0]:
        st.subheader("Daily Operations Summary")
        day_meta = meta[pd.to_datetime(meta["Date"], errors="coerce").dt.normalize() == selected_date] if not meta.empty else pd.DataFrame()
        if not day_meta.empty:
            show_cols = [c for c in ["Sheet", "Crew", "Shift", "Supervisor", "Headings Blasted", "Rings Blasted", "Total Truck Loads", "Buckets to Mini Crusher", "Shotcrete", "Long Anchors Meters Drilled", "Longholes Drilled (m)"] if c in day_meta.columns]
            st.dataframe(day_meta[show_cols], width="stretch", hide_index=True)
        st.plotly_chart(px.bar(area_summary, x="Area", y=["Productive Hours", "Delay Hours"], barmode="stack", title="Productive and delay hours by area"), width="stretch")

    # 2. Areas worked today
    with tabs[1]:
        st.subheader("Areas Worked Today")
        status_filter = st.multiselect("Status", sorted(area_summary["Status"].unique()), default=sorted(area_summary["Status"].unique()), key="dispatch_area_status")
        st.dataframe(area_summary[area_summary["Status"].isin(status_filter)], width="stretch", hide_index=True)

    # 3. Work Type Intelligence
    with tabs[2]:
        st.subheader("Work Type Intelligence")
        productive_day = day[day["Productive"]].groupby("Activity", as_index=False)["Hours"].sum().sort_values("Hours", ascending=False)
        c1, c2 = st.columns(2)
        if productive_day.empty:
            c1.info("No productive work types recorded.")
        else:
            c1.plotly_chart(px.pie(productive_day, names="Activity", values="Hours", hole=.45, title="Productive effort distribution"), width="stretch")
            by_area = day[day["Productive"]].groupby(["Area", "Activity"], as_index=False)["Hours"].sum()
            c2.plotly_chart(px.bar(by_area, x="Area", y="Hours", color="Activity", title="Work types by area"), width="stretch")
        st.dataframe(productive_day, width="stretch", hide_index=True)

    # 4. Breakdowns & Delays incl seismic red rating
    with tabs[3]:
        st.subheader("Breakdowns, Delays & Seismic Red Ratings")
        delay_day = day[day["Delay"]].copy()
        if delay_day.empty:
            st.success("No delay hours were recorded for this date.")
        else:
            d1 = delay_day.groupby("Delay Category", as_index=False)["Hours"].sum().sort_values("Hours", ascending=False)
            d2 = delay_day.groupby("Area", as_index=False)["Hours"].sum().sort_values("Hours", ascending=False)
            c1, c2 = st.columns(2)
            c1.plotly_chart(px.bar(d1, x="Delay Category", y="Hours", title="Delay hours by cause"), width="stretch")
            c2.plotly_chart(px.bar(d2, x="Area", y="Hours", title="Delay burden by area"), width="stretch")
            st.dataframe(delay_day[["Area", "Shift", "Machine", "Time Slot", "Delay Category", "Hours"]], width="stretch", hide_index=True)

    # 5. Seven-day activity tracker
    with tabs[4]:
        st.subheader("Seven-Day Area Activity")
        st.caption("● = productive work, D = delay only, — = no dispatch record")
        st.dataframe(tracker, width="stretch", hide_index=True)
        inactive = tracker[tracker["Days Since Productive Work"] >= 4]
        if not inactive.empty:
            st.warning("Attention areas: " + ", ".join(inactive["Area"].astype(str).tolist()))

    # 6. Undercut lead/lag using ring numbers from map/manual table
    with tabs[5]:
        st.subheader("Undercut Lead & Lag")
        st.caption("Upload the undercut map for visual reference, then enter or import the current ring at each undercut heading. Adjacent headings are compared automatically.")
        map_upload = st.file_uploader("Undercut ring map (PNG/JPG)", type=["png", "jpg", "jpeg"], key="dispatch_ring_map")
        if map_upload:
            st.image(Image.open(map_upload), caption="Undercut ring map", width="stretch")
        s1, s2 = st.columns(2)
        limit = int(s1.number_input("Maximum allowed lead/lag (rings)", min_value=1, value=10, step=1, key="dispatch_limit"))
        warning_limit = int(s2.number_input("Approaching-limit warning (rings)", min_value=1, max_value=limit, value=min(8, limit), step=1, key="dispatch_warning"))
        ring_upload = st.file_uploader("Optional ring-position table (CSV/XLSX: Date, Heading, Ring)", type=["csv", "xlsx", "xls"], key="dispatch_ring_table")
        imported_rings = prepare_ring_table(_read_ring_upload(ring_upload), selected_date)
        base = saved_rings.copy()
        if not imported_rings.empty:
            base = pd.concat([base, imported_rings], ignore_index=True) if not base.empty else imported_rings
        current = base[base["Date"] <= selected_date].sort_values("Date").drop_duplicates("Heading", keep="last") if not base.empty else pd.DataFrame(columns=["Date", "Heading", "Ring"])
        if current.empty:
            current = pd.DataFrame({"Date": [selected_date] * 4, "Heading": ["UC21S", "UC22S", "UC23S", "UC24S"], "Ring": [0, 0, 0, 0]})
        edited = st.data_editor(current[["Date", "Heading", "Ring"]], num_rows="dynamic", width="stretch", key="dispatch_ring_editor")
        if st.button("Save ring positions", type="primary", key="save_dispatch_rings"):
            clean = prepare_ring_table(edited, selected_date)
            combined = pd.concat([saved_rings, clean], ignore_index=True) if not saved_rings.empty else clean
            combined = combined.drop_duplicates(["Date", "Heading"], keep="last")
            combined.to_csv(DISPATCH_RING_FILE, index=False)
            st.success("Ring positions saved.")
            st.rerun()
        leadlag_live = lead_lag_table(prepare_ring_table(edited, selected_date), selected_date, limit, warning_limit)
        if leadlag_live.empty:
            st.info("Enter valid undercut headings and ring numbers to calculate lead and lag.")
        else:
            st.dataframe(leadlag_live, width="stretch", hide_index=True)
            chart = px.bar(leadlag_live, x="Heading", y="Ring", color="Status", title="Undercut ring profile")
            st.plotly_chart(chart, width="stretch")

    # 7. Equipment Intelligence
    with tabs[6]:
        st.subheader("Equipment Intelligence")
        eq = day[day["Machine"].astype(str).str.strip().ne("")].copy()
        if eq.empty:
            st.info("No machine numbers were recorded.")
        else:
            equipment = eq.groupby("Machine", as_index=False).agg(
                Recorded_Hours=("Hours", "sum"),
                Productive_Hours=("Productive", "sum"),
                Delay_Hours=("Delay", "sum"),
                Areas=("Area", "nunique"),
            )
            equipment["Utilisation %"] = (100 * equipment["Productive_Hours"] / equipment["Recorded_Hours"]).round(1)
            equipment = equipment.sort_values(["Delay_Hours", "Utilisation %"], ascending=[False, True])
            st.plotly_chart(px.bar(equipment, x="Machine", y=["Productive_Hours", "Delay_Hours"], barmode="stack", title="Machine productive and delay hours"), width="stretch")
            st.dataframe(equipment, width="stretch", hide_index=True)

    # 8. Crew productivity and shift comparison
    with tabs[7]:
        st.subheader("Crew & Shift Comparison")
        compare = activity[(activity["Date"] >= selected_date - pd.Timedelta(days=6)) & (activity["Date"] <= selected_date)].copy()
        crew = compare.groupby(["Crew", "Shift"], as_index=False).agg(
            Productive_Hours=("Productive", "sum"), Delay_Hours=("Delay", "sum"), Areas_Worked=("Area", "nunique")
        )
        crew["Productivity %"] = (100 * crew["Productive_Hours"] / (crew["Productive_Hours"] + crew["Delay_Hours"]).replace(0, pd.NA)).fillna(0).round(1)
        st.plotly_chart(px.bar(crew, x="Crew", y="Productivity %", color="Shift", barmode="group", title="Seven-day crew and shift productivity comparison"), width="stretch")
        st.dataframe(crew.sort_values("Productivity %", ascending=False), width="stretch", hide_index=True)

    # 9. Operational Trends
    with tabs[8]:
        st.subheader("Operational Trends")
        trend = activity.groupby("Date", as_index=False).agg(Productive_Hours=("Productive", "sum"), Delay_Hours=("Delay", "sum"), Areas=("Area", "nunique"))
        trend = trend.sort_values("Date")
        st.plotly_chart(px.line(trend, x="Date", y=["Productive_Hours", "Delay_Hours"], markers=True, title="Productive and delay trend"), width="stretch")
        delay_trend = activity[activity["Delay"]].groupby(["Date", "Delay Category"], as_index=False)["Hours"].sum()
        if not delay_trend.empty:
            st.plotly_chart(px.area(delay_trend, x="Date", y="Hours", color="Delay Category", title="Delay-cause trend"), width="stretch")

    # 10. Daily Dispatch Intelligence narrative
    with tabs[9]:
        st.subheader("Daily Dispatch Intelligence")
        narrative = dispatch_intelligence_text(activity, area_summary, tracker, leadlag, selected_date)
        st.info(narrative)
        priorities = []
        if not tracker.empty:
            priorities += [f"Restore productive activity at {a}." for a in tracker.loc[tracker["Days Since Productive Work"] >= 7, "Area"].head(5)]
        delayed_areas = area_summary.sort_values("Delay Hours", ascending=False).head(3) if not area_summary.empty else pd.DataFrame()
        priorities += [f"Review the delay burden at {r['Area']} ({r['Delay Hours']:.1f} h)." for _, r in delayed_areas.iterrows() if r["Delay Hours"] > 0]
        if priorities:
            st.markdown("### Priority actions")
            for p in priorities:
                st.write("• " + p)
        st.download_button("Download daily intelligence note", narrative.encode("utf-8"), f"dispatch_intelligence_{selected_date.date()}.txt", "text/plain")

    # Bonus: Mine Health Index
    with tabs[10]:
        st.subheader("Mine Health Index")
        gauge = go.Figure(go.Indicator(
            mode="gauge+number", value=health_score,
            title={"text": health_label},
            gauge={"axis": {"range": [0, 100]}, "steps": [
                {"range": [0, 45], "color": "#fee2e2"}, {"range": [45, 65], "color": "#ffedd5"},
                {"range": [65, 80], "color": "#fef9c3"}, {"range": [80, 100], "color": "#dcfce7"},
            ]},
        ))
        gauge.update_layout(height=330)
        st.plotly_chart(gauge, width="stretch")
        components = pd.DataFrame({"Component": list(health_components), "Score": list(health_components.values())})
        st.plotly_chart(px.bar(components, x="Component", y="Score", range_y=[0, 100], title="Health-index components"), width="stretch")
        st.caption("The index is a transparent operational indicator based on productive time, active areas, seven-day continuity and undercut sequence. It supports—not replaces—professional operational judgement.")
