import streamlit as st
import pdfplumber
import re
import pandas as pd
from pathlib import Path
from datetime import datetime
import plotly.graph_objects as go
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Geotechnical Draw Intelligence", layout="wide")

DATA_FILE = Path("gdi_history.csv")

DEFAULT_EP_THRESHOLD = 10000
DEFAULT_ET_THRESHOLD = 10

st.title("⛏️ Geotechnical Draw Intelligence (GDI)")
st.markdown("Lift II East seismic draw intelligence from IMS daily PDF reports.")


# ---------------- FUNCTIONS ----------------

def find_value(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def clean_report_date(last_event):
    if not last_event:
        return ""
    match = re.search(r"(\d{1,2})\s+([A-Za-z]{3})", last_event)
    if not match:
        return last_event

    day = int(match.group(1))
    month_text = match.group(2).title()
    month_map = {
        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
        "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
        "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
    }

    month = month_map.get(month_text)
    if not month:
        return last_event

    year = datetime.now().year
    return f"{year}-{month:02d}-{day:02d}"


def load_history():
    if DATA_FILE.exists():
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame()


def save_history(history):
    history.to_csv(DATA_FILE, index=False)


def save_or_update_record(record):
    history = load_history()

    if not history.empty and "Report Date" in history.columns:
        if record["Report Date"] in history["Report Date"].astype(str).values:
            history.loc[history["Report Date"].astype(str) == record["Report Date"], list(record.keys())] = list(record.values())
        else:
            history = pd.concat([history, pd.DataFrame([record])], ignore_index=True)
    else:
        history = pd.DataFrame([record])

    save_history(history)


def get_status(ep, et, ep_threshold, et_threshold):
    if et is None:
        return "Waiting for tonnes"
    if ep < ep_threshold and et < et_threshold:
        return "Efficient Draw"
    if ep >= ep_threshold and et < et_threshold:
        return "Brittle but Efficient"
    if ep < ep_threshold and et >= et_threshold:
        return "Ductile but Costly"
    return "High-Risk Draw"


def cave_health_score(status):
    scores = {
        "Efficient Draw": 90,
        "Brittle but Efficient": 65,
        "Ductile but Costly": 55,
        "High-Risk Draw": 30,
        "Waiting for tonnes": 0,
    }
    return scores.get(status, 0)


def recommended_action(status):
    actions = {
        "Efficient Draw": "Maintain current draw strategy and continue routine monitoring.",
        "Brittle but Efficient": "Continue draw, but monitor seismic response closely due to elevated E/P.",
        "Ductile but Costly": "Review draw distribution and tonnes efficiency because seismic cost per tonne is elevated.",
        "High-Risk Draw": "Escalate for geotechnical review. Monitor draw, seismicity and nearby workplaces closely.",
        "Waiting for tonnes": "Enter tonnes drawn to finalise the draw compliance classification.",
    }
    return actions.get(status, "")


def trend_direction(history):
    if history.empty or len(history) < 2 or "E/P" not in history.columns:
        return "Insufficient history"

    hist = history.copy()
    hist["E/P"] = pd.to_numeric(hist["E/P"], errors="coerce")
    hist = hist.dropna(subset=["E/P"])

    if len(hist) < 2:
        return "Insufficient history"

    previous = hist["E/P"].iloc[-2]
    current = hist["E/P"].iloc[-1]

    if current > previous * 1.10:
        return "Deteriorating"
    elif current < previous * 0.90:
        return "Improving"
    else:
        return "Stable"


def plot_matrix(ep, et, history, ep_threshold, et_threshold):
    max_x = max(et_threshold * 2, et * 1.4)
    max_y = max(ep_threshold * 2, ep * 1.4)

    fig = go.Figure()

    fig.add_shape(type="rect", x0=0, x1=et_threshold, y0=0, y1=ep_threshold,
                  fillcolor="green", opacity=0.12, line_width=0)
    fig.add_shape(type="rect", x0=et_threshold, x1=max_x, y0=0, y1=ep_threshold,
                  fillcolor="orange", opacity=0.14, line_width=0)
    fig.add_shape(type="rect", x0=0, x1=et_threshold, y0=ep_threshold, y1=max_y,
                  fillcolor="gold", opacity=0.16, line_width=0)
    fig.add_shape(type="rect", x0=et_threshold, x1=max_x, y0=ep_threshold, y1=max_y,
                  fillcolor="red", opacity=0.12, line_width=0)

    fig.add_shape(type="line", x0=et_threshold, x1=et_threshold, y0=0, y1=max_y,
                  line=dict(dash="dash", width=2))
    fig.add_shape(type="line", x0=0, x1=max_x, y0=ep_threshold, y1=ep_threshold,
                  line=dict(dash="dash", width=2))

    if not history.empty and "E/P" in history.columns and "E/T" in history.columns:
        hist = history.copy()
        hist["E/P"] = pd.to_numeric(hist["E/P"], errors="coerce")
        hist["E/T"] = pd.to_numeric(hist["E/T"], errors="coerce")
        hist = hist.dropna(subset=["E/P", "E/T"])

        if not hist.empty:
            fig.add_trace(go.Scatter(
                x=hist["E/T"],
                y=hist["E/P"],
                mode="markers",
                name="Saved Records",
                text=hist["Report Date"],
                marker=dict(size=10),
                hovertemplate="Date: %{text}<br>E/T: %{x:.2f}<br>E/P: %{y:.2f}<extra></extra>",
            ))

    fig.add_trace(go.Scatter(
        x=[et],
        y=[ep],
        mode="markers+text",
        name="Current Lift II East",
        text=["Current Lift II East"],
        textposition="top center",
        marker=dict(size=22),
        hovertemplate="Current Lift II East<br>E/T: %{x:.2f}<br>E/P: %{y:.2f}<extra></extra>",
    ))

    fig.add_annotation(x=et_threshold * 0.45, y=ep_threshold * 0.5, text="Efficient Draw", showarrow=False)
    fig.add_annotation(x=et_threshold * 1.45, y=ep_threshold * 0.5, text="Ductile but Costly", showarrow=False)
    fig.add_annotation(x=et_threshold * 0.5, y=ep_threshold * 1.45, text="Brittle but Efficient", showarrow=False)
    fig.add_annotation(x=et_threshold * 1.45, y=ep_threshold * 1.45, text="High-Risk Draw", showarrow=False)

    fig.update_layout(
        title="Draw Compliance Matrix",
        xaxis_title="Energy / Tonnes (E/T)",
        yaxis_title="Energy / Potency (E/P)",
        xaxis=dict(range=[0, max_x]),
        yaxis=dict(range=[0, max_y]),
        height=650,
    )

    return fig


def build_word_report(comment, record):
    doc = Document()
    doc.add_heading("Geotechnical Draw Intelligence Report", 0)
    doc.add_heading("Lift II East Summary", level=1)

    for key, value in record.items():
        doc.add_paragraph(f"{key}: {value}")

    doc.add_heading("Management Comment", level=1)
    doc.add_paragraph(comment)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------- SIDEBAR ----------------

st.sidebar.header("Settings")
ep_threshold = st.sidebar.number_input("E/P Threshold", min_value=0.0, value=float(DEFAULT_EP_THRESHOLD))
et_threshold = st.sidebar.number_input("E/T Threshold", min_value=0.0, value=float(DEFAULT_ET_THRESHOLD))

st.sidebar.markdown("---")
st.sidebar.write("Thresholds can be adjusted as GDI calibration improves.")


# ---------------- INPUTS ----------------

uploaded_pdf = st.file_uploader("Upload IMS Daily PDF", type=["pdf"])
uploaded_excel = st.file_uploader("Optional: Upload Production Tonnes Excel", type=["xlsx", "xls", "csv"])


history = load_history()
tonnes_from_excel = None

if uploaded_excel:
    if uploaded_excel.name.endswith(".csv"):
        tonnes_df = pd.read_csv(uploaded_excel)
    else:
        tonnes_df = pd.read_excel(uploaded_excel)

    st.subheader("Production Tonnes File Preview")
    st.dataframe(tonnes_df.head(), use_container_width=True)

    tonnes_col = st.selectbox("Select Tonnes Column", tonnes_df.columns)
    tonnes_from_excel = pd.to_numeric(tonnes_df[tonnes_col], errors="coerce").sum()
    st.success(f"Tonnes from Excel: {tonnes_from_excel:,.2f}")


# ---------------- MAIN PDF PROCESSING ----------------

if uploaded_pdf:
    full_text = ""

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

    st.success("PDF uploaded and read successfully.")

    if "Lift_II_E" in full_text:
        st.success("Lift II East section found.")
    else:
        st.warning("Lift II East section not found.")

    last_event = find_value(r"Last event Date and Time:\s*(.+)", full_text)
    new_events = find_value(r"Number of new events:\s*(\d+)", full_text)
    total_energy = find_value(r"Total Energy:\s*([0-9.eE+-]+)", full_text)
    total_potency = find_value(r"Total Potency:\s*([0-9.eE+-]+)", full_text)
    current_activity = find_value(r"Current normalized activity rate:\s*([0-9.]+)", full_text)
    medium_activity = find_value(r"Medium term normalized activity rate:\s*([0-9.]+)", full_text)

    st.header("Extracted Lift II East Data")

    col1, col2, col3 = st.columns(3)

    col1.metric("New Events", new_events or "Not found")
    col1.metric("Total Energy (J)", total_energy or "Not found")

    col2.metric("Total Potency (m³)", total_potency or "Not found")
    col2.metric("Current Activity Rate", current_activity or "Not found")

    col3.metric("Medium Term Activity Rate", medium_activity or "Not found")
    col3.metric("Last Event", last_event or "Not found")

    st.header("Tonnes Input")

    input_mode = st.radio("Tonnes Source", ["Manual Entry", "Excel Upload"], horizontal=True)

    if input_mode == "Excel Upload" and tonnes_from_excel is not None:
        tonnes = tonnes_from_excel
        st.info(f"Using tonnes from Excel: {tonnes:,.2f}")
    else:
        tonnes = st.number_input("Tonnes Drawn", min_value=0.0, value=0.0, step=1.0)

    if total_energy and total_potency:
        energy = float(total_energy)
        potency = float(total_potency)
        events = int(new_events) if new_events else 0

        ep = energy / potency
        et = energy / tonnes if tonnes > 0 else None
        status = get_status(ep, et, ep_threshold, et_threshold)
        health = cave_health_score(status)
        action = recommended_action(status)

        st.header("Draw Intelligence Results")

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Energy / Potency (E/P)", f"{ep:,.2f}")
        r2.metric("Energy / Tonnes (E/T)", f"{et:,.2f}" if et is not None else "Waiting")
        r3.metric("Draw Status", status)
        r4.metric("Cave Health Score", f"{health}/100" if health else "Waiting")

        st.header("Recommended Action")
        st.info(action)

        if et is not None:
            st.header("Draw Compliance Matrix")
            st.plotly_chart(plot_matrix(ep, et, history, ep_threshold, et_threshold), use_container_width=True)

        st.header("Save / Update Daily Record")

        suggested_date = clean_report_date(last_event)
        report_date = st.text_input("Report Date", value=suggested_date)

        if not history.empty and "Report Date" in history.columns and report_date in history["Report Date"].astype(str).values:
            st.warning("A record for this date already exists. Saving will update the existing record.")

        comment = ""

        if et is None:
            comment = (
                f"Lift II East recorded {events} new seismic events, with total energy of {energy:,.0f} J "
                f"and total potency of {potency:.3f} m³. The calculated E/P ratio is {ep:,.2f}. "
                f"E/T and the final draw compliance quadrant cannot be confirmed until tonnes drawn are entered."
            )
        else:
            comment = (
                f"Lift II East recorded {events} new seismic events, with total energy of {energy:,.0f} J "
                f"and total potency of {potency:.3f} m³. The calculated E/P ratio is {ep:,.2f}, "
                f"while E/T is {et:,.2f}. Based on the draw compliance matrix, Lift II East plots within the "
                f"'{status}' quadrant, with a cave health score of {health}/100. {action}"
            )

        if st.button("Save or Update Record"):
            record = {
                "Report Date": report_date,
                "New Events": events,
                "Energy J": energy,
                "Potency m3": potency,
                "Tonnes": tonnes,
                "E/P": ep,
                "E/T": et if et is not None else "",
                "Current Activity Rate": current_activity,
                "Medium Activity Rate": medium_activity,
                "Status": status,
                "Cave Health Score": health,
                "Recommended Action": action,
            }

            save_or_update_record(record)
            st.success("Record saved or updated successfully.")

        st.header("Management Comment")
        st.text_area("Copy-ready comment", value=comment, height=170)

        if et is not None:
            record_for_report = {
                "Report Date": report_date,
                "New Events": events,
                "Energy J": energy,
                "Potency m3": potency,
                "Tonnes": tonnes,
                "E/P": round(ep, 2),
                "E/T": round(et, 2),
                "Status": status,
                "Cave Health Score": health,
                "Recommended Action": action,
            }

            word_file = build_word_report(comment, record_for_report)

            st.download_button(
                "Download Word Report",
                data=word_file,
                file_name=f"GDI_Report_{report_date}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    with st.expander("Show full extracted PDF text"):
        st.text(full_text)

else:
    st.info("Upload an IMS PDF to begin.")


# ---------------- HISTORY SECTION ----------------

st.header("GDI History")

history = load_history()

if history.empty:
    st.info("No saved records yet.")
else:
    st.dataframe(history, use_container_width=True)

    csv = history.to_csv(index=False).encode("utf-8")
    st.download_button("Download History CSV", csv, "gdi_history.csv", "text/csv")

    st.subheader("Delete Record")
    delete_date = st.selectbox("Select date to delete", history["Report Date"].astype(str).tolist())

    if st.button("Delete Selected Record"):
        history = history[history["Report Date"].astype(str) != delete_date]
        save_history(history)
        st.success(f"Deleted record for {delete_date}. Refresh the app to update the table.")

    history["E/P"] = pd.to_numeric(history["E/P"], errors="coerce")
    history["E/T"] = pd.to_numeric(history["E/T"], errors="coerce")

    st.subheader("Trend Direction")
    trend = trend_direction(history)

    if trend == "Improving":
        st.success("Trend: Improving")
    elif trend == "Deteriorating":
        st.error("Trend: Deteriorating")
    elif trend == "Stable":
        st.info("Trend: Stable")
    else:
        st.warning("Trend: Insufficient history")

    st.subheader("E/P Trend")
    st.line_chart(history.set_index("Report Date")["E/P"])

    st.subheader("E/T Trend")
    st.line_chart(history.set_index("Report Date")["E/T"])
