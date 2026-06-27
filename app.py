import streamlit as st
import pdfplumber
import re
import pandas as pd
from pathlib import Path
import plotly.graph_objects as go

st.set_page_config(page_title="Geotechnical Draw Intelligence", layout="wide")

DATA_FILE = Path("gdi_history.csv")

EP_THRESHOLD = 10000
ET_THRESHOLD = 10

st.title("⛏️ Geotechnical Draw Intelligence (GDI)")
st.markdown("Lift II East seismic draw intelligence from IMS daily PDF reports.")

uploaded_pdf = st.file_uploader("Upload IMS Daily PDF", type=["pdf"])


def find_value(pattern, text):
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


def load_history():
    if DATA_FILE.exists():
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame()


def save_record(record):
    history = load_history()
    history = pd.concat([history, pd.DataFrame([record])], ignore_index=True)
    history.to_csv(DATA_FILE, index=False)


def get_status(ep, et):
    if et is None:
        return "Waiting for tonnes"
    if ep < EP_THRESHOLD and et < ET_THRESHOLD:
        return "Efficient Draw"
    elif ep >= EP_THRESHOLD and et < ET_THRESHOLD:
        return "Brittle but Efficient"
    elif ep < EP_THRESHOLD and et >= ET_THRESHOLD:
        return "Ductile but Costly"
    else:
        return "High-Risk Draw"


def plot_matrix(ep, et, history=None):
    max_x = max(ET_THRESHOLD * 2, et * 1.4)
    max_y = max(EP_THRESHOLD * 2, ep * 1.4)

    fig = go.Figure()

    fig.add_shape(
        type="line",
        x0=ET_THRESHOLD,
        x1=ET_THRESHOLD,
        y0=0,
        y1=max_y,
        line=dict(dash="dash", width=2),
    )

    fig.add_shape(
        type="line",
        x0=0,
        x1=max_x,
        y0=EP_THRESHOLD,
        y1=EP_THRESHOLD,
        line=dict(dash="dash", width=2),
    )

    if history is not None and not history.empty and "E/P" in history.columns and "E/T" in history.columns:
        hist = history.copy()
        hist["E/P"] = pd.to_numeric(hist["E/P"], errors="coerce")
        hist["E/T"] = pd.to_numeric(hist["E/T"], errors="coerce")
        hist = hist.dropna(subset=["E/P", "E/T"])

        if not hist.empty:
            fig.add_trace(
                go.Scatter(
                    x=hist["E/T"],
                    y=hist["E/P"],
                    mode="markers",
                    name="Saved History",
                    marker=dict(size=10),
                    text=hist["Report Date"] if "Report Date" in hist.columns else None,
                    hovertemplate=
                    "Saved Record<br>"
                    "Date: %{text}<br>"
                    "E/T: %{x:.2f}<br>"
                    "E/P: %{y:.2f}<extra></extra>",
                )
            )

    fig.add_trace(
        go.Scatter(
            x=[et],
            y=[ep],
            mode="markers+text",
            name="Current Lift II East",
            text=["Current Lift II East"],
            textposition="top center",
            marker=dict(size=18),
            hovertemplate=
            "Current Lift II East<br>"
            "E/T: %{x:.2f}<br>"
            "E/P: %{y:.2f}<extra></extra>",
        )
    )

    fig.add_annotation(
        x=ET_THRESHOLD * 0.35,
        y=EP_THRESHOLD * 0.5,
        text="Efficient Draw",
        showarrow=False,
    )

    fig.add_annotation(
        x=ET_THRESHOLD * 1.5,
        y=EP_THRESHOLD * 0.5,
        text="Ductile but Costly",
        showarrow=False,
    )

    fig.add_annotation(
        x=ET_THRESHOLD * 0.5,
        y=EP_THRESHOLD * 1.5,
        text="Brittle but Efficient",
        showarrow=False,
    )

    fig.add_annotation(
        x=ET_THRESHOLD * 1.5,
        y=EP_THRESHOLD * 1.5,
        text="High-Risk Draw",
        showarrow=False,
    )

    fig.update_layout(
        title="Draw Compliance Matrix",
        xaxis_title="Energy / Tonnes (E/T)",
        yaxis_title="Energy / Potency (E/P)",
        xaxis=dict(range=[0, max_x]),
        yaxis=dict(range=[0, max_y]),
        height=650,
    )

    return fig


history = load_history()

if uploaded_pdf:
    full_text = ""

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

    st.success("PDF uploaded and read successfully.")

    last_event = find_value(r"Last event Date and Time:\s*(.+)", full_text)
    new_events = find_value(r"Number of new events:\s*(\d+)", full_text)
    total_energy = find_value(r"Total Energy:\s*([0-9.eE+-]+)", full_text)
    total_potency = find_value(r"Total Potency:\s*([0-9.eE+-]+)", full_text)
    current_activity = find_value(r"Current normalized activity rate:\s*([0-9.]+)", full_text)
    medium_activity = find_value(r"Medium term normalized activity rate:\s*([0-9.]+)", full_text)

    if "Lift_II_E" in full_text:
        st.success("Lift II East section found.")
    else:
        st.warning("Lift II East section not found.")

    st.header("Extracted Lift II East Data")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("New Events", new_events or "Not found")
        st.metric("Total Energy (J)", total_energy or "Not found")

    with col2:
        st.metric("Total Potency (m³)", total_potency or "Not found")
        st.metric("Current Activity Rate", current_activity or "Not found")

    with col3:
        st.metric("Medium Term Activity Rate", medium_activity or "Not found")
        st.metric("Last Event", last_event or "Not found")

    st.header("Manual Tonnes Input")

    tonnes = st.number_input("Tonnes Drawn", min_value=0.0, value=0.0)

    if total_energy and total_potency:
        energy = float(total_energy)
        potency = float(total_potency)
        events = int(new_events) if new_events else 0

        ep = energy / potency
        et = energy / tonnes if tonnes > 0 else None
        status = get_status(ep, et)

        st.header("Draw Intelligence Results")

        c1, c2 = st.columns(2)

        with c1:
            st.metric("Energy / Potency (E/P)", f"{ep:,.2f}")

        with c2:
            if et is not None:
                st.metric("Energy / Tonnes (E/T)", f"{et:,.2f}")
            else:
                st.metric("Energy / Tonnes (E/T)", "Waiting for tonnes")

        st.header("Draw Status")

        if status == "Waiting for tonnes":
            st.info("Enter tonnes drawn to calculate E/T and matrix quadrant.")
        elif status == "Efficient Draw":
            st.success(status)
        elif status in ["Brittle but Efficient", "Ductile but Costly"]:
            st.warning(status)
        else:
            st.error(status)

        if et is not None:
            st.header("Draw Compliance Matrix")
            fig = plot_matrix(ep, et, history)
            st.plotly_chart(fig, use_container_width=True)

        st.header("Save Daily Record")

        report_date = st.text_input("Report Date", value=last_event or "")

        if st.button("Save Record"):
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
            }

            save_record(record)
            st.success("Record saved to history.")

        st.header("Auto Interpretation")

        if et is None:
            st.write(
                f"Lift II East recorded {events} new events, with total energy of {energy:,.0f} J "
                f"and total potency of {potency:.3f} m³. The calculated E/P is {ep:,.2f}. "
                f"E/T and matrix quadrant cannot be finalised until tonnes drawn are entered."
            )
        else:
            st.write(
                f"Lift II East recorded {events} new events, with total energy of {energy:,.0f} J "
                f"and total potency of {potency:.3f} m³. The calculated E/P is {ep:,.2f}, "
                f"while E/T is {et:,.2f}. Based on the draw compliance matrix, "
                f"Lift II East plots within the '{status}' quadrant."
            )

    with st.expander("Show full extracted PDF text"):
        st.text(full_text)

else:
    st.info("Upload an IMS PDF to begin.")


st.header("GDI History")

history = load_history()

if history.empty:
    st.info("No saved records yet.")
else:
    st.dataframe(history, use_container_width=True)

    history["E/P"] = pd.to_numeric(history["E/P"], errors="coerce")
    history["E/T"] = pd.to_numeric(history["E/T"], errors="coerce")

    st.subheader("E/P Trend")
    st.line_chart(history.set_index("Report Date")["E/P"])

    st.subheader("E/T Trend")
    st.line_chart(history.set_index("Report Date")["E/T"])