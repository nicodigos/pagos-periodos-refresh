from datetime import date
from uuid import uuid4

import streamlit as st

from utils.ms_graph_excel import finish_redirect_flow, get_redirect_login_url, get_token_silent
from utils.pagos_periodos_sync import (
    pagos_periodos_workbook_path,
    sync_pagos_periodos_lookup_tables,
)
from utils.work98_generator import (
    generate_work98_schedule_validation_export,
    generate_work98_workbook,
)

st.set_page_config(page_title="Pagos Periodos Refresh", layout="centered")


def render_microsoft_login() -> str | None:
    qp = st.query_params
    auth_code = qp.get("code")
    auth_state = qp.get("state")

    if auth_code:
        expected_state = st.session_state.get("oauth_state")
        try:
            if expected_state and auth_state != expected_state:
                raise RuntimeError("Invalid OAuth state. Please try connecting again.")
            token = finish_redirect_flow(str(auth_code))
            st.session_state["graph_token"] = token
            st.session_state.pop("oauth_state", None)
            st.query_params.clear()
            st.success("Connected to Microsoft.")
            st.rerun()
        except Exception as exc:
            st.session_state.pop("oauth_state", None)
            st.query_params.clear()
            st.error(f"Login failed: {exc}")

    token = get_token_silent()
    if token:
        st.session_state["graph_token"] = token
        return token

    st.warning("Microsoft connection required.")
    if not st.session_state.get("oauth_state"):
        st.session_state["oauth_state"] = str(uuid4())

    try:
        login_url = get_redirect_login_url(st.session_state["oauth_state"])
        st.link_button("Connect to Microsoft", login_url, type="primary")
    except Exception as exc:
        st.error(f"Could not build login URL: {exc}")
    return None


st.title("Pagos Periodos Refresh")
st.caption("Recarga las listas de contractors y buildings desde Microsoft hacia las tablas usadas por los dropdowns del admin.")
st.code(pagos_periodos_workbook_path(), language="text")

token = render_microsoft_login()
if not token:
    st.stop()

refresh_clicked = st.button("Refresh Pagos Periodos ID", type="primary", use_container_width=True)

if refresh_clicked:
    try:
        with st.spinner("Refreshing Pagos Periodos lookup tables..."):
            result = sync_pagos_periodos_lookup_tables(token)
        st.success(
            "Pagos Periodos ID updated. "
            f"Contractors: {result['contractors']}. Buildings: {result['buildings']}."
        )
        st.info("Refresh the admin page where you were working and open the dropdown again.")
    except Exception as exc:
        st.error(f"Refresh failed: {exc}")

st.divider()
st.subheader("Work Export")
st.caption("Genera un Excel con hojas Work1..WorkN a partir de la vista vw_jobs_with_plan_building_and_pagos_periodos.")

def normalized_export_date(value: object) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, tuple):
        return normalized_export_date(value[0] if value else None)
    if isinstance(value, date):
        return value
    return None


filter_start_col, filter_end_col = st.columns(2)
with filter_start_col:
    filter_start_date = st.date_input("Filter start date", value=None, format="YYYY-MM-DD")
with filter_end_col:
    filter_end_date = st.date_input("Filter end date", value=None, format="YYYY-MM-DD")
filter_start_date = normalized_export_date(filter_start_date)
filter_end_date = normalized_export_date(filter_end_date)

invalid_date_range = bool(
    filter_start_date and filter_end_date and filter_start_date > filter_end_date
)
if invalid_date_range:
    st.error("Filter start date must be on or before filter end date.")

workbook_col, schedule_col = st.columns(2)
with workbook_col:
    if st.button("Generate Work Workbook", use_container_width=True, disabled=invalid_date_range):
        try:
            with st.spinner("Generating Work workbook..."):
                filename, workbook_bytes = generate_work98_workbook(token, start_date=filter_start_date, end_date=filter_end_date)
            st.session_state["work98_filename"] = filename
            st.session_state["work98_bytes"] = workbook_bytes
            st.success("Work workbook generated.")
        except Exception as exc:
            st.error(f"Work export failed: {exc}")

    workbook_bytes = st.session_state.get("work98_bytes")
    workbook_filename = st.session_state.get("work98_filename", "work_export.xlsx")
    if workbook_bytes:
        st.download_button(
            "Download Work Workbook",
            data=workbook_bytes,
            file_name=workbook_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with schedule_col:
    if st.button("Generate Schedule Validation View", use_container_width=True):
        try:
            with st.spinner("Generating schedule validation view..."):
                schedule_filename, schedule_bytes = generate_work98_schedule_validation_export()
            st.session_state["work98_schedule_filename"] = schedule_filename
            st.session_state["work98_schedule_bytes"] = schedule_bytes
            st.success("Schedule validation view generated.")
        except Exception as exc:
            st.error(f"Schedule validation export failed: {exc}")

    schedule_bytes = st.session_state.get("work98_schedule_bytes")
    schedule_filename = st.session_state.get("work98_schedule_filename", "work_export_schedule_validation.csv")
    if schedule_bytes:
        st.download_button(
            "Download Schedule Validation View",
            data=schedule_bytes,
            file_name=schedule_filename,
            mime="text/csv",
            use_container_width=True,
        )
