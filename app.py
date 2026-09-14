"""
DDL Automation — Web App
=========================
Deploy this for free on Streamlit Community Cloud (share.streamlit.io):

1. Create a free GitHub account (if you don't have one) and a new repo, e.g. "ddl-automation".
2. Upload these 3 files to it: app.py, process_ddl.py, requirements.txt
3. Go to https://share.streamlit.io, sign in with GitHub, click "New app",
   pick your repo, set main file to "app.py", click Deploy.
4. You'll get a permanent link like https://your-name-ddl-automation.streamlit.app
   Bookmark it. Open it every 15 days, upload the raw DDL file, download the result.

Your staff mapping (staff_mapping.csv) is edited right inside the app below —
no need to touch code or GitHub again once it's deployed.
"""

import streamlit as st
import pandas as pd
import io
import tempfile
import os

from process_ddl import load_raw, load_staff_map, write_sheet, WANTED_COLUMNS
import openpyxl
import csv

st.set_page_config(page_title="DDL Automation", layout="centered")
st.title("DDL Report Automation")
st.caption("Upload the raw DDL export → get a print-ready, staff-split workbook back.")

# --- Staff mapping editor -----------------------------------------------
st.subheader("1. Staff mapping (DTC Code → Staff Name)")
st.write("Edit this once. It's remembered in your browser session — for a permanent "
         "save, download it and re-upload next time, or ask to wire up persistent storage.")

if "staff_df" not in st.session_state:
    default_path = os.path.join(os.path.dirname(__file__), "staff_mapping.csv")
    if os.path.exists(default_path):
        st.session_state.staff_df = pd.read_csv(default_path)
    else:
        st.session_state.staff_df = pd.DataFrame({"dtc_code": [], "staff_name": []})

edited = st.data_editor(
    st.session_state.staff_df,
    num_rows="dynamic",
    use_container_width=True,
    key="staff_editor",
)

col1, col2 = st.columns(2)
with col1:
    csv_bytes = edited.to_csv(index=False).encode("utf-8")
    st.download_button("Download this mapping as CSV", csv_bytes, "staff_mapping.csv", "text/csv")
with col2:
    uploaded_mapping = st.file_uploader("...or load a saved mapping CSV", type=["csv"], key="mapping_upload")
    if uploaded_mapping is not None:
        edited = pd.read_csv(uploaded_mapping)
        st.session_state.staff_df = edited
        st.success("Mapping loaded.")

# --- Raw DDL upload -------------------------------------------------------
st.subheader("2. Upload the raw DDL file")
raw_file = st.file_uploader("DDL export (.xls or .xlsx)", type=["xls", "xlsx"], key="raw_upload")

if raw_file is not None and st.button("Process file", type="primary"):
    with st.spinner("Processing..."):
        with tempfile.TemporaryDirectory() as tmp:
            raw_path = os.path.join(tmp, raw_file.name)
            with open(raw_path, "wb") as f:
                f.write(raw_file.getbuffer())

            mapping_path = os.path.join(tmp, "staff_mapping.csv")
            edited.to_csv(mapping_path, index=False)

            banner_rows, headers, data_rows = load_raw(raw_path)
            staff_map = load_staff_map(mapping_path)
            dtc_idx = headers.index("DTC Code")

            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            write_sheet(wb, "Report", banner_rows, headers, data_rows)

            by_staff = {}
            for row in data_rows:
                dtc = str(row[dtc_idx]).strip()
                staff = staff_map.get(dtc)
                if staff:
                    by_staff.setdefault(staff, []).append(row)
            for staff, rows in by_staff.items():
                write_sheet(wb, staff, banner_rows, headers, rows)

            out_buf = io.BytesIO()
            wb.save(out_buf)
            out_buf.seek(0)

            st.success(f"Done — {len(data_rows)} total consumers, split into {len(by_staff)} staff sheets.")
            for staff, rows in by_staff.items():
                st.write(f"- **{staff}**: {len(rows)} consumers")

            st.download_button(
                "Download formatted workbook",
                out_buf,
                file_name="DDL_Report_Formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
