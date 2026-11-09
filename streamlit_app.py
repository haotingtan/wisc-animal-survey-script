"""Minimal survey intake flow for the Wisconsin animal survey project."""

from __future__ import annotations

from io import BytesIO

import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="Survey Intake", page_icon="📋", layout="centered")
st.title("Survey Intake")
st.write(
    "Upload the survey data file and the original survey PDF. "
    "For Excel workbooks with multiple sheets, choose the sheet to use."
)
st.info(
    "This first step only collects the files and sheet choice. It does not display, "
    "analyze, or save response data."
)

st.subheader("1. Survey data")
survey_file = st.file_uploader(
    "Upload survey data (Excel or CSV)",
    type=["xlsx", "xlsm", "csv"],
    key="survey_data_upload",
    help="Supported formats: .xlsx, .xlsm, and .csv.",
)

sheet_names: list[str] = []
selected_sheet: str | None = None
dataset_ready = False

if survey_file is not None:
    file_bytes = survey_file.getvalue()
    if not file_bytes:
        st.error("The survey data file is empty.")
    elif survey_file.name.lower().endswith(".csv"):
        dataset_ready = True
        sheet_names = ["CSV data"]
        selected_sheet = "CSV data"
        st.text(f"Uploaded: {survey_file.name} · {len(file_bytes):,} bytes")
        st.caption("CSV files contain one table, so no sheet selection is needed.")
    else:
        try:
            workbook = load_workbook(BytesIO(file_bytes), read_only=True, data_only=False)
            try:
                sheet_names = workbook.sheetnames
            finally:
                workbook.close()
            if not sheet_names:
                st.error("No worksheets were found in this workbook.")
            else:
                dataset_ready = True
                st.text(f"Uploaded: {survey_file.name} · {len(file_bytes):,} bytes")
                if len(sheet_names) == 1:
                    selected_sheet = sheet_names[0]
                    st.text(f"One worksheet found: {selected_sheet}")
                else:
                    selected_sheet = st.selectbox(
                        "Which worksheet should be used?",
                        sheet_names,
                        key="survey_sheet_choice",
                    )
        except Exception as exc:  # openpyxl reports several format-specific errors
            st.error(f"Could not read the Excel workbook: {exc}")

st.subheader("2. Original survey")
survey_pdf = st.file_uploader(
    "Upload the original survey instrument (PDF)",
    type=["pdf"],
    key="survey_pdf_upload",
)
pdf_ready = False
if survey_pdf is not None:
    pdf_bytes = survey_pdf.getvalue()
    if not pdf_bytes.startswith(b"%PDF-"):
        st.error("The uploaded file does not appear to be a valid PDF.")
    else:
        pdf_ready = True
        st.text(f"Uploaded: {survey_pdf.name} · {len(pdf_bytes):,} bytes")

st.divider()
if dataset_ready and pdf_ready:
    st.success("Both files are ready for the next step.")
    st.text(f"Selected data source: {survey_file.name}")
    st.text(f"Selected worksheet: {selected_sheet}")
    st.text(f"Survey instrument: {survey_pdf.name}")
else:
    st.caption("Upload both files to complete this step.")
