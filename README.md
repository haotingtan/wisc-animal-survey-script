# wisc-animal-survey-script

## Streamlit survey intake

This app starts a survey review flow by collecting:

1. A survey data workbook (`.xlsx` or `.xlsm`) or CSV file.
2. The original survey instrument as a PDF.
3. A worksheet choice when the workbook contains multiple sheets.

At this stage, the app only reads Excel worksheet names. It does not display or
analyze response rows, and it does not write uploaded files to disk.

Run locally:

```bash
python3 -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

For teacher access through Streamlit Community Cloud, deploy `streamlit_app.py`
from this repository and set the app to private, then invite the teacher by
email. Keep source workbooks, PDFs, and Streamlit secrets out of Git; the
`.gitignore` excludes common source data files under `source/`.
