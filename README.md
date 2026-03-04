# UK Tenancy Agreement Generator

Streamlit application for generating UK tenancy agreements from the supplied Word template and apartment CSV.

## Features

- Reads the provided property CSV and auto-fills the selected apartment details.
- Detects and maps the template placeholders once, then reuses each value everywhere it appears.
- Supports a dynamic number of tenants and rebuilds the party/signature sections to match.
- Extracts ID details from uploaded image files using `tesseract` when available.
- Generates a `.docx` agreement and can also create a PDF using the built-in macOS `textutil` and `cupsfilter` tools.

## Run

```bash
python3 -m pip install -r requirements.txt
streamlit run app.py
```

The app defaults to:

- Template: `/Users/diegogallardo/Desktop/LTO-Template-Formatted-v2.docx`
- Property CSV: `/Users/diegogallardo/Downloads/Apartments.csv`

You can override either path in the Streamlit sidebar.
