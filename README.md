# Excel Image Extractor (Flask)

This app extracts images from an Excel workbook and downloads them as a ZIP file.

## What it does

- Upload `.xlsx` or `.xlsm` file
- Choose a sheet
- Extract images anchored in **Column A**
- Name files using nearest value above/in **Column D**
- Download ZIP containing:
  - `images/...`
  - `summary.txt`

## Project structure

```
app.py
templates/
  index.html
static/
  styles.css
  app.js
```

## Run locally

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Start the app:

   ```bash
   python app.py
   ```

3. Open:

   ```
   http://127.0.0.1:5000
   ```

## Notes for Dataiku

- Frontend calls try these API paths in order:
  - `backend/...`
  - `/backend/...`
  - `/...`
- Backend exposes both `/...` and `/backend/...` routes for compatibility.