# Excel Image Extractor

This project contains an Excel image extraction webapp.

## Dataiku-first setup (recommended)

If your target is Dataiku only, use the ready-to-paste bundle in:

```
dataiku_webapp/
```

Detailed steps:

- `dataiku_webapp/README_DATAIKU.md`

## What it does

- Upload `.xlsx` or `.xlsm` file
- Choose a sheet
- Extract images anchored in **Column A**
- Name files using nearest value above/in **Column D**
- Download ZIP containing:
  - `images/...`
  - `summary.txt`

## Included implementations

- **Dataiku webapp bundle** (copy/paste into DSS tabs)
- **Standalone Flask app** (optional local run)

## Local Flask structure (optional)

```
app.py
templates/
  index.html
static/
  styles.css
  app.js
dataiku_webapp/
  backend.py
  index.html
  styles.css
  script.js
  README_DATAIKU.md
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