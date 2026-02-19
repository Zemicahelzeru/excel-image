# Dataiku Webapp Setup (Copy/Paste Ready)

Use these files directly in a **Dataiku Standard Webapp** (with Python backend):

- `index.html` -> HTML tab
- `styles.css` -> CSS tab
- `script.js` -> JavaScript tab
- `backend.py` -> Python backend tab

## 1) Create the webapp

1. In Dataiku, create a **Webapp**.
2. Choose **Standard** (Python backend).

## 2) Add Python dependency

The backend needs:

- `openpyxl`

Add it to the webapp/project code environment (or admin-installed env), then restart the backend.

## 3) Paste code into tabs

Copy each file content into the matching Dataiku tab:

- HTML: `index.html`
- CSS: `styles.css`
- JavaScript: `script.js`
- Backend: `backend.py`

## 4) Run and test

1. Save webapp.
2. Start/restart backend.
3. Upload an `.xlsx` or `.xlsm`.
4. Select sheet and process.

## Troubleshooting

### `Error: Failed to fetch` on Process

This is usually a backend/runtime issue in Dataiku, not a frontend validation issue.

Check:

1. **Backend state** in the webapp editor (must be running).
2. **Backend logs** for Python exceptions or timeout.
3. `openpyxl` installed in the webapp code environment.
4. Restart webapp backend after any dependency/code change.

If you see `Webapp backend start timeout`, fix the backend environment first (packages / startup errors), then restart.

## Behavior

- Extracts only images anchored in **Column A**
- Uses nearest value above/in **Column D** as image filename
- Downloads a ZIP containing:
  - `images/...`
  - `summary.txt`
