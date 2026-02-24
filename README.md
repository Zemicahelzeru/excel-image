<div class="header">
    <div class="container">
        <h1>Excel Image Extractor</h1>
    </div>
</div>

<div class="container">
    <div class="card">
        <div class="instructions">
            <h2>Instructions</h2>
            <ol>
                <li>Upload your Excel file containing images</li>
                <li>Select the sheet containing your images</li>
                <li>Paste your Excel Column A into the grid</li>
                <li>Each row uses its own image; names keep row suffix (e.g., <code>ABC120_R12</code>)</li>
            </ol>
        </div>
    </div>

    <div class="card">
        <div class="upload-form">
            <form id="uploadForm">
                <div class="file-input">
                    <label for="excelFile" class="sr-only">Excel file</label>
                    <input type="file" accept=".xlsx,.xlsm" required id="excelFile">
                </div>
                <button type="submit" class="button" id="uploadButton">Upload Excel File</button>
            </form>
        </div>

        <div class="sheet-selection" id="sheetSelection">
            <h3>Select Sheet</h3>
            <div class="sheet-list" id="sheetList"></div>
        </div>

        <div class="paste-image-column-section" id="pasteImageColumnSection" style="display:none;">
            <h3>Paste Column A (Image Rows)</h3>
            <p>Click a cell and paste Excel Column A values. Non-empty rows are treated as image rows.</p>
            <label for="startRowInput">First Excel row number</label>
            <input id="startRowInput" type="number" min="1" value="1" />
            <div class="excel-grid-container">
                <table class="excel-grid">
                    <thead>
                        <tr>
                            <th class="row-header">Row</th>
                            <th>Column A</th>
                        </tr>
                    </thead>
                    <tbody id="gridBody"></tbody>
                </table>
            </div>
            <div class="grid-actions">
                <button type="button" class="button" id="validateGridButton">Validate Grid</button>
                <button type="button" class="button btn-secondary" id="clearGridButton">Clear Grid</button>
            </div>
            <div id="gridStatus" class="status" role="status" aria-live="polite"></div>
            <button class="button" id="processButton" disabled>Process Selected Sheet</button>
        </div>

        <div class="spinner" id="spinner" aria-hidden="true"></div>
        <div class="progress-bar" id="progressBar" aria-hidden="true">
            <div class="progress" id="progress"></div>
        </div>
        <div id="status" class="status" role="status" aria-live="polite"></div>
    </div>
</div>
