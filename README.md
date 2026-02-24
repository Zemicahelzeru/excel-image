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
                <li>Preview mapping to validate naming accuracy</li>
                <li>Process to download ZIP with deterministic row-based names</li>
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
            <div class="grid-actions">
                <button type="button" class="button" id="previewMappingButton" disabled>Preview Mapping</button>
                <button class="button" id="processButton" disabled>Process Selected Sheet</button>
            </div>
            <div id="mappingPreview" style="display:none;"></div>
        </div>

        <div class="spinner" id="spinner" aria-hidden="true"></div>
        <div class="progress-bar" id="progressBar" aria-hidden="true">
            <div class="progress" id="progress"></div>
        </div>
        <div id="status" class="status" role="status" aria-live="polite"></div>
    </div>
</div>
