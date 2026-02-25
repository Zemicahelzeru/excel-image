<div class="header">
    <div class="container">
        <h1>Image and Vendor Material Extractor</h1>
    </div>
</div>

<div class="container">
    <div class="instructions">
        <h2>Instructions</h2>
        <ol>
            <li>Use <strong>Column A</strong> for images and <strong>Column B</strong> for Vendor Material #.</li>
            <li>Paste images while Column A is selected. Multiple images are placed in consecutive rows.</li>
            <li>Paste vendor codes in Column B (same row as each image).</li>
            <li>Add more rows from right-click menu or keep pasting; the grid auto-expands.</li>
            <li>Click <strong>Generate Zip</strong> to download files named <code>VENDOR_R{row}</code>.</li>
        </ol>
    </div>

    <div class="grid-labels">
        <span><strong>Column A:</strong> Images</span>
        <span><strong>Column B:</strong> Vendor Material #</span>
    </div>

    <div id="excel-grid" style="height: 460px; width: 100%;"></div>
    <button id="processButton" class="process-btn">Generate Zip</button>
    <div id="status"></div>
</div>

<link href="styles.css" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.js"></script>
<script src="script.js"></script>
