<div class="header">
    <div class="container">
        <h1>Image and Vendor Material Extractor</h1>
    </div>
</div>

<div class="container">
    <div class="instructions">
        <h2>Instructions</h2>
        <ol>
            <li>Click any cell in <strong>Column A (Images)</strong>.</li>
            <li>Paste images from clipboard. If multiple images are available, they fill rows in order.</li>
            <li>Paste vendor material codes into <strong>Column B (Vendor Material #)</strong>.</li>
            <li>Click <strong>Generate Zip</strong> to download files named by vendor code and row.</li>
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

<link href="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.js"></script>
<script src="script.js"></script>
