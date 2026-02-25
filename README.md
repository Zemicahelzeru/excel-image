<!-- index.html -->
<div class="container">
    <h2>Image and Vendor Code Extractor</h2>
    <div id="excel-grid" style="height: 400px; width: 100%;"></div>
    <button onclick="processData()" class="process-btn">Generate Zip</button>
    <div id="status"></div>
</div>

<!-- Include external libraries -->
<link href="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/handsontable/dist/handsontable.full.min.js"></script>

<!-- Include our CSS and JS -->
<link href="static/styles.css" rel="stylesheet">
<script src="static/main.js"></script>

/* styles.css */
.container {
    padding: 20px;
    max-width: 1200px;
    margin: 0 auto;
}

.process-btn {
    margin-top: 20px;
    padding: 10px 20px;
    background: #041E42;
    color: white;
    border: none;
    cursor: pointer;
}

.image-cell {
    height: 50px;
    width: 50px;
    background-size: contain;
    background-repeat: no-repeat;
    background-position: center;
}

#status {
    margin-top: 10px;
    padding: 10px;
    border-radius: 4px;
}

.status-error {
    background-color: #ffe6e6;
    color: #ff0000;
}

.status-success {
    background-color: #e6ffe6;
    color: #008000;
}


// main.js
document.addEventListener('DOMContentLoaded', function() {
    initializeGrid();
});

function initializeGrid() {
    const container = document.getElementById('excel-grid');
    const hot = new Handsontable(container, {
        data: [
            ['Image', 'Vendor Code'],
            ['', ''],
            ['', ''],
            ['', '']
        ],
        colHeaders: true,
        rowHeaders: true,
        height: 'auto',
        licenseKey: 'non-commercial-and-evaluation',
        contextMenu: true,
        columns: [
            {
                type: 'text',
                renderer: imageRenderer
            },
            { type: 'text' }
        ],
        afterPaste: handleAfterPaste
    });

    // Store hot instance globally
    window.hotInstance = hot;
    
    // Add paste event listener
    document.addEventListener('paste', handleImagePaste);
}

function imageRenderer(instance, td, row, col, prop, value, cellProperties) {
    if (value && value.startsWith('data:image')) {
        td.innerHTML = `<div class="image-cell" style="background-image: url(${value})"></div>`;
    } else {
        Handsontable.renderers.TextRenderer.apply(this, arguments);
    }
}

function handleAfterPaste(data, coords) {
    if (coords[0].startCol === 0) {
        processImagePaste(data, coords[0].startRow);
    }
}

function processImagePaste(clipboardData, startRow) {
    const dataArray = typeof clipboardData === 'string' ? [clipboardData] : clipboardData;
    
    dataArray.forEach((item, index) => {
        if (item && item[0] && item[0].startsWith('data:image')) {
            window.hotInstance.setDataAtCell(startRow + index, 0, item[0]);
        }
    });
}

function handleImagePaste(e) {
    const activeCell = window.hotInstance.getSelectedLast();
    if (!activeCell || activeCell[1] !== 0) return;
    
    const items = e.clipboardData.items;
    let imageCount = 0;
    
    for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
            const blob = items[i].getAsFile();
            const reader = new FileReader();
            reader.onload = function(event) {
                window.hotInstance.setDataAtCell(activeCell[0] + imageCount, 0, event.target.result);
                imageCount++;
            };
            reader.readAsDataURL(blob);
        }
    }
}

function processData() {
    const data = window.hotInstance.getData();
    const processedData = data.slice(1).filter(row => row[0] || row[1]);
    
    // Call backend API
    getDataikuAPI()
        .call('/process_data', {
            'data': JSON.stringify(processedData)
        })
        .then(handleSuccess)
        .catch(handleError);
}

function handleSuccess(response) {
    document.getElementById('status').className = 'status-success';
    document.getElementById('status').textContent = 'Processing successful!';
    window.location.href = response.downloadUrl;
}

function handleError(error) {
    document.getElementById('status').className = 'status-error';
    document.getElementById('status').textContent = 'Error: ' + error;
}


# app.py
import dataiku
from flask import Flask, request, send_file
import json
import base64
import io
import zipfile

app = dataiku.get_custom_variables()["webapp"]

@app.route('/')
def serve_page():
    return app.get_template_resource('index.html')

@app.route('/process_data', methods=['POST'])
def process_data():
    try:
        data = json.loads(request.form['data'])
        memory_file = io.BytesIO()
        
        with zipfile.ZipFile(memory_file, 'w') as zf:
            for index, row in enumerate(data):
                if row[0] and row[1]:  # If both image and vendor code exist
                    image_data = row[0].split(',')[1]
                    image_bytes = base64.b64decode(image_data)
                    filename = f"{row[1]}.jpg"
                    zf.writestr(filename, image_bytes)
        
        memory_file.seek(0)
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name='vendor_images.zip'
        )
    
    except Exception as e:
        return json.dumps({'error': str(e)}), 400
