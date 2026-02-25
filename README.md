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
