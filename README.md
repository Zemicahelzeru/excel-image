(function () {
  "use strict";
  // Prevent double initialization
  if (window.__PASTE_GRID_APP_READY__) return;
  window.__PASTE_GRID_APP_READY__ = true;

  // --- Utilities ---
  function apiUrl(path) {
    const normalized = path.startsWith("/") ? path : "/" + path;
    return (typeof getWebAppBackendUrl === "function") ? getWebAppBackendUrl(normalized) : "backend" + normalized;
  }

  function showStatus(message, type) {
    const status = document.getElementById("status");
    if (!status) return;
    status.className = "status-" + (type || "info");
    status.textContent = message;
  }

  function escapeHtml(text) {
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }

  // --- Renderers ---
  function imageRenderer(instance, td, row, col, prop, value) {
    if (value && typeof value === "string" && value.indexOf("data:image") === 0) {
      td.innerHTML = '<div class="image-cell" style="background-image:url(' + escapeHtml(value) + '); width:50px; height:50px; background-size:contain; background-repeat:no-repeat;"></div>';
      return td;
    }
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    return td;
  }

  // --- Image Processing ---
  async function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  function isProbablyRealImage(dataUrl) {
    return dataUrl && dataUrl.indexOf("data:image/") === 0 && dataUrl.length > 1000;
  }

  /**
   * FIX: This function ensures that for every image in the list, 
   * it moves DOWN one row in the grid.
   */
  function setImagesSequentially(hot, startRow, dataUrls) {
    const changes = [];
    dataUrls.forEach((url, idx) => {
      // Column 0 is "A - Images"
      changes.push([startRow + idx, 0, url]);
    });
    
    if (changes.length > 0) {
      hot.setDataAtCell(changes, "image-paste");
    }
  }

  async function handleClipboardPaste(e, hot) {
    const selected = hot.getSelectedLast();
    if (!selected) return;

    const startRow = selected[0];
    const startCol = selected[1];

    // Only run this custom logic if pasting into Column A
    if (startCol !== 0) return;

    const items = Array.from(e.clipboardData.items || []);
    const collected = [];

    // Capture Image files directly (best for copied files/screenshots)
    for (const item of items) {
      if (item.type.indexOf("image/") === 0) {
        const file = item.getAsFile();
        if (file) {
          const dataUrl = await blobToDataUrl(file);
          if (isProbablyRealImage(dataUrl)) collected.push(dataUrl);
        }
      }
    }

    if (collected.length > 0) {
      // STOP the browser from doing a normal text-paste
      e.preventDefault();
      e.stopImmediatePropagation();
      
      setImagesSequentially(hot, startRow, collected);
      showStatus("Pasted " + collected.length + " image(s) across rows.", "success");
    }
  }

  // --- Backend Integration ---
  function processData() {
    const hot = window.hotInstance;
    if (!hot) return;
    
    const rawData = hot.getData();
    const rows = [];
    rawData.forEach((row, idx) => {
      if (row[0] && row[1]) { // Both Image and Vendor Code must exist
        rows.push({
          row_number: idx + 1,
          image_data_url: row[0],
          vendor_code: row[1].toString().trim()
        });
      }
    });

    if (!rows.length) {
      showStatus("Error: Missing image or vendor code in rows.", "error");
      return;
    }

    const formData = new FormData();
    formData.append("data", JSON.stringify(rows));
    showStatus("Sending to backend...", "info");

    fetch(apiUrl("/process_data"), { method: "POST", body: formData })
      .then(res => res.ok ? res.blob() : Promise.reject("Server Error"))
      .then(blob => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "vendor_images.zip";
        a.click();
        showStatus("ZIP downloaded!", "success");
      })
      .catch(err => showStatus(err, "error"));
  }

  // --- Initialization ---
  function initializeGrid() {
    const container = document.getElementById("excel-grid");
    if (!container) return;

    const hot = new Handsontable(container, {
      data: Array.from({ length: 100 }, () => ["", ""]),
      colHeaders: ["A - Images", "B - Vendor Material #"],
      rowHeaders: true,
      height: 500,
      stretchH: "all",
      rowHeights: 60, // Give space for images
      columns: [{ renderer: imageRenderer }, { type: "text" }],
      licenseKey: "non-commercial-and-evaluation"
    });

    window.hotInstance = hot;

    // Listen for paste specifically on the container to intercept it
    container.addEventListener("paste", (e) => {
      handleClipboardPaste(e, hot);
    }, true); 

    const btn = document.getElementById("processButton");
    if (btn) btn.addEventListener("click", processData);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initializeGrid);
  } else {
    initializeGrid();
  }
})();
