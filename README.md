(function () {
  "use strict";
  if (window.__PASTE_GRID_APP_READY__) return;
  window.__PASTE_GRID_APP_READY__ = true;

  function apiUrl(path) {
    const normalized = path.startsWith("/") ? path : "/" + path;
    if (typeof getWebAppBackendUrl === "function") {
      return getWebAppBackendUrl(normalized);
    }
    return "backend" + normalized;
  }

  function showStatus(message, type) {
    const status = document.getElementById("status");
    if (!status) return;
    const kind = type || "info";
    status.className = "status-" + kind;
    status.textContent = message;
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function imageRenderer(instance, td, row, col, prop, value) {
    if (value && typeof value === "string" && value.indexOf("data:image") === 0) {
      td.innerHTML = '<div class="image-cell" style="background-image:url(' + escapeHtml(value) + ')"></div>';
      return td;
    }
    Handsontable.renderers.TextRenderer.apply(this, arguments);
    return td;
  }

  function parseImagesFromHtml(html) {
    if (!html) return [];
    const matches = html.match(/src\s*=\s*["'](data:image\/[^"']+)["']/gi) || [];
    return matches
      .map(function (m) {
        const mm = m.match(/["'](data:image\/[^"']+)["']/i);
        return mm ? mm[1] : null;
      })
      .filter(Boolean);
  }

  function setImagesSequentially(hot, startRow, dataUrls) {
    dataUrls.forEach(function (url, idx) {
      hot.setDataAtCell(startRow + idx, 0, url, "image-paste");
    });
  }

  function handleClipboardPaste(e, hot) {
    const selected = hot.getSelectedLast();
    if (!selected) return;
    const startRow = selected[0];
    const startCol = selected[1];
    if (startCol !== 0) return;

    const html = e.clipboardData ? e.clipboardData.getData("text/html") : "";
    const htmlImages = parseImagesFromHtml(html);
    const items = Array.from((e.clipboardData && e.clipboardData.items) || []);
    const imageItems = items.filter(function (it) {
      return it.type && it.type.indexOf("image/") === 0;
    });

    if (!htmlImages.length && !imageItems.length) return;
    e.preventDefault();

    if (htmlImages.length) {
      setImagesSequentially(hot, startRow, htmlImages);
      showStatus("Pasted " + htmlImages.length + " image(s).", "success");
      return;
    }

    const dataUrls = new Array(imageItems.length);
    let completed = 0;
    imageItems.forEach(function (item, idx) {
      const blob = item.getAsFile();
      if (!blob) {
        completed += 1;
        return;
      }
      const reader = new FileReader();
      reader.onload = function (event) {
        dataUrls[idx] = event.target.result;
        completed += 1;
        if (completed === imageItems.length) {
          const ordered = dataUrls.filter(Boolean);
          setImagesSequentially(hot, startRow, ordered);
          showStatus("Pasted " + ordered.length + " image(s).", "success");
        }
      };
      reader.onerror = function () {
        completed += 1;
        if (completed === imageItems.length) {
          const ordered = dataUrls.filter(Boolean);
          setImagesSequentially(hot, startRow, ordered);
          showStatus("Some images could not be read from clipboard.", ordered.length ? "info" : "error");
        }
      };
      reader.readAsDataURL(blob);
    });
  }

  function buildRowsForBackend(rawData) {
    const rows = [];
    rawData.forEach(function (row, idx) {
      const image = row[0];
      const vendor = (row[1] || "").toString().trim();
      if (!image || !vendor) return;
      rows.push({
        row_number: idx + 1,
        image_data_url: image,
        vendor_code: vendor
      });
    });
    return rows;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename || "vendor_images.zip";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function processData() {
    const hot = window.hotInstance;
    if (!hot) return;
    const rows = buildRowsForBackend(hot.getData());
    if (!rows.length) {
      showStatus("Need both image (A) and vendor code (B) in at least one row.", "error");
      return;
    }

    const formData = new FormData();
    formData.append("data", JSON.stringify(rows));
    showStatus("Generating ZIP...", "info");

    fetch(apiUrl("/process_data"), { method: "POST", body: formData })
      .then(async function (response) {
        if (!response.ok) {
          const text = await response.text();
          throw new Error(text || "Process failed");
        }
        return response.blob();
      })
      .then(function (blob) {
        downloadBlob(blob, "vendor_images.zip");
        showStatus("ZIP generated successfully.", "success");
      })
      .catch(function (error) {
        showStatus("Error: " + error.message, "error");
      });
  }

  function initializeGrid() {
    const container = document.getElementById("excel-grid");
    if (!container || typeof Handsontable === "undefined") {
      showStatus("Handsontable failed to load.", "error");
      return;
    }

    const seedRows = [];
    for (let i = 0; i < 500; i += 1) {
      seedRows.push(["", ""]);
    }

    const hot = new Handsontable(container, {
      data: seedRows,
      colHeaders: ["A - Images", "B - Vendor Material #"],
      rowHeaders: true,
      height: 460,
      stretchH: "all",
      contextMenu: ["row_above", "row_below", "remove_row", "undo", "redo"],
      minSpareRows: 100,
      manualRowResize: true,
      licenseKey: "non-commercial-and-evaluation",
      columns: [{ renderer: imageRenderer }, { type: "text" }],
      afterPaste: function (_data, coords) {
        if (coords && coords[0] && coords[0].startCol === 1) {
          showStatus("Vendor codes pasted in Column B.", "info");
        }
      }
    });

    window.hotInstance = hot;
    document.addEventListener("paste", function (e) {
      handleClipboardPaste(e, hot);
    });

    const processButton = document.getElementById("processButton");
    if (processButton) {
      processButton.addEventListener("click", processData);
    }
  }

  document.addEventListener("DOMContentLoaded", initializeGrid);
})();
