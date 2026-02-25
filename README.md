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
    const matches = html.match(/src\s*=\s*["']((?:data:image\/|blob:)[^"']+)["']/gi) || [];
    return matches
      .map(function (m) {
        const mm = m.match(/["']((?:data:image\/|blob:)[^"']+)["']/i);
        return mm ? mm[1] : null;
      })
      .filter(Boolean);
  }

  function blobToDataUrl(blob) {
    return new Promise(function (resolve, reject) {
      const reader = new FileReader();
      reader.onload = function (event) {
        resolve(event.target.result);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  async function blobUrlToDataUrl(blobUrl) {
    if (!blobUrl || blobUrl.indexOf("blob:") !== 0) return null;
    try {
      const response = await fetch(blobUrl);
      if (!response.ok) return null;
      const blob = await response.blob();
      if (!blob || !blob.type || blob.type.indexOf("image/") !== 0) return null;
      return await blobToDataUrl(blob);
    } catch (_err) {
      return null;
    }
  }

  function isProbablyRealImage(dataUrl) {
    if (!dataUrl || dataUrl.indexOf("data:image/") !== 0) return false;
    // Filter tiny clipboard artifacts/icons that appear as fake "images".
    return dataUrl.length > 1800;
  }

  function setImagesSequentially(hot, startRow, dataUrls) {
    dataUrls.forEach(function (url, idx) {
      hot.setDataAtCell(startRow + idx, 0, url, "image-paste");
    });
  }

  function pasteVendorCodesFromText(hot, startRow, text) {
    if (!text) return 0;
    const rows = text.replace(/\r/g, "").split("\n").filter(Boolean);
    let pasted = 0;
    for (let i = 0; i < rows.length; i += 1) {
      const cols = rows[i].split("\t");
      // If copied range has at least 2 columns, col index 1 is vendor.
      if (cols.length > 1 && cols[1] && cols[1].trim()) {
        hot.setDataAtCell(startRow + i, 1, cols[1].trim(), "vendor-paste");
        pasted += 1;
      }
    }
    return pasted;
  }

  async function handleClipboardPaste(e, hot) {
    const selected = hot.getSelectedLast();
    if (!selected) return;
    const startRow = selected[0];
    const startCol = selected[1];
    if (startCol !== 0) return;

    const html = e.clipboardData ? e.clipboardData.getData("text/html") : "";
    const plainText = e.clipboardData ? e.clipboardData.getData("text/plain") : "";
    const htmlImages = parseImagesFromHtml(html);
    const items = Array.from((e.clipboardData && e.clipboardData.items) || []);
    const files = Array.from((e.clipboardData && e.clipboardData.files) || []);
    const imageItems = items.filter(function (it) {
      return it.type && it.type.indexOf("image/") === 0;
    });
    const imageFiles = files.filter(function (f) {
      return f && f.type && f.type.indexOf("image/") === 0;
    });

    if (!imageItems.length && !imageFiles.length && !htmlImages.length) return;
    e.preventDefault();
    e.stopPropagation();
    if (typeof e.stopImmediatePropagation === "function") {
      e.stopImmediatePropagation();
    }

    const collected = [];

    for (let i = 0; i < imageFiles.length; i += 1) {
      try {
        const converted = await blobToDataUrl(imageFiles[i]);
        if (converted) collected.push(converted);
      } catch (_err) {
        // Keep going.
      }
    }

    for (let i = 0; i < imageItems.length; i += 1) {
      const item = imageItems[i];
      const blob = item.getAsFile();
      if (!blob) continue;
      try {
        const converted = await blobToDataUrl(blob);
        if (converted) collected.push(converted);
      } catch (_err) {
        // Keep going.
      }
    }

    // Also parse HTML image sources (blob:/data:image) for multi-image clipboard cases.
    for (let i = 0; i < htmlImages.length; i += 1) {
      const src = htmlImages[i];
      if (!src) continue;
      if (src.indexOf("data:image/") === 0) {
        collected.push(src);
      } else if (src.indexOf("blob:") === 0) {
        const converted = await blobUrlToDataUrl(src);
        if (converted) collected.push(converted);
      }
    }

    // Fallback only when browser did not expose image items/files/html images.
    if (!collected.length && htmlImages.length) {
      for (let i = 0; i < htmlImages.length; i += 1) {
        if (htmlImages[i] && htmlImages[i].indexOf("data:image/") === 0) {
          collected.push(htmlImages[i]);
        }
      }
    }

    const filtered = collected.filter(isProbablyRealImage);
    if (!filtered.length) {
      showStatus("No readable images found in clipboard.", "error");
      return;
    }

    setImagesSequentially(hot, startRow, filtered);
    const vendorCount = pasteVendorCodesFromText(hot, startRow, plainText);
    if (vendorCount > 0) {
      showStatus(
        "Pasted " + filtered.length + " image(s) and " + vendorCount + " vendor code(s).",
        "success"
      );
    } else {
      showStatus("Pasted " + filtered.length + " image(s).", "success");
    }
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
      handleClipboardPaste(e, hot).catch(function () {
        showStatus("Clipboard paste failed. Try paste again.", "error");
      });
    }, true);

    const processButton = document.getElementById("processButton");
    if (processButton) {
      processButton.addEventListener("click", processData);
    }
  }

  document.addEventListener("DOMContentLoaded", initializeGrid);
})();
