(function () {
  "use strict";

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
      const row = startRow + idx;
      hot.setDataAtCell(row, 0, url, "image-paste");
    });
  }

  function handleClipboardPaste(e, hot) {
    const selected = hot.getSelectedLast();
    if (!selected) return;
    const startRow = selected[0];
    const startCol = selected[1];
    if (startCol !== 0) return;

    const items = Array.from((e.clipboardData && e.clipboardData.items) || []);
    const imageItems = items.filter(function (it) {
      return it.type && it.type.indexOf("image/") === 0;
    });
    const html = e.clipboardData ? e.clipboardData.getData("text/html") : "";
    const htmlImages = parseImagesFromHtml(html);

    if (!imageItems.length && !htmlImages.length) return;
    e.preventDefault();

    if (htmlImages.length) {
      setImagesSequentially(hot, startRow, htmlImages);
      showStatus("Pasted " + htmlImages.length + " image(s) from HTML clipboard.", "success");
      return;
    }

    let completed = 0;
    const dataUrls = new Array(imageItems.length);
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
          showStatus("Pasted " + ordered.length + " image(s) from clipboard.", "success");
        }
      };
      reader.onerror = function () {
        completed += 1;
        if (completed === imageItems.length) {
          const ordered = dataUrls.filter(Boolean);
          setImagesSequentially(hot, startRow, ordered);
          showStatus("Some clipboard images failed to load.", ordered.length ? "info" : "error");
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
        vendor_code: vendor,
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
      showStatus("No valid rows found. Need both Image and Vendor Code.", "error");
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
      showStatus("Grid library failed to load.", "error");
      return;
    }

    const seedRows = [];
    for (let i = 0; i < 300; i += 1) {
      seedRows.push(["", ""]);
    }

    const hot = new Handsontable(container, {
      data: seedRows,
      colHeaders: ["Image", "Vendor Code"],
      rowHeaders: true,
      height: 460,
      stretchH: "all",
      contextMenu: true,
      manualRowResize: true,
      licenseKey: "non-commercial-and-evaluation",
      columns: [{ renderer: imageRenderer }, { type: "text" }],
      afterPaste: function (_data, coords) {
        // If plain text paste starts at vendor column, keep default behavior.
        if (coords && coords[0] && coords[0].startCol === 1) {
          showStatus("Vendor codes pasted.", "info");
        }
      },
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
(function () {
  "use strict";

  function apiUrl(path) {
    const normalized = path.startsWith("/") ? path : "/" + path;
    if (typeof getWebAppBackendUrl === "function") {
      return getWebAppBackendUrl(normalized);
    }
    return "backend" + normalized;
  }

  function buildCandidateUrls(path) {
    const normalized = path.startsWith("/") ? path : "/" + path;
    if (typeof getWebAppBackendUrl === "function") {
      return [getWebAppBackendUrl(normalized)];
    }
    return ["backend" + normalized, "/backend" + normalized, normalized];
  }

  function apiFetch(path, options) {
    const urls = buildCandidateUrls(path);
    let lastError = null;
    function tryAt(index) {
      if (index >= urls.length) {
        throw lastError || new Error("Backend unreachable");
      }
      return fetch(urls[index], options || {})
        .then(function (response) {
          if (response.status === 404) {
            lastError = new Error("Endpoint not found: " + urls[index]);
            return tryAt(index + 1);
          }
          return response;
        })
        .catch(function (error) {
          lastError = error;
          return tryAt(index + 1);
        });
    }
    return tryAt(0);
  }

  function init() {
    const uploadForm = document.getElementById("uploadForm");
    const sheetSelection = document.getElementById("sheetSelection");
    const sheetList = document.getElementById("sheetList");
    const previewMappingButton = document.getElementById("previewMappingButton");
    const mappingPreview = document.getElementById("mappingPreview");
    const processButton = document.getElementById("processButton");
    const spinner = document.getElementById("spinner");
    const progressBar = document.getElementById("progressBar");
    const progress = document.getElementById("progress");
    const status = document.getElementById("status");
    const uploadButton = document.getElementById("uploadButton");
    const excelFileInput = document.getElementById("excelFile");

    if (
      !uploadForm ||
      !sheetSelection ||
      !sheetList ||
      !previewMappingButton ||
      !mappingPreview ||
      !processButton ||
      !spinner ||
      !progressBar ||
      !progress ||
      !status ||
      !uploadButton ||
      !excelFileInput
    ) {
      return;
    }

    let progressInterval = null;
    let selectedSheetName = "";

    function showStatus(message, type) {
      const statusType = type || "info";
      status.className = "status " + statusType;
      status.innerHTML =
        '<div class="status-message">' +
        (statusType === "error" ? "❌ " : statusType === "success" ? "✅ " : "ℹ️ ") +
        message +
        "</div>";
    }

    function clearStatus() {
      status.className = "status";
      status.textContent = "";
    }

    function showSpinner() {
      spinner.style.display = "block";
    }

    function hideSpinner() {
      spinner.style.display = "none";
    }

    function showProgressBar() {
      if (progressInterval) {
        clearInterval(progressInterval);
      }
      progressBar.style.display = "block";
      progress.style.width = "0%";
      let width = 0;
      progressInterval = setInterval(function () {
        if (width >= 90) {
          clearInterval(progressInterval);
          progressInterval = null;
          return;
        }
        width += 1;
        progress.style.width = width + "%";
      }, 50);
    }

    function hideProgressBar() {
      if (progressInterval) {
        clearInterval(progressInterval);
        progressInterval = null;
      }
      progress.style.width = "100%";
      setTimeout(function () {
        progressBar.style.display = "none";
        progress.style.width = "0%";
      }, 400);
    }

    function parseError(response) {
      return response
        .clone()
        .json()
        .then(function (body) {
          return body.message || body.error || JSON.stringify(body);
        })
        .catch(function () {
          return response
            .clone()
            .text()
            .then(function (text) {
              return text || ("Request failed (" + response.status + ")");
            })
            .catch(function () {
              return "Request failed (" + response.status + ")";
            });
        });
    }

    function resetSheetSelection() {
      sheetSelection.style.display = "none";
      sheetList.innerHTML = "";
      selectedSheetName = "";
      processButton.disabled = true;
      previewMappingButton.disabled = true;
      mappingPreview.style.display = "none";
      mappingPreview.innerHTML = "";
    }

    function showSheetSelection(sheets) {
      sheetList.innerHTML = "";
      sheets.forEach(function (sheet) {
        const option = document.createElement("div");
        option.className = "sheet-option";
        option.textContent = sheet;
        option.addEventListener("click", function () {
          document.querySelectorAll(".sheet-option").forEach(function (el) {
            el.classList.remove("selected");
          });
          option.classList.add("selected");
          selectedSheetName = sheet;
          previewMappingButton.disabled = false;
          processButton.disabled = false;
          mappingPreview.style.display = "none";
          mappingPreview.innerHTML = "";
          showStatus("Sheet selected. Preview mapping or process directly.", "info");
        });
        sheetList.appendChild(option);
      });
      sheetSelection.style.display = "block";
    }

    function getSheets(formData) {
      return apiFetch("/get_sheets", { method: "POST", body: formData }).then(function (response) {
        if (!response.ok) {
          return parseError(response).then(function (message) {
            throw new Error(message);
          });
        }
        return response.json();
      });
    }

    function extractImages(formData) {
      return apiFetch("/extract_images", { method: "POST", body: formData }).then(function (response) {
        if (!response.ok) {
          return parseError(response).then(function (message) {
            throw new Error(message);
          });
        }
        return response.blob();
      });
    }

    function previewMapping(formData) {
      return apiFetch("/preview_mapping", { method: "POST", body: formData }).then(function (response) {
        if (!response.ok) {
          return parseError(response).then(function (message) {
            throw new Error(message);
          });
        }
        return response.json();
      });
    }

    function renderPreview(preview) {
      const warnings = preview.warnings || [];
      const pivot = preview.pivot || [];
      let html =
        "<h4>Mapping Preview</h4>" +
        "<div>Target rows: " +
        preview.target_row_count +
        " | Mapped rows: " +
        preview.mapped_row_count +
        " | Mode: " +
        (preview.extraction_mode || "none") +
        "</div>";

      if (warnings.length) {
        html += "<ul>";
        warnings.forEach(function (w) {
          html += "<li class='preview-warning'>" + w + "</li>";
        });
        html += "</ul>";
      } else {
        html += "<div>No warnings detected.</div>";
      }

      html +=
        "<table class='preview-table'><thead><tr><th>Vendor</th><th>Rows</th><th>Missing</th><th>Unique Images</th><th>Row List</th></tr></thead><tbody>";
      pivot.slice(0, 200).forEach(function (item) {
        html +=
          "<tr><td>" +
          (item.vendor_code || "[NO_VENDOR]") +
          "</td><td>" +
          item.row_count +
          "</td><td>" +
          item.missing_images +
          "</td><td>" +
          item.unique_image_count +
          "</td><td>" +
          (item.rows || []).join(",") +
          "</td></tr>";
      });
      html += "</tbody></table>";
      mappingPreview.innerHTML = html;
      mappingPreview.style.display = "block";
    }

    function buildDownloadName() {
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      return "Excel_Images_" + timestamp + ".zip";
    }

    function runHealthCheck() {
      apiFetch("/health")
        .then(function (response) {
          if (!response.ok) {
            throw new Error("Health check failed");
          }
          return response.json();
        })
        .then(function (data) {
          if (data.status !== "ok") {
            showStatus("Warning: Backend health check returned unexpected status.", "error");
          }
        })
        .catch(function () {
          showStatus("Warning: Cannot connect to backend. Check your webapp backend logs.", "error");
        });
    }

    previewMappingButton.addEventListener("click", function () {
      if (!selectedSheetName) {
        showStatus("Select a sheet first.", "error");
        return;
      }
      const file = excelFileInput.files && excelFileInput.files[0];
      if (!file) {
        showStatus("Please upload an Excel file first.", "error");
        return;
      }

      const formData = new FormData();
      formData.append("file", file);
      formData.append("sheet_name", selectedSheetName);

      previewMappingButton.disabled = true;
      showSpinner();
      showStatus("Building mapping preview...", "info");
      previewMapping(formData)
        .then(function (result) {
          if (!result || result.status !== "ok" || !result.preview) {
            throw new Error("Preview response is invalid.");
          }
          renderPreview(result.preview);
          showStatus("Preview ready. Check warnings before processing.", "success");
        })
        .catch(function (error) {
          showStatus("Preview error: " + error.message, "error");
        })
        .finally(function () {
          hideSpinner();
          previewMappingButton.disabled = false;
        });
    });

    excelFileInput.addEventListener("change", function (event) {
      const file = event.target.files && event.target.files[0];
      uploadButton.textContent = file ? "Upload: " + file.name : "Upload Excel File";
      resetSheetSelection();
      clearStatus();
    });

    uploadForm.addEventListener("submit", function (event) {
      event.preventDefault();
      resetSheetSelection();
      clearStatus();

      const file = excelFileInput.files && excelFileInput.files[0];
      if (!file) {
        showStatus("Please select an Excel file.", "error");
        return;
      }
      if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
        showStatus("Please upload a valid Excel file (.xlsx or .xlsm).", "error");
        return;
      }
      if (file.size > 50 * 1024 * 1024) {
        showStatus("File too large. Please upload a file up to 50MB.", "error");
        return;
      }

      const formData = new FormData();
      formData.append("file", file);

      uploadButton.disabled = true;
      showSpinner();
      showStatus("Reading Excel file...", "info");

      getSheets(formData)
        .then(function (data) {
          if (!Array.isArray(data.sheets) || data.sheets.length === 0) {
            throw new Error("No sheets found in the Excel file.");
          }
          showSheetSelection(data.sheets);
          showStatus("Select a sheet containing images.", "info");
        })
        .catch(function (error) {
          showStatus("Error: " + error.message, "error");
        })
        .finally(function () {
          hideSpinner();
          uploadButton.disabled = false;
        });
    });

    processButton.addEventListener("click", function () {
      if (!selectedSheetName) {
        showStatus("Please select a sheet.", "error");
        return;
      }
      const file = excelFileInput.files && excelFileInput.files[0];
      if (!file) {
        showStatus("Please upload an Excel file first.", "error");
        return;
      }

      const formData = new FormData();
      formData.append("file", file);
      formData.append("sheet_name", selectedSheetName);

      processButton.disabled = true;
      showSpinner();
      showProgressBar();
      showStatus("Processing images...", "info");

      extractImages(formData)
        .then(function (blob) {
          const url = window.URL.createObjectURL(blob);
          const anchor = document.createElement("a");
          anchor.href = url;
          anchor.download = buildDownloadName();
          document.body.appendChild(anchor);
          anchor.click();
          document.body.removeChild(anchor);
          window.URL.revokeObjectURL(url);
          showStatus("Images extracted successfully! Check your Downloads folder.", "success");
        })
        .catch(function (error) {
          showStatus("Error: " + error.message, "error");
        })
        .finally(function () {
          hideSpinner();
          hideProgressBar();
          processButton.disabled = false;
        });
    });

    runHealthCheck();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
