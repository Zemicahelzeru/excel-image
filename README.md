(function () {
  "use strict";

  function apiUrl(path) {
    const normalized = path.startsWith("/") ? path : "/" + path;
    if (typeof getWebAppBackendUrl === "function") {
      return getWebAppBackendUrl(normalized);
    }
    return "backend" + normalized;
  }

  function init() {
    const uploadForm = document.getElementById("uploadForm");
    const sheetSelection = document.getElementById("sheetSelection");
    const sheetList = document.getElementById("sheetList");
    const pasteImageColumnSection = document.getElementById("pasteImageColumnSection");
    const gridBody = document.getElementById("gridBody");
    const startRowInput = document.getElementById("startRowInput");
    const validateGridButton = document.getElementById("validateGridButton");
    const clearGridButton = document.getElementById("clearGridButton");
    const gridStatus = document.getElementById("gridStatus");
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
      !pasteImageColumnSection ||
      !gridBody ||
      !startRowInput ||
      !validateGridButton ||
      !clearGridButton ||
      !gridStatus ||
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
    let validatedRows = [];
    const gridSize = 200;

    function showStatus(message, type) {
      const statusType = type || "info";
      status.className = "status " + statusType;
      status.innerHTML =
        '<div class="status-message">' +
        (statusType === "error" ? "❌ " : statusType === "success" ? "✅ " : "ℹ️ ") +
        message +
        "</div>";
    }

    function showGridStatus(message, type) {
      const statusType = type || "info";
      gridStatus.className = "status " + statusType;
      gridStatus.innerHTML =
        '<div class="status-message">' +
        (statusType === "error" ? "❌ " : statusType === "success" ? "✅ " : "ℹ️ ") +
        message +
        "</div>";
    }

    function clearStatus() {
      status.className = "status";
      status.textContent = "";
    }

    function clearGridStatus() {
      gridStatus.className = "status";
      gridStatus.textContent = "";
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

    function initializeGrid() {
      gridBody.innerHTML = "";
      for (let i = 1; i <= gridSize; i += 1) {
        const row = document.createElement("tr");
        row.innerHTML =
          '<td class="row-num">' +
          i +
          "</td><td><input type='text' class='grid-cell' data-grid-row='" +
          i +
          "' placeholder='Paste row " +
          i +
          "' /></td>";
        gridBody.appendChild(row);
      }
    }

    function clearGrid() {
      const cells = gridBody.querySelectorAll(".grid-cell");
      cells.forEach(function (cell) {
        cell.value = "";
      });
      validatedRows = [];
      processButton.disabled = true;
      clearGridStatus();
    }

    function collectValidatedRows() {
      const firstExcelRow = Math.max(1, parseInt(startRowInput.value || "1", 10) || 1);
      const rows = [];
      const cells = gridBody.querySelectorAll(".grid-cell");
      cells.forEach(function (cell) {
        const value = (cell.value || "").trim();
        if (!value) {
          return;
        }
        const gridRow = parseInt(cell.getAttribute("data-grid-row") || "0", 10);
        if (!gridRow) {
          return;
        }
        rows.push(firstExcelRow + gridRow - 1);
      });
      return rows;
    }

    function resetSheetSelection() {
      sheetSelection.style.display = "none";
      sheetList.innerHTML = "";
      pasteImageColumnSection.style.display = "none";
      selectedSheetName = "";
      validatedRows = [];
      processButton.disabled = true;
      clearGridStatus();
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
          pasteImageColumnSection.style.display = "block";
          validatedRows = [];
          processButton.disabled = true;
          clearGridStatus();
          showStatus("Paste Column A and click Validate Grid.", "info");
        });
        sheetList.appendChild(option);
      });
      sheetSelection.style.display = "block";
    }

    function getSheets(formData) {
      return fetch(apiUrl("/get_sheets"), { method: "POST", body: formData }).then(function (response) {
        if (!response.ok) {
          return parseError(response).then(function (message) {
            throw new Error(message);
          });
        }
        return response.json();
      });
    }

    function extractImages(formData) {
      return fetch(apiUrl("/extract_images"), { method: "POST", body: formData }).then(function (response) {
        if (!response.ok) {
          return parseError(response).then(function (message) {
            throw new Error(message);
          });
        }
        return response.blob();
      });
    }

    function buildDownloadName() {
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      return "Excel_Images_" + timestamp + ".zip";
    }

    function runHealthCheck() {
      fetch(apiUrl("/health"))
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

    gridBody.addEventListener("paste", function (event) {
      const target = event.target;
      if (!target || !target.classList || !target.classList.contains("grid-cell")) {
        return;
      }
      event.preventDefault();
      const text = (event.clipboardData || window.clipboardData).getData("text");
      const rows = text.split(/\r?\n/);
      const startGridRow = parseInt(target.getAttribute("data-grid-row") || "1", 10);
      rows.forEach(function (line, idx) {
        const rowNumber = startGridRow + idx;
        const cell = gridBody.querySelector(".grid-cell[data-grid-row='" + rowNumber + "']");
        if (cell) {
          cell.value = (line || "").trim();
        }
      });
    });

    validateGridButton.addEventListener("click", function () {
      const rows = collectValidatedRows();
      if (!rows.length) {
        showGridStatus("No non-empty values found in pasted Column A.", "error");
        processButton.disabled = true;
        return;
      }
      validatedRows = rows;
      showGridStatus("Detected " + rows.length + " image rows. Ready to process.", "success");
      processButton.disabled = false;
      showStatus("Grid validated. You can now process the sheet.", "success");
    });

    clearGridButton.addEventListener("click", function () {
      clearGrid();
      showStatus("Grid cleared.", "info");
    });

    excelFileInput.addEventListener("change", function (event) {
      const file = event.target.files && event.target.files[0];
      uploadButton.textContent = file ? "Upload: " + file.name : "Upload Excel File";
      resetSheetSelection();
      clearGrid();
      clearStatus();
    });

    uploadForm.addEventListener("submit", function (event) {
      event.preventDefault();
      resetSheetSelection();
      clearGrid();
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
      if (!validatedRows.length) {
        showStatus("Please validate pasted Column A first.", "error");
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
      formData.append("pasted_image_rows", JSON.stringify(validatedRows));

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

    initializeGrid();
    runHealthCheck();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
