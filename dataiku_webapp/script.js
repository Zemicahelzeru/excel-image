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

    function friendlyErrorMessage(error, phase) {
      const message = (error && error.message) || String(error || "");
      if (/Failed to fetch/i.test(message)) {
        return (
          "Cannot reach backend during " +
          phase +
          ". In Dataiku this usually means backend crash, backend timeout, or backend restart. " +
          "Open webapp backend logs and restart the backend."
        );
      }
      return message || "Unexpected error";
    }

    function resetSheetSelection() {
      sheetSelection.style.display = "none";
      sheetList.innerHTML = "";
      processButton.disabled = true;
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
          processButton.disabled = false;
          showStatus('Click "Process Selected Sheet" to extract images', "info");
        });

        sheetList.appendChild(option);
      });
      sheetSelection.style.display = "block";
    }

    function getSheets(formData) {
      return fetch(apiUrl("/get_sheets"), {
        method: "POST",
        body: formData,
        credentials: "same-origin",
        cache: "no-store",
      }).then(
        function (response) {
          if (!response.ok) {
            return parseError(response).then(function (message) {
              throw new Error(message);
            });
          }
          return response.json();
        }
      );
    }

    function extractImages(formData) {
      return fetch(apiUrl("/extract_images"), {
        method: "POST",
        body: formData,
        credentials: "same-origin",
        cache: "no-store",
      }).then(
        function (response) {
          if (!response.ok) {
            return parseError(response).then(function (message) {
              throw new Error(message);
            });
          }
          return response.blob();
        }
      );
    }

    function buildDownloadName() {
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
      return "Excel_Images_" + timestamp + ".zip";
    }

    function runHealthCheck() {
      fetch(apiUrl("/health"), { credentials: "same-origin", cache: "no-store" })
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
          showStatus("Error: " + friendlyErrorMessage(error, "sheet loading"), "error");
          console.error("get_sheets error:", error);
        })
        .finally(function () {
          hideSpinner();
          uploadButton.disabled = false;
        });
    });

    processButton.addEventListener("click", function () {
      const selectedSheet = document.querySelector(".sheet-option.selected");
      if (!selectedSheet) {
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
      formData.append("sheet_name", selectedSheet.textContent);

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
          showStatus("Error: " + friendlyErrorMessage(error, "image extraction"), "error");
          console.error("extract_images error:", error);
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
