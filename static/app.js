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

let progressInterval = null;

const API_PREFIXES = ["backend", "/backend", ""];

function showStatus(message, type = "info") {
  status.className = `status ${type}`;
  status.innerHTML = `
    <div class="status-message">
      ${type === "error" ? "❌ " : type === "success" ? "✅ " : "ℹ️ "}
      ${message}
    </div>
  `;
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
  progressInterval = setInterval(() => {
    if (width >= 90) {
      clearInterval(progressInterval);
      progressInterval = null;
      return;
    }
    width += 1;
    progress.style.width = `${width}%`;
  }, 50);
}

function hideProgressBar() {
  if (progressInterval) {
    clearInterval(progressInterval);
    progressInterval = null;
  }
  progress.style.width = "100%";
  setTimeout(() => {
    progressBar.style.display = "none";
    progress.style.width = "0%";
  }, 400);
}

async function parseError(response) {
  const clone = response.clone();
  try {
    const body = await clone.json();
    return body.message || body.error || JSON.stringify(body);
  } catch (_) {
    try {
      return await clone.text();
    } catch (err) {
      return `Request failed (${response.status})`;
    }
  }
}

function buildCandidateUrls(path) {
  const normalized = path.startsWith("/") ? path : `/${path}`;
  return API_PREFIXES.map((prefix) => (prefix ? `${prefix}${normalized}` : normalized));
}

async function apiFetch(path, options = {}) {
  const urls = buildCandidateUrls(path);
  let lastError = null;

  for (const url of urls) {
    try {
      const response = await fetch(url, options);
      if (response.status !== 404) {
        return response;
      }
      lastError = new Error(`Endpoint not found: ${url}`);
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Backend unreachable");
}

async function getSheets(formData) {
  const response = await apiFetch("/get_sheets", { method: "POST", body: formData });
  if (!response.ok) {
    throw new Error(await parseError(response));
  }
  return response.json();
}

async function extractImages(formData) {
  const response = await apiFetch("/extract_images", { method: "POST", body: formData });
  if (!response.ok) {
    throw new Error(await parseError(response));
  }
  return response.blob();
}

function resetSheetSelection() {
  sheetSelection.style.display = "none";
  sheetList.innerHTML = "";
  processButton.disabled = true;
}

function showSheetSelection(sheets) {
  sheetList.innerHTML = "";

  sheets.forEach((sheet) => {
    const option = document.createElement("div");
    option.className = "sheet-option";
    option.textContent = sheet;

    option.addEventListener("click", function () {
      document.querySelectorAll(".sheet-option").forEach((el) => el.classList.remove("selected"));
      this.classList.add("selected");
      processButton.disabled = false;
      showStatus('Click "Process Selected Sheet" to extract images', "info");
    });

    sheetList.appendChild(option);
  });

  sheetSelection.style.display = "block";
}

function buildDownloadName() {
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `Excel_Images_${timestamp}.zip`;
}

async function runHealthCheck() {
  try {
    const response = await apiFetch("/health");
    if (!response.ok) {
      throw new Error(await parseError(response));
    }
    const data = await response.json();
    if (data.status !== "ok") {
      showStatus("Warning: Backend health check returned unexpected status.", "error");
    }
  } catch (_) {
    showStatus("Warning: Cannot connect to backend. Check server logs.", "error");
  }
}

excelFileInput.addEventListener("change", (event) => {
  const fileName = event.target.files[0]?.name || "";
  uploadButton.textContent = fileName ? `Upload: ${fileName}` : "Upload Excel File";
  resetSheetSelection();
  clearStatus();
});

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  resetSheetSelection();
  clearStatus();

  const file = excelFileInput.files[0];
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

  try {
    const data = await getSheets(formData);
    if (data.status === "error") {
      throw new Error(data.message || "Failed to read sheet names.");
    }
    if (!Array.isArray(data.sheets) || data.sheets.length === 0) {
      throw new Error("No sheets found in the Excel file.");
    }

    showSheetSelection(data.sheets);
    showStatus("Select a sheet containing images.", "info");
  } catch (error) {
    showStatus(`Error: ${error.message}`, "error");
    console.error("get_sheets error:", error);
  } finally {
    hideSpinner();
    uploadButton.disabled = false;
  }
});

processButton.addEventListener("click", async () => {
  const selectedSheet = document.querySelector(".sheet-option.selected");
  if (!selectedSheet) {
    showStatus("Please select a sheet.", "error");
    return;
  }

  const file = excelFileInput.files[0];
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

  try {
    const blob = await extractImages(formData);
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = buildDownloadName();
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);

    showStatus("Images extracted successfully! Check your Downloads folder.", "success");
  } catch (error) {
    showStatus(`Error: ${error.message}`, "error");
    console.error("extract_images error:", error);
  } finally {
    hideSpinner();
    hideProgressBar();
    processButton.disabled = false;
  }
});

runHealthCheck();
