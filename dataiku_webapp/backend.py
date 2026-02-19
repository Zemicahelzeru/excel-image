import io
import os
import re
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from flask import Flask, after_this_request, jsonify, request, send_file
from werkzeug.exceptions import HTTPException

# In Dataiku webapps, "app" is usually already provided.
# This fallback keeps the file runnable outside Dataiku for local testing.
try:
    app  # type: ignore[name-defined]
except NameError:
    app = Flask(__name__)

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024


def _json_error(message, status_code=400):
    return jsonify(status="error", message=message), status_code


def _detect_ext(data):
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if data.startswith(b"GIF8"):
        return "gif"
    return "png"


def _safe_name(value):
    value = str(value).strip() if value not in (None, "") else "Image"
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-") or "Image"


def _anchor_row_col(img):
    anchor = getattr(img, "anchor", None)
    if anchor:
        from_cell = getattr(anchor, "_from", None)
        if from_cell:
            return from_cell.row + 1, from_cell.col + 1
    return None, None


def _read_up(ws, row, col):
    if not row:
        return None
    for r in range(row, 0, -1):
        value = ws.cell(r, col).value
        if value not in (None, ""):
            return value
    return None


def _next_unique_filename(base_name, ext, seen):
    candidate = "{0}.{1}".format(base_name, ext)
    if candidate not in seen:
        seen.add(candidate)
        return candidate

    counter = 2
    while True:
        candidate = "{0}_{1}.{2}".format(base_name, counter, ext)
        if candidate not in seen:
            seen.add(candidate)
            return candidate
        counter += 1


def _get_uploaded_file():
    file_obj = request.files.get("file")
    if not file_obj or not file_obj.filename:
        return None, "No file uploaded"

    ext = Path(file_obj.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return None, "Please upload a valid Excel file (.xlsx or .xlsm)"

    file_bytes = file_obj.read()
    if not file_bytes:
        return None, "Uploaded file is empty"
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        return None, "File too large. Please upload a file up to 50MB."

    return file_bytes, None


def _send_zip_response(stream, filename):
    send_kwargs = dict(
        mimetype="application/zip",
        as_attachment=True,
        conditional=False,
        etag=False,
        max_age=0,
    )
    try:
        return send_file(stream, download_name=filename, **send_kwargs)
    except TypeError:
        # Compatibility with older Flask versions.
        return send_file(stream, attachment_filename=filename, **send_kwargs)


@app.errorhandler(Exception)
def _handle_unexpected_error(exc):
    if isinstance(exc, HTTPException):
        return exc
    app.logger.exception("Unhandled webapp backend exception")
    return _json_error("Backend error while processing request: {0}".format(exc), 500)


@app.route("/health")
def health():
    return jsonify(status="ok")


@app.route("/get_sheets", methods=["POST"])
def get_sheets():
    file_bytes, error = _get_uploaded_file()
    if error:
        return _json_error(error, 400)

    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception as exc:
        return _json_error("Could not read Excel file: {0}".format(exc), 400)

    if not sheet_names:
        return _json_error("No sheets found in the Excel file", 400)

    return jsonify(status="ok", sheets=sheet_names)


@app.route("/extract_images", methods=["POST"])
def extract_images():
    file_bytes, error = _get_uploaded_file()
    if error:
        return _json_error(error, 400)

    sheet_name = (request.form.get("sheet_name") or "").strip()
    if not sheet_name:
        return _json_error("Missing sheet_name", 400)

    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, keep_links=False)
    except Exception as exc:
        return _json_error("Could not open workbook: {0}".format(exc), 400)

    if sheet_name not in wb.sheetnames:
        wb.close()
        return _json_error('Sheet "{0}" not found in workbook'.format(sheet_name), 400)

    ws = wb[sheet_name]
    images = list(getattr(ws, "_images", []) or [])

    extracted_count = 0
    skipped_count = 0
    skipped_reasons = []
    seen_filenames = set()

    temp_file = tempfile.NamedTemporaryFile(
        prefix="dataiku_excel_images_",
        suffix=".zip",
        delete=False,
    )
    zip_path = temp_file.name
    temp_file.close()

    try:
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            for idx, img in enumerate(images, start=1):
                row, col = _anchor_row_col(img)
                if col != 1:
                    skipped_count += 1
                    skipped_reasons.append(
                        "Image #{0}: skipped (not in Column A). Found column={1}.".format(
                            idx, col or "unknown"
                        )
                    )
                    continue

                vendor = _read_up(ws, row, 4)
                safe_vendor = _safe_name(vendor) if vendor else "Row_{0}".format(row or idx)

                try:
                    image_data = img._data()
                except Exception as exc:
                    skipped_count += 1
                    skipped_reasons.append(
                        "Image #{0}: could not read image data ({1}).".format(idx, exc)
                    )
                    continue

                ext = _detect_ext(image_data)
                filename = _next_unique_filename(safe_vendor, ext, seen_filenames)
                zip_file.writestr("images/{0}".format(filename), image_data)
                extracted_count += 1

            summary_lines = [
                "Excel Image Extraction Summary",
                "==============================",
                "Generated at: {0}Z".format(datetime.utcnow().isoformat()),
                "Sheet: {0}".format(sheet_name),
                "Total images found in sheet: {0}".format(len(images)),
                "Extracted images: {0}".format(extracted_count),
                "Skipped images: {0}".format(skipped_count),
                "",
                "Rules:",
                "- Images are extracted only when anchored in Column A.",
                "- File names are based on nearest non-empty value above/in Column D.",
                "",
            ]
            if skipped_reasons:
                summary_lines.append("Skipped details:")
                summary_lines.extend("- {0}".format(reason) for reason in skipped_reasons)

            zip_file.writestr("summary.txt", "\n".join(summary_lines))
    except Exception:
        try:
            os.remove(zip_path)
        except OSError:
            pass
        wb.close()
        raise

    wb.close()

    if extracted_count == 0:
        try:
            os.remove(zip_path)
        except OSError:
            pass
        return _json_error(
            "No images were extracted. Ensure images are in Column A and not empty.",
            400,
        )

    @after_this_request
    def _cleanup_temp_file(response):
        try:
            os.remove(zip_path)
        except OSError:
            pass
        return response

    return _send_zip_response(zip_path, "images.zip")
