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


def _normalize_ext(ext, data):
    ext = (ext or "").lower().replace(".", "")
    if ext == "jpeg":
        return "jpg"
    if ext in {"png", "jpg", "gif", "bmp", "tif", "tiff", "webp"}:
        return ext
    return _detect_ext(data)


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


def _natural_sort_key(text):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


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


def _detect_vendor_column(ws):
    max_row = min(ws.max_row or 1, 10)
    max_col = min(ws.max_column or 1, 20)

    header_candidate = None
    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            value = ws.cell(row_idx, col_idx).value
            if isinstance(value, str):
                label = value.strip().lower()
                if "vendor" in label and "material" in label:
                    return col_idx
                if "vendor" in label or "material" in label:
                    header_candidate = header_candidate or col_idx

    if header_candidate:
        return header_candidate

    # Fallback on most populated candidate business columns.
    fallback_candidates = [4, 2, 3, 1]
    max_data_row = min(ws.max_row or 1, 5000)
    best_col = 4
    best_score = -1
    for col_idx in fallback_candidates:
        if col_idx > (ws.max_column or 1):
            continue
        score = 0
        for row_idx in range(2, max_data_row + 1):
            value = ws.cell(row_idx, col_idx).value
            if value not in (None, ""):
                score += 1
        if score > best_score:
            best_col = col_idx
            best_score = score
    return best_col


def _collect_vendor_values(ws, vendor_col, limit=None):
    vendors = []
    max_row = ws.max_row or 1
    for row_idx in range(1, max_row + 1):
        value = ws.cell(row_idx, vendor_col).value
        if value in (None, ""):
            continue
        raw = str(value).strip()
        lower = raw.lower()
        if lower in {"image", "vendor material", "vendor", "material"}:
            continue
        vendors.append(_safe_name(raw))
        if limit and len(vendors) >= limit:
            break
    return vendors


def _extract_media_images(file_bytes):
    media = []
    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as archive:
        media_files = [
            name for name in archive.namelist() if name.lower().startswith("xl/media/")
        ]
        media_files.sort(key=_natural_sort_key)
        for media_name in media_files:
            try:
                data = archive.read(media_name)
            except Exception:
                continue
            if not data:
                continue
            ext = _normalize_ext(Path(media_name).suffix, data)
            media.append(
                {
                    "source": media_name,
                    "ext": ext,
                    "data": data,
                }
            )
    return media


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
    vendor_col = _detect_vendor_column(ws)
    images = list(getattr(ws, "_images", []) or [])

    extracted_count = 0
    skipped_count = 0
    skipped_reasons = []
    seen_filenames = set()
    extraction_mode = "openpyxl_anchor_col_a"

    temp_file = tempfile.NamedTemporaryFile(
        prefix="dataiku_excel_images_",
        suffix=".zip",
        delete=False,
    )
    zip_path = temp_file.name
    temp_file.close()

    try:
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            # Pass 1: strict Column A behavior for classic floating images.
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

                vendor = _read_up(ws, row, vendor_col)
                safe_vendor = _safe_name(vendor) if vendor else "Row_{0}".format(row or idx)

                try:
                    image_data = img._data()
                except Exception as exc:
                    skipped_count += 1
                    skipped_reasons.append(
                        "Image #{0}: could not read image data ({1}).".format(idx, exc)
                    )
                    continue

                ext = _normalize_ext(None, image_data)
                filename = _next_unique_filename(safe_vendor, ext, seen_filenames)
                zip_file.writestr("images/{0}".format(filename), image_data)
                extracted_count += 1

            # Pass 2: if no extracted images but images exist, use any anchored column.
            if extracted_count == 0 and images:
                extraction_mode = "openpyxl_any_anchor_column"
                for idx, img in enumerate(images, start=1):
                    row, col = _anchor_row_col(img)
                    vendor = _read_up(ws, row, vendor_col)
                    safe_vendor = _safe_name(vendor) if vendor else "Row_{0}".format(row or idx)

                    try:
                        image_data = img._data()
                    except Exception as exc:
                        skipped_count += 1
                        skipped_reasons.append(
                            "Image #{0}: could not read image data ({1}).".format(idx, exc)
                        )
                        continue

                    ext = _normalize_ext(None, image_data)
                    filename = _next_unique_filename(safe_vendor, ext, seen_filenames)
                    zip_file.writestr("images/{0}".format(filename), image_data)
                    extracted_count += 1

                if extracted_count > 0:
                    skipped_reasons.append(
                        "No image anchor detected in Column A; fallback used all anchor columns."
                    )

            # Pass 3: fallback for newer Excel 'Picture/in-cell image' files.
            if extracted_count == 0:
                media_images = _extract_media_images(file_bytes)
                if media_images:
                    extraction_mode = "xlsx_media_fallback"
                    vendor_values = _collect_vendor_values(ws, vendor_col, limit=len(media_images))
                    if vendor_values:
                        media_images = media_images[: len(vendor_values)]

                    for idx, media in enumerate(media_images, start=1):
                        if vendor_values and idx - 1 < len(vendor_values):
                            safe_vendor = vendor_values[idx - 1]
                        else:
                            safe_vendor = "Image_{0}".format(idx)
                        filename = _next_unique_filename(
                            safe_vendor,
                            media["ext"],
                            seen_filenames,
                        )
                        zip_file.writestr("images/{0}".format(filename), media["data"])
                        extracted_count += 1

                    skipped_reasons.append(
                        "Used XLSX media fallback mode because no readable image anchors were found."
                    )

            summary_lines = [
                "Excel Image Extraction Summary",
                "==============================",
                "Generated at: {0}Z".format(datetime.utcnow().isoformat()),
                "Sheet: {0}".format(sheet_name),
                "Extraction mode: {0}".format(extraction_mode),
                "Detected vendor column: {0}".format(vendor_col),
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
