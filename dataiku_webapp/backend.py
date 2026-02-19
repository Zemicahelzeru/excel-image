import io
import os
import posixpath
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from flask import Flask, after_this_request, jsonify, request, send_file
from werkzeug.exceptions import HTTPException

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None

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
        return None, None, "No file uploaded"

    ext = Path(file_obj.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return None, None, "Please upload a valid Excel file (.xlsx or .xlsm)"

    file_bytes = file_obj.read()
    if not file_bytes:
        return None, None, "Uploaded file is empty"
    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        return None, None, "File too large. Please upload a file up to 50MB."

    return file_bytes, file_obj.filename, None


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


def _normalize_label(value):
    if not isinstance(value, str):
        return ""
    return " ".join(value.strip().lower().split())


def _is_image_header_label(label):
    if not label:
        return False
    return (
        label == "image"
        or "image" in label
        or "picture" in label
        or "photo" in label
    )


def _is_vendor_header_label(label):
    return bool(label and ("vendor" in label and "material" in label))


def _is_material_header_label(label):
    if not label:
        return False
    if "vendor" in label:
        return False
    return (
        "original material" in label
        or label.startswith("material")
        or "material #" in label
    )


def _safe_folder_name(filename):
    name = (filename or "").strip()
    if not name:
        return "Excel_Images"
    # Keep original casing/spaces as much as possible while preventing invalid paths.
    name = name.replace("/", "_").replace("\\", "_")
    name = re.sub(r"[\x00-\x1f]+", "", name).strip()
    return name or "Excel_Images"


def _detect_vendor_column(ws):
    max_row = min(ws.max_row or 1, 60)
    max_col = min(ws.max_column or 1, 30)

    header_candidate = None
    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            label = _normalize_label(ws.cell(row_idx, col_idx).value)
            if _is_vendor_header_label(label):
                return col_idx
            if "vendor" in label or ("material" in label and "original" not in label):
                header_candidate = header_candidate or col_idx

    if header_candidate:
        return header_candidate

    fallback_candidates = [4, 2, 3, 1]
    max_data_row = min(ws.max_row or 1, 10000)
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


def _detect_material_column(ws, vendor_col):
    max_row = min(ws.max_row or 1, 60)
    max_col = min(ws.max_column or 1, 40)
    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            label = _normalize_label(ws.cell(row_idx, col_idx).value)
            if _is_material_header_label(label):
                return col_idx
    # Common layout: Vendor in D, Original Material in F.
    if vendor_col and vendor_col + 2 <= (ws.max_column or 1):
        return vendor_col + 2
    return None


def _detect_layout(ws):
    max_row = min(ws.max_row or 1, 60)
    max_col = min(ws.max_column or 1, 40)
    image_header = None
    vendor_header = None
    material_header = None

    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            label = _normalize_label(ws.cell(row_idx, col_idx).value)
            if not label:
                continue
            if image_header is None and _is_image_header_label(label):
                image_header = (row_idx, col_idx)
            if vendor_header is None and _is_vendor_header_label(label):
                vendor_header = (row_idx, col_idx)
            if material_header is None and _is_material_header_label(label):
                material_header = (row_idx, col_idx)

    image_col = image_header[1] if image_header else 1
    vendor_col = vendor_header[1] if vendor_header else _detect_vendor_column(ws)
    material_col = material_header[1] if material_header else _detect_material_column(ws, vendor_col)

    header_rows = [
        row
        for row, _ in [image_header, vendor_header, material_header]
        if row is not None
    ]
    start_row = max(header_rows) + 1 if header_rows else 2

    return {
        "image_col": image_col,
        "vendor_col": vendor_col,
        "material_col": material_col,
        "start_row": start_row,
        "image_header_row": image_header[0] if image_header else None,
        "vendor_header_row": vendor_header[0] if vendor_header else None,
        "material_header_row": material_header[0] if material_header else None,
    }


def _row_code(ws, row_idx, vendor_col, material_col):
    vendor_value = ws.cell(row_idx, vendor_col).value if vendor_col else None
    vendor_text = str(vendor_value).strip() if vendor_value not in (None, "") else ""
    vendor_label = _normalize_label(vendor_text)
    if vendor_text and not _is_vendor_header_label(vendor_label):
        return _safe_name(vendor_text), "vendor"

    material_value = ws.cell(row_idx, material_col).value if material_col else None
    material_text = str(material_value).strip() if material_value not in (None, "") else ""
    material_label = _normalize_label(material_text)
    if material_text and not _is_material_header_label(material_label):
        return _safe_name("MAT_{0}".format(material_text)), "material"

    return None, None


def _candidate_rows_for_media(ws, start_row, image_col, vendor_col, material_col):
    rows = []
    max_row = ws.max_row or start_row
    for row_idx in range(start_row, max_row + 1):
        code, _ = _row_code(ws, row_idx, vendor_col, material_col)
        image_cell = ws.cell(row_idx, image_col).value if image_col else None
        image_label = _normalize_label(image_cell)
        has_image_hint = image_label in {"image", "picture", "photo"}
        if code or has_image_hint:
            rows.append(row_idx)
    if rows:
        return rows

    # Last fallback: use rows that have any vendor/material value.
    for row_idx in range(start_row, max_row + 1):
        vendor_value = ws.cell(row_idx, vendor_col).value if vendor_col else None
        material_value = ws.cell(row_idx, material_col).value if material_col else None
        if vendor_value not in (None, "") or material_value not in (None, ""):
            rows.append(row_idx)
    return rows


def _resolve_zip_path(base_path, target):
    if not target:
        return None
    clean_target = target.replace("\\", "/")
    if clean_target.startswith("/"):
        return clean_target.lstrip("/")
    base_dir = posixpath.dirname(base_path)
    return posixpath.normpath(posixpath.join(base_dir, clean_target))


def _read_relationships(archive, rels_path):
    rels = {}
    if rels_path not in archive.namelist():
        return rels
    rel_root = ET.fromstring(archive.read(rels_path))
    ns_rel = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
    for rel in rel_root.findall("rel:Relationship", ns_rel):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rel_id and target:
            rels[rel_id] = target
    return rels


def _sheet_path_for_name(archive, sheet_name):
    workbook_path = "xl/workbook.xml"
    workbook_rels_path = "xl/_rels/workbook.xml.rels"
    if workbook_path not in archive.namelist():
        return None

    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    wb_root = ET.fromstring(archive.read(workbook_path))
    workbook_rels = _read_relationships(archive, workbook_rels_path)
    for sheet in wb_root.findall(".//main:sheets/main:sheet", ns):
        if sheet.attrib.get("name") != sheet_name:
            continue
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = workbook_rels.get(rid)
        return _resolve_zip_path(workbook_path, target)
    return None


def _extract_drawing_images_for_sheet(file_bytes, sheet_name):
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    entries = []
    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as archive:
        sheet_path = _sheet_path_for_name(archive, sheet_name)
        if not sheet_path or sheet_path not in archive.namelist():
            return entries

        sheet_root = ET.fromstring(archive.read(sheet_path))
        sheet_rels_path = "{0}/_rels/{1}.rels".format(
            posixpath.dirname(sheet_path),
            posixpath.basename(sheet_path),
        )
        sheet_rels = _read_relationships(archive, sheet_rels_path)
        drawing_rids = []
        for drawing in sheet_root.findall(".//main:drawing", ns):
            rid = drawing.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if rid:
                drawing_rids.append(rid)

        item_idx = 0
        for rid in drawing_rids:
            drawing_target = sheet_rels.get(rid)
            drawing_path = _resolve_zip_path(sheet_path, drawing_target)
            if not drawing_path or drawing_path not in archive.namelist():
                continue

            drawing_root = ET.fromstring(archive.read(drawing_path))
            drawing_rels_path = "{0}/_rels/{1}.rels".format(
                posixpath.dirname(drawing_path),
                posixpath.basename(drawing_path),
            )
            drawing_rels = _read_relationships(archive, drawing_rels_path)

            for anchor in drawing_root.findall("xdr:twoCellAnchor", ns) + drawing_root.findall(
                "xdr:oneCellAnchor", ns
            ):
                from_node = anchor.find("xdr:from", ns)
                row = None
                col = None
                if from_node is not None:
                    row_text = from_node.findtext("xdr:row", default=None, namespaces=ns)
                    col_text = from_node.findtext("xdr:col", default=None, namespaces=ns)
                    if row_text is not None and row_text.isdigit():
                        row = int(row_text) + 1
                    if col_text is not None and col_text.isdigit():
                        col = int(col_text) + 1

                blip = anchor.find(".//a:blip", ns)
                if blip is None:
                    continue
                embed = blip.attrib.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                )
                target = drawing_rels.get(embed)
                media_path = _resolve_zip_path(drawing_path, target)
                if not media_path or media_path not in archive.namelist():
                    continue

                data = archive.read(media_path)
                if not data:
                    continue
                item_idx += 1
                entries.append(
                    {
                        "row": row,
                        "col": col,
                        "ext": _normalize_ext(Path(media_path).suffix, data),
                        "data": data,
                        "source": "drawing:{0}".format(item_idx),
                    }
                )

    entries.sort(key=lambda item: (item.get("row") or 10**9, item.get("col") or 10**9, item["source"]))
    return entries


def _extract_openpyxl_images(images):
    entries = []
    for idx, img in enumerate(images, start=1):
        row, col = _anchor_row_col(img)
        try:
            image_data = img._data()
        except Exception:
            continue
        entries.append(
            {
                "row": row,
                "col": col,
                "ext": _normalize_ext(None, image_data),
                "data": image_data,
                "source": "openpyxl:{0}".format(idx),
            }
        )
    entries.sort(key=lambda item: (item.get("row") or 10**9, item.get("col") or 10**9, item["source"]))
    return entries


def _cell_ref_to_row_col(cell_ref):
    match = re.match(r"^([A-Za-z]+)(\d+)$", cell_ref or "")
    if not match:
        return None, None
    col_letters = match.group(1).upper()
    row_idx = int(match.group(2))
    col_idx = 0
    for ch in col_letters:
        col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
    return row_idx, col_idx


def _extract_dispimg_row_map(archive, sheet_path, image_col, start_row):
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    row_map = {}
    sheet_root = ET.fromstring(archive.read(sheet_path))
    for cell in sheet_root.findall(".//main:c", ns):
        cell_ref = cell.attrib.get("r")
        row_idx, col_idx = _cell_ref_to_row_col(cell_ref)
        if row_idx is None or col_idx is None:
            continue
        if col_idx != image_col or row_idx < start_row:
            continue
        formula = cell.findtext("main:f", default="", namespaces=ns) or ""
        if "DISPIMG" not in formula.upper():
            continue
        key_match = re.search(r'DISPIMG\(\s*"([^"]+)"', formula, flags=re.IGNORECASE)
        if not key_match:
            key_match = re.search(r"DISPIMG\(\s*'([^']+)'", formula, flags=re.IGNORECASE)
        if not key_match:
            continue
        key = key_match.group(1).strip()
        if key:
            row_map[row_idx] = key
    return row_map


def _find_cellimages_part_path(archive):
    workbook_rels = _read_relationships(archive, "xl/_rels/workbook.xml.rels")
    for target in workbook_rels.values():
        if "cellimage" in target.lower():
            part_path = _resolve_zip_path("xl/workbook.xml", target)
            if part_path in archive.namelist():
                return part_path

    for candidate in ("xl/cellimages.xml", "xl/cellImages.xml"):
        if candidate in archive.namelist():
            return candidate
    return None


def _extract_cellimages_by_key(archive, cellimages_path):
    ns = {
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }
    rels_path = "{0}/_rels/{1}.rels".format(
        posixpath.dirname(cellimages_path),
        posixpath.basename(cellimages_path),
    )
    rels_map = _read_relationships(archive, rels_path)
    root = ET.fromstring(archive.read(cellimages_path))
    images_by_key = {}

    for pic in root.findall(".//xdr:pic", ns):
        c_nv_pr = pic.find(".//xdr:cNvPr", ns)
        if c_nv_pr is None:
            continue
        key_name = (c_nv_pr.attrib.get("name") or "").strip()
        if not key_name:
            continue
        blip = pic.find(".//a:blip", ns)
        if blip is None:
            continue
        rel_id = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        target = rels_map.get(rel_id)
        media_path = _resolve_zip_path(cellimages_path, target)
        if not media_path or media_path not in archive.namelist():
            continue
        data = archive.read(media_path)
        if not data:
            continue
        item = {
            "data": data,
            "ext": _normalize_ext(Path(media_path).suffix, data),
            "source": "cellimages:{0}".format(key_name),
        }
        images_by_key[key_name] = item
        images_by_key[key_name.upper()] = item
    return images_by_key


def _extract_dispimg_entries(file_bytes, sheet_name, image_col, start_row):
    entries = []
    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as archive:
        sheet_path = _sheet_path_for_name(archive, sheet_name)
        if not sheet_path or sheet_path not in archive.namelist():
            return entries
        row_key_map = _extract_dispimg_row_map(archive, sheet_path, image_col, start_row)
        if not row_key_map:
            return entries

        cellimages_path = _find_cellimages_part_path(archive)
        if not cellimages_path:
            return entries
        images_by_key = _extract_cellimages_by_key(archive, cellimages_path)
        if not images_by_key:
            return entries

        for row_idx in sorted(row_key_map.keys()):
            key = row_key_map[row_idx]
            image_item = images_by_key.get(key) or images_by_key.get(key.upper())
            if not image_item:
                continue
            entries.append(
                {
                    "row": row_idx,
                    "col": image_col,
                    "ext": image_item["ext"],
                    "data": image_item["data"],
                    "source": "dispimg:{0}".format(key),
                }
            )
    return entries


def _collect_target_rows(ws, start_row, vendor_col, material_col):
    rows = []
    max_row = ws.max_row or start_row
    for row_idx in range(start_row, max_row + 1):
        code, _ = _row_code(ws, row_idx, vendor_col, material_col)
        if code:
            rows.append(row_idx)
    return rows


def _assign_entries_to_rows(target_rows, entries):
    if not target_rows or not entries:
        return []

    row_entries = [entry for entry in entries if entry.get("row") is not None]
    no_row_entries = [entry for entry in entries if entry.get("row") is None]

    if not row_entries:
        mapped = []
        for idx, row_idx in enumerate(target_rows):
            mapped.append((row_idx, no_row_entries[idx % len(no_row_entries)]))
        return mapped

    row_entries.sort(key=lambda item: item["row"])
    mapped = []
    pointer = 0
    current = row_entries[0]

    for row_idx in sorted(target_rows):
        while pointer + 1 < len(row_entries) and row_entries[pointer + 1]["row"] <= row_idx:
            pointer += 1
            current = row_entries[pointer]
        mapped.append((row_idx, current))
    return mapped


def _maybe_upscale_image(data, ext, scale_factor=3):
    if Image is None:
        return data, ext, False

    try:
        with Image.open(io.BytesIO(data)) as img:
            # Upscale only when image is visually small.
            if max(img.size) >= 300:
                return data, ext, False

            resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.BICUBIC)
            resized = img.resize((img.width * scale_factor, img.height * scale_factor), resampling)
            output = io.BytesIO()
            # Save as PNG to avoid additional quality loss from JPEG re-encoding.
            resized.save(output, format="PNG")
            return output.getvalue(), "png", True
    except Exception:
        return data, ext, False


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
    file_bytes, _filename, error = _get_uploaded_file()
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
    file_bytes, original_filename, error = _get_uploaded_file()
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
    layout = _detect_layout(ws)
    image_col = layout["image_col"]
    vendor_col = layout["vendor_col"]
    material_col = layout["material_col"]
    start_row = layout["start_row"]
    target_rows = _collect_target_rows(ws, start_row, vendor_col, material_col)

    openpyxl_images = list(getattr(ws, "_images", []) or [])
    openpyxl_entries = _extract_openpyxl_images(openpyxl_images)
    drawing_entries = _extract_drawing_images_for_sheet(file_bytes, sheet_name)
    dispimg_entries = _extract_dispimg_entries(file_bytes, sheet_name, image_col, start_row)
    media_images = _extract_media_images(file_bytes)

    extracted_count = 0
    skipped_count = 0
    skipped_reasons = []
    seen_filenames = set()
    extraction_mode = "none"
    upscaled_count = 0

    excel_name_no_ext = Path(original_filename).stem if original_filename else "Excel_Images"
    root_folder = _safe_folder_name(excel_name_no_ext)
    download_name = "{0}.zip".format(root_folder)

    temp_file = tempfile.NamedTemporaryFile(
        prefix="dataiku_excel_images_",
        suffix=".zip",
        delete=False,
    )
    zip_path = temp_file.name
    temp_file.close()

    try:
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            source_entries = []
            if dispimg_entries:
                source_entries = dispimg_entries
                extraction_mode = "dispimg_cellimages"
            elif drawing_entries:
                source_entries = drawing_entries
                extraction_mode = "drawing_anchor"
            elif openpyxl_entries:
                source_entries = openpyxl_entries
                extraction_mode = "openpyxl_anchor"
            elif media_images:
                source_entries = [
                    {
                        "row": None,
                        "col": image_col,
                        "ext": media["ext"],
                        "data": media["data"],
                        "source": media["source"],
                    }
                    for media in media_images
                ]
                extraction_mode = "xlsx_media_fallback"

            image_cache = {}

            def _prepared_image(entry):
                cache_key = entry.get("source") or id(entry)
                if cache_key in image_cache:
                    return image_cache[cache_key]
                new_data, new_ext, did_upscale = _maybe_upscale_image(entry["data"], entry["ext"], scale_factor=3)
                image_cache[cache_key] = (new_data, new_ext, did_upscale)
                return image_cache[cache_key]

            # Primary behavior requested: one output image per vendor/material row.
            if source_entries and target_rows:
                mapped_rows = _assign_entries_to_rows(target_rows, source_entries)
                for row_idx, entry in mapped_rows:
                    safe_code, code_source = _row_code(ws, row_idx, vendor_col, material_col)
                    if not safe_code:
                        safe_code = "Row_{0}".format(row_idx)
                    if code_source == "material":
                        skipped_reasons.append(
                            "Row {0}: vendor missing, used material fallback MAT_.".format(row_idx)
                        )

                    out_data, out_ext, did_upscale = _prepared_image(entry)
                    if did_upscale:
                        upscaled_count += 1

                    filename = _next_unique_filename(safe_code, out_ext, seen_filenames)
                    zip_file.writestr("{0}/{1}".format(root_folder, filename), out_data)
                    extracted_count += 1

                if len(mapped_rows) != len(target_rows):
                    skipped_count += abs(len(target_rows) - len(mapped_rows))
                    skipped_reasons.append(
                        "Target rows ({0}) and mapped images ({1}) mismatch.".format(
                            len(target_rows), len(mapped_rows)
                        )
                    )

            # If no vendor/material rows were found, still export discovered images.
            elif source_entries:
                for idx, entry in enumerate(source_entries, start=1):
                    safe_code = "Image_{0}".format(idx)
                    out_data, out_ext, did_upscale = _prepared_image(entry)
                    if did_upscale:
                        upscaled_count += 1
                    filename = _next_unique_filename(safe_code, out_ext, seen_filenames)
                    zip_file.writestr("{0}/{1}".format(root_folder, filename), out_data)
                    extracted_count += 1

            if target_rows and extracted_count and extracted_count != len(target_rows):
                skipped_count += abs(len(target_rows) - extracted_count)
                skipped_reasons.append(
                    "Final count check failed: extracted {0}, vendor/material rows {1}.".format(
                        extracted_count, len(target_rows)
                    )
                )

            summary_lines = [
                "Excel Image Extraction Summary",
                "==============================",
                "Generated at: {0}Z".format(datetime.utcnow().isoformat()),
                "Sheet: {0}".format(sheet_name),
                "Workbook file: {0}".format(original_filename or ""),
                "Output root folder: {0}".format(root_folder),
                "Extraction mode: {0}".format(extraction_mode),
                "Detected image column: {0}".format(image_col),
                "Detected vendor column: {0}".format(vendor_col),
                "Detected material column: {0}".format(material_col or "none"),
                "Data start row: {0}".format(start_row),
                "Vendor/material target rows: {0}".format(len(target_rows)),
                "DISPIMG/cellimages entries: {0}".format(len(dispimg_entries)),
                "Openpyxl anchored images: {0}".format(len(openpyxl_entries)),
                "Drawing anchored images: {0}".format(len(drawing_entries)),
                "XLSX media items: {0}".format(len(media_images)),
                "Upscaled images (3x): {0}".format(upscaled_count),
                "Extracted images: {0}".format(extracted_count),
                "Skipped images: {0}".format(skipped_count),
                "",
                "Rules:",
                "- Row mapping starts after detected header rows.",
                "- One output file per vendor/material row when rows are detected.",
                "- Preferred name is Vendor Material from detected vendor column.",
                "- If Vendor is empty and Original Material exists, file uses MAT_<material>.",
                "",
            ]
            if skipped_reasons:
                summary_lines.append("Skipped details:")
                summary_lines.extend("- {0}".format(reason) for reason in skipped_reasons)

            zip_file.writestr("{0}/summary.txt".format(root_folder), "\n".join(summary_lines))
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

    return _send_zip_response(zip_path, download_name)
