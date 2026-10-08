import csv
import io
import os
import re
import zipfile
from pathlib import Path

import fitz  # PyMuPDF
import pytesseract
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)
from PIL import Image, UnidentifiedImageError
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

ELIGIBLE_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".pptx",
    ".png",
    ".jpg",
    ".jpeg",
    ".txt",
    ".xlsx",
    ".csv",
    ".zip",
}

BLOCKED_EXTENSIONS = {
    ".exe", ".dll", ".msi",
    ".bat", ".cmd", ".ps1", ".sh", ".js", ".vbs", ".jar", ".scr",
    ".docm", ".xlsm", ".pptm",
    ".doc", ".ppt", ".xls",
    ".rar", ".7z",
}

# Limits help reduce obvious resource-exhaustion / ZIP-bomb risks.
MAX_ZIP_DEPTH = 5
MAX_ZIP_FILES = 500
MAX_ZIP_TOTAL_UNCOMPRESSED = 100 * 1024 * 1024  # 100 MB
MAX_ZIP_ENTRY_SIZE = 25 * 1024 * 1024            # 25 MB
MAX_INPUT_FILE_SIZE = 100 * 1024 * 1024          # 100 MB

ALLOWED_FORMAT_MESSAGE = (
    ".pdf, .docx, .pptx, .png, .jpg, .jpeg, "
    ".txt, .xlsx, .csv, .zip"
)


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------

def normalize_text(text):
    """Apply light normalization without destroying meaningful content."""
    if text is None:
        return ""

    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Remove control characters except tab and newline.
    text = "".join(
        ch for ch in text
        if ch in ("\n", "\t") or ch >= " "
    )

    # Remove trailing spaces while preserving line structure.
    lines = [line.rstrip() for line in text.split("\n")]

    # Collapse runs of more than two blank lines.
    normalized = []
    blank_count = 0

    for line in lines:
        if line.strip() == "":
            blank_count += 1
            if blank_count <= 2:
                normalized.append("")
        else:
            blank_count = 0
            normalized.append(line)

    return "\n".join(normalized).strip()


def make_item(file_name, file_type, location_type, location, text,
              extraction_method, ocr_used=False, source_path=None):
    """Create the common structure passed to the next module."""
    item = {
        "file_name": file_name,
        "file_type": file_type,
        "location_type": location_type,
        "location": location,
        "text": normalize_text(text),
        "extraction_method": extraction_method,
        "ocr_used": bool(ocr_used),
    }

    if source_path:
        item["source_path"] = source_path

    return item


def success_result(file_name, file_type, content, **extra):
    result = {
        "success": True,
        "file_name": file_name,
        "file_type": file_type,
        "content": content,
    }
    result.update(extra)
    return result


def error_result(file_name, error, file_type=None):
    result = {
        "success": False,
        "file_name": file_name,
        "error": error,
    }
    if file_type:
        result["file_type"] = file_type
    return result


def extension_of(file_name):
    return Path(file_name).suffix.lower()


# ---------------------------------------------------------------------------
# File type validation
# ---------------------------------------------------------------------------

def detect_actual_file_type(data):
    """
    Detect the actual type from file signatures / container structure.

    Returns one of:
    pdf, docx, pptx, xlsx, zip, png, jpg, jpeg, txt, csv, unknown
    """
    if data.startswith(b"%PDF-"):
        return "pdf"

    try:
        with Image.open(io.BytesIO(data)) as img:
            image_format = img.format

        if image_format == "PNG":
            return "png"

        if image_format == "JPEG":
            return "jpeg"

    except (UnidentifiedImageError, OSError):
        pass

    # ZIP / OOXML containers.
    if data.startswith(b"PK\x03\x04") or data.startswith(b"PK\x05\x06"):
        try:
            with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                names = set(zf.namelist())

                if "[Content_Types].xml" in names:
                    name_text = "\n".join(names)

                    if "word/document.xml" in names:
                        return "docx"

                    if "ppt/presentation.xml" in names:
                        return "pptx"

                    if "xl/workbook.xml" in names:
                        return "xlsx"

                return "zip"
        except (zipfile.BadZipFile, OSError):
            return "zip"

    # Text-based formats cannot be identified perfectly from magic bytes.
    # They are validated further by decoding and CSV parsing.
    try:
        decoded = data.decode("utf-8-sig")
        if is_probably_text(decoded):
            return "text"
    except UnicodeDecodeError:
        pass

    return "unknown"


def is_probably_text(text):
    """Reject text containing a large amount of suspicious binary data."""
    if not text:
        return True

    sample = text[:10000]
    bad = 0

    for ch in sample:
        if ch == "\x00":
            bad += 1
        elif ord(ch) < 32 and ch not in ("\n", "\r", "\t", "\f"):
            bad += 1

    return (bad / max(len(sample), 1)) < 0.01


def validate_extension_and_type(file_name, data):
    """
    Validate both the claimed extension and the actual content.

    TXT/CSV are text formats and cannot be distinguished perfectly by magic
    bytes, so they receive additional text validation.
    """
    ext = extension_of(file_name)

    if ext in BLOCKED_EXTENSIONS:
        return False, (
            f"Blocked file type: {ext}. "
            "This file format is not allowed."
        )

    if ext not in ELIGIBLE_EXTENSIONS:
        return False, (
            f"Unsupported file type: {ext or '[no extension]'}\n\n"
            f"Allowed formats: {ALLOWED_FORMAT_MESSAGE}"
        )

    actual = detect_actual_file_type(data)

    expected_map = {
        ".pdf": {"pdf"},
        ".docx": {"docx"},
        ".pptx": {"pptx"},
        ".xlsx": {"xlsx"},
        ".zip": {"zip"},
        ".png": {"png"},
        ".jpg": {"jpeg", "jpg"},
        ".jpeg": {"jpeg", "jpg"},
        ".txt": {"text"},
        ".csv": {"text"},
    }

    if ext not in expected_map:
        return False, "Unable to validate the file type."

    if actual not in expected_map[ext]:
        return False, (
            f"File type mismatch: the extension '{ext}' does not match "
            f"the detected file content ('{actual}')."
        )

    if ext == ".txt":
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = data.decode("utf-16")
            except UnicodeDecodeError:
                return False, "The TXT file could not be decoded as text."

        if not is_probably_text(text):
            return False, "The TXT file appears to contain binary data."

    if ext == ".csv":
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = data.decode("utf-16")
            except UnicodeDecodeError:
                return False, "The CSV file could not be decoded as text."

        if not is_probably_text(text):
            return False, "The CSV file appears to contain binary data."

    return True, None


# ---------------------------------------------------------------------------
# OCR helpers
# ---------------------------------------------------------------------------

def ocr_image(image):
    """Run Tesseract OCR on a PIL image."""
    try:
        text = pytesseract.image_to_string(image)
    except Exception as exc:
        raise RuntimeError(f"OCR failed: {exc}") from exc

    if not normalize_text(text):
        raise RuntimeError("OCR completed but produced no usable text.")

    return text


def ocr_pdf_page(page):
    """Render one PDF page and run OCR."""
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        image = Image.open(io.BytesIO(pix.tobytes("png")))
        return ocr_image(image)
    except Exception as exc:
        raise RuntimeError(f"PDF page OCR failed: {exc}") from exc


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def extract_pdf(data, file_name, source_path=None):
    content = []
    ocr_pages = 0

    try:
        document = fitz.open(stream=data, filetype="pdf")
    except Exception as exc:
        raise RuntimeError(f"Could not open PDF: {exc}") from exc

    try:
        if document.needs_pass:
            raise RuntimeError("Password-protected PDF files are not supported.")

        if document.page_count == 0:
            raise RuntimeError("The PDF contains no pages.")

        for page_number in range(document.page_count):
            page = document.load_page(page_number)
            text = normalize_text(page.get_text("text"))

            if text:
                content.append(
                    make_item(
                        file_name=file_name,
                        file_type="pdf",
                        location_type="page",
                        location=page_number + 1,
                        text=text,
                        extraction_method="pymupdf",
                        ocr_used=False,
                        source_path=source_path,
                    )
                )
            else:
                try:
                    ocr_text = ocr_pdf_page(page)
                    ocr_pages += 1

                    content.append(
                        make_item(
                            file_name=file_name,
                            file_type="pdf",
                            location_type="page",
                            location=page_number + 1,
                            text=ocr_text,
                            extraction_method="tesseract",
                            ocr_used=True,
                            source_path=source_path,
                        )
                    )
                except RuntimeError as exc:
                    raise RuntimeError(
                        f"OCR failed on PDF page {page_number + 1}: {exc}"
                    ) from exc
    finally:
        document.close()

    if not content:
        raise RuntimeError("No usable text could be extracted from the PDF.")

    return content, ocr_pages


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------

def extract_docx(data, file_name, source_path=None):
    try:
        document = Document(io.BytesIO(data))
    except Exception as exc:
        raise RuntimeError(f"Could not open DOCX: {exc}") from exc

    content = []
    location = 0

    # Paragraphs, including headings and list paragraphs.
    for paragraph in document.paragraphs:
        text = normalize_text(paragraph.text)
        if text:
            location += 1
            content.append(
                make_item(
                    file_name,
                    "docx",
                    "paragraph",
                    location,
                    text,
                    "python-docx",
                    False,
                    source_path,
                )
            )

    # Tables.
    for table_index, table in enumerate(document.tables, start=1):
        rows = []

        for row in table.rows:
            values = [normalize_text(cell.text) for cell in row.cells]
            rows.append(" | ".join(values))

        table_text = "\n".join(rows).strip()

        if table_text:
            content.append(
                make_item(
                    file_name,
                    "docx",
                    "table",
                    table_index,
                    table_text,
                    "python-docx",
                    False,
                    source_path,
                )
            )

    if not content:
        raise RuntimeError("The DOCX contains no usable text.")

    return content


# ---------------------------------------------------------------------------
# PPTX
# ---------------------------------------------------------------------------

def extract_pptx(data, file_name, source_path=None):
    try:
        presentation = Presentation(io.BytesIO(data))
    except Exception as exc:
        raise RuntimeError(f"Could not open PPTX: {exc}") from exc

    content = []

    for slide_number, slide in enumerate(presentation.slides, start=1):
        parts = []

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = normalize_text(shape.text)
                if text:
                    parts.append(text)

            # Table cells are not always represented adequately by shape.text.
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    values = [
                        normalize_text(cell.text)
                        for cell in row.cells
                    ]
                    row_text = " | ".join(values)
                    if row_text:
                        parts.append(row_text)

        slide_text = "\n".join(parts)

        if slide_text:
            content.append(
                make_item(
                    file_name,
                    "pptx",
                    "slide",
                    slide_number,
                    slide_text,
                    "python-pptx",
                    False,
                    source_path,
                )
            )

    if not content:
        raise RuntimeError("The PPTX contains no usable text.")

    return content


# ---------------------------------------------------------------------------
# Images
# ---------------------------------------------------------------------------

def extract_image(data, file_name, file_type, source_path=None):
    try:
        image = Image.open(io.BytesIO(data))
        image.verify()
    except (UnidentifiedImageError, OSError) as exc:
        raise RuntimeError(f"Invalid image file: {exc}") from exc

    try:
        image = Image.open(io.BytesIO(data))
        text = ocr_image(image)
    except Exception as exc:
        raise RuntimeError(f"Image OCR failed: {exc}") from exc

    return [
        make_item(
            file_name,
            file_type,
            "image",
            1,
            text,
            "tesseract",
            True,
            source_path,
        )
    ]


# ---------------------------------------------------------------------------
# TXT
# ---------------------------------------------------------------------------

def decode_text_file(data):
    encodings = ("utf-8-sig", "utf-16", "cp1252")

    for encoding in encodings:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise RuntimeError("The text file could not be decoded.")


def extract_txt(data, file_name, source_path=None):
    text = decode_text_file(data)
    text = normalize_text(text)

    if not text:
        raise RuntimeError("The TXT file is empty.")

    return [
        make_item(
            file_name,
            "txt",
            "text",
            1,
            text,
            "plain_text",
            False,
            source_path,
        )
    ]


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------

def extract_xlsx(data, file_name, source_path=None):
    try:
        workbook = load_workbook(
            filename=io.BytesIO(data),
            read_only=True,
            data_only=True,
        )
    except Exception as exc:
        raise RuntimeError(f"Could not open XLSX: {exc}") from exc

    content = []

    try:
        for worksheet in workbook.worksheets:
            rows = []

            for row in worksheet.iter_rows(values_only=True):
                values = [
                    "" if value is None else str(value)
                    for value in row
                ]

                # Keep rows that contain at least one meaningful cell.
                if any(value.strip() for value in values):
                    rows.append(" | ".join(values))

            worksheet_text = "\n".join(rows).strip()

            if worksheet_text:
                content.append(
                    make_item(
                        file_name,
                        "xlsx",
                        "worksheet",
                        worksheet.title,
                        worksheet_text,
                        "openpyxl",
                        False,
                        source_path,
                    )
                )
    finally:
        workbook.close()

    if not content:
        raise RuntimeError("The XLSX contains no usable cell values.")

    return content


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

def extract_csv(data, file_name, source_path=None):
    text = decode_text_file(data)

    try:
        reader = csv.reader(io.StringIO(text))
        rows = []

        for row in reader:
            rows.append(" | ".join(cell for cell in row))

    except csv.Error as exc:
        raise RuntimeError(f"Invalid CSV file: {exc}") from exc

    csv_text = normalize_text("\n".join(rows))

    if not csv_text:
        raise RuntimeError("The CSV file is empty.")

    return [
        make_item(
            file_name,
            "csv",
            "csv",
            1,
            csv_text,
            "csv",
            False,
            source_path,
        )
    ]


# ---------------------------------------------------------------------------
# ZIP security and extraction
# ---------------------------------------------------------------------------

def validate_zip_entries(zf, depth):
    """Validate every non-directory ZIP entry before processing."""
    if depth > MAX_ZIP_DEPTH:
        raise RuntimeError(
            f"Nested ZIP depth exceeds the allowed limit of {MAX_ZIP_DEPTH}."
        )

    entries = [
        info for info in zf.infolist()
        if not info.is_dir()
    ]

    if len(entries) > MAX_ZIP_FILES:
        raise RuntimeError(
            f"ZIP contains too many files. Maximum allowed is {MAX_ZIP_FILES}."
        )

    total_size = 0

    for info in entries:
        if info.file_size > MAX_ZIP_ENTRY_SIZE:
            raise RuntimeError(
                f"ZIP entry '{info.filename}' is too large."
            )

        total_size += info.file_size

        if total_size > MAX_ZIP_TOTAL_UNCOMPRESSED:
            raise RuntimeError(
                "ZIP uncompressed content exceeds the allowed size limit."
            )

        # Reject absolute paths and path traversal.
        normalized = info.filename.replace("\\", "/")

        if normalized.startswith("/") or re.match(r"^[A-Za-z]:/", normalized):
            raise RuntimeError(
                f"Unsafe ZIP path detected: {info.filename}"
            )

        parts = Path(normalized).parts
        if ".." in parts:
            raise RuntimeError(
                f"Unsafe ZIP path detected: {info.filename}"
            )

        # Encrypted ZIP entries have flag bit 0 set.
        if info.flag_bits & 0x1:
            raise RuntimeError(
                f"Password-protected ZIP entry is not supported: "
                f"{info.filename}"
            )


def extract_zip(data, file_name, source_path=None, depth=0):
    try:
        zf = zipfile.ZipFile(io.BytesIO(data), "r")
    except zipfile.BadZipFile as exc:
        raise RuntimeError(f"Corrupted ZIP file: {exc}") from exc

    content = []
    processed_files = []

    try:
        validate_zip_entries(zf, depth)

        entries = [
            info for info in zf.infolist()
            if not info.is_dir()
        ]

        for info in entries:
            entry_name = info.filename
            entry_ext = extension_of(entry_name)

            if entry_ext in BLOCKED_EXTENSIONS:
                raise RuntimeError(
                    f"Blocked file type detected inside ZIP: {entry_name}"
                )

            if entry_ext not in ELIGIBLE_EXTENSIONS:
                raise RuntimeError(
                    f"Unsupported file type inside ZIP: {entry_name}. "
                    f"Allowed formats: {ALLOWED_FORMAT_MESSAGE}"
                )

            try:
                entry_data = zf.read(info)
            except Exception as exc:
                raise RuntimeError(
                    f"Could not read ZIP entry '{entry_name}': {exc}"
                ) from exc

            valid, error = validate_extension_and_type(
                entry_name,
                entry_data,
            )

            if not valid:
                raise RuntimeError(
                    f"Invalid ZIP entry '{entry_name}': {error}"
                )

            # Nested ZIPs are recursively inspected.
            if entry_ext == ".zip":
                nested_content = extract_zip(
                    entry_data,
                    Path(entry_name).name,
                    source_path=entry_name,
                    depth=depth + 1,
                )

                for item in nested_content:
                    item["source_path"] = (
                        f"{entry_name}/{item.get('source_path', item['file_name'])}"
                    )
                    content.append(item)

                processed_files.append(entry_name)
                continue

            entry_result = extract_bytes_by_extension(
                entry_data,
                Path(entry_name).name,
                entry_ext,
                source_path=entry_name,
            )

            content.extend(entry_result)
            processed_files.append(entry_name)

    finally:
        zf.close()

    if not content:
        raise RuntimeError("The ZIP contains no usable eligible files.")

    return content


# ---------------------------------------------------------------------------
# Dispatcher
# ---------------------------------------------------------------------------

def extract_bytes_by_extension(data, file_name, ext, source_path=None):
    if ext == ".pdf":
        content, _ = extract_pdf(data, file_name, source_path)
        return content

    if ext == ".docx":
        return extract_docx(data, file_name, source_path)

    if ext == ".pptx":
        return extract_pptx(data, file_name, source_path)

    if ext in {".png", ".jpg", ".jpeg"}:
        file_type = "jpeg" if ext in {".jpg", ".jpeg"} else "png"
        return extract_image(
            data,
            file_name,
            file_type,
            source_path,
        )

    if ext == ".txt":
        return extract_txt(data, file_name, source_path)

    if ext == ".xlsx":
        return extract_xlsx(data, file_name, source_path)

    if ext == ".csv":
        return extract_csv(data, file_name, source_path)

    if ext == ".zip":
        return extract_zip(
            data,
            file_name,
            source_path=source_path,
            depth=0,
        )

    raise RuntimeError(f"Unsupported file type: {ext}")


# ---------------------------------------------------------------------------
# Main public function
# ---------------------------------------------------------------------------

def extract_document(file_path):
    """
    Main entry point for the PII-Shield Document Extraction + OCR module.

    Returns:
        {
            "success": True/False,
            "file_name": "...",
            "file_type": "...",
            "content": [...],
            ...
        }
    """
    path = Path(file_path)

    if not path.exists():
        return error_result(
            path.name,
            "File does not exist."
        )

    if not path.is_file():
        return error_result(
            path.name,
            "The supplied path is not a file."
        )

    try:
        file_size = path.stat().st_size
    except OSError as exc:
        return error_result(
            path.name,
            f"Could not read file information: {exc}"
        )

    if file_size == 0:
        return error_result(
            path.name,
            "The document is empty."
        )

    if file_size > MAX_INPUT_FILE_SIZE:
        return error_result(
            path.name,
            "The input file is too large for this module."
        )

    file_name = path.name
    ext = extension_of(file_name)

    if ext in BLOCKED_EXTENSIONS:
        return error_result(
            file_name,
            f"Blocked file type: {ext}. "
            "This file format is not allowed."
        )

    if ext not in ELIGIBLE_EXTENSIONS:
        return error_result(
            file_name,
            f"Unsupported file type: {ext or '[no extension]'}\n\n"
            f"Allowed formats: {ALLOWED_FORMAT_MESSAGE}"
        )

    try:
        data = path.read_bytes()
    except OSError as exc:
        return error_result(
            file_name,
            f"Could not read the file: {exc}"
        )

    valid, error = validate_extension_and_type(file_name, data)

    if not valid:
        return error_result(file_name, error)

    try:
        if ext == ".pdf":
            content, ocr_pages = extract_pdf(
                data,
                file_name,
                source_path=file_name,
            )

            return success_result(
                file_name,
                "pdf",
                content,
                pages_processed=len(content),
                pages_using_ocr=ocr_pages,
            )

        if ext == ".zip":
            content = extract_zip(
                data,
                file_name,
                source_path=file_name,
                depth=0,
            )

            return success_result(
                file_name,
                "zip",
                content,
                files_processed=len({
                    item.get("source_path", item["file_name"])
                    for item in content
                }),
            )

        content = extract_bytes_by_extension(
            data,
            file_name,
            ext,
            source_path=file_name,
        )

        return success_result(
            file_name,
            ext.lstrip("."),
            content,
        )

    except Exception as exc:
        return error_result(
            file_name,
            str(exc),
            file_type=ext.lstrip("."),
        )


# ---------------------------------------------------------------------------
# Direct testing
# ---------------------------------------------------------------------------

def print_result(result):
    print("\n" + "=" * 60)

    if not result.get("success"):
        print("File rejected / extraction failed.")
        print(f"File: {result.get('file_name', 'Unknown')}")
        print(f"Reason: {result.get('error', 'Unknown error')}")
        print("=" * 60)
        return

    print(f"File: {result['file_name']}")
    print(f"Type: {result['file_type'].upper()}")
    print("Status: Success")

    if result["file_type"] == "pdf":
        print(f"Pages processed: {result.get('pages_processed', 0)}")
        print(f"Pages using OCR: {result.get('pages_using_ocr', 0)}")

    if result["file_type"] == "zip":
        print(f"Files processed: {result.get('files_processed', 0)}")

    print("-" * 60)

    # Direct testing should be useful, but avoid dumping huge documents.
    max_items = 20
    items = result.get("content", [])

    for index, item in enumerate(items[:max_items], start=1):
        print(
            f"\n[{index}] "
            f"{item['location_type'].upper()}: {item['location']}"
        )
        print(f"Method: {item['extraction_method']}")
        print(f"OCR: {'Yes' if item['ocr_used'] else 'No'}")

        if item.get("source_path"):
            print(f"Source: {item['source_path']}")

        text = item.get("text", "")
        preview = text[:3000]

        print("\n" + preview)

        if len(text) > len(preview):
            print("\n[Text preview truncated]")

    if len(items) > max_items:
        print(
            f"\nOnly the first {max_items} items are shown "
            f"out of {len(items)} extracted items."
        )

    print("=" * 60)


if __name__ == "__main__":
    print("PII-Shield - Document Extraction + OCR")
    print("Supported: PDF, DOCX, PPTX, PNG, JPG, JPEG, TXT, XLSX, CSV, ZIP")
    print()

    file_path = input("Enter document path: ").strip().strip('"')

    if not file_path:
        print("No document path was provided.")
    else:
        result = extract_document(file_path)
        print_result(result)
