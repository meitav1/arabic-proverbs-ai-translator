
import json
import random
import re
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from translate_and_merge import (
    translate_scanned_pdf,
    read_instructions,
    PDF_FOLDER,
    INSTRUCTIONS_FILE,
)

# ========== SETTINGS ==========

MAX_WORKERS = 15
BATCH_SIZE = 150
WORD_UPDATE_EVERY = 10
MAX_RETRIES = 4

PDF_DIR = Path(PDF_FOLDER)
OUTPUT_DIR = Path("translated_pages")
OLD_DIR = Path("parallel_results")
REPORT_DIR = Path("translation_reports")
WORD_FILE = Path("full_book_translation.docx")

# ==============================

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
REPORT_DIR.mkdir(parents=True, exist_ok=True)

print_lock = threading.Lock()
stop_event = threading.Event()


def page_number(path):
    match = re.search(r"(\d+)$", path.stem)
    if not match:
        raise ValueError(f"Missing page number: {path}")
    return int(match.group(1))


def valid_file(path):
    return path.exists() and path.stat().st_size > 0


def import_previous_results(pdf_files):
    imported = 0

    for pdf in pdf_files:
        destination = OUTPUT_DIR / (pdf.stem + ".txt")

        if valid_file(destination):
            continue

        candidates = list(
            OLD_DIR.glob(f"workers_*/{pdf.stem}.txt")
        )

        # Include the earlier single-page measurement.
        candidates.append(
            Path("measurement_results") /
            (pdf.stem + ".txt")
        )

        for source in candidates:
            if valid_file(source):
                shutil.copy2(source, destination)
                imported += 1
                break

    return imported



def create_word(pdf_files):
    """
    Build the Word document from saved translations,
    preserving original page order and special tables.
    """
    doc = Document()
    count = 0

    def add_regular_line(line):
        paragraph = doc.add_paragraph(line)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.paragraph_format.rtl = True

    def add_special_table(rows, page):
        if not rows:
            raise ValueError(f"Empty table on page {page}")

        table = doc.add_table(rows=0, cols=2)
        table.autofit = True

        if page == 614:
            table.style = "Table Grid"

        for right_text, left_text in rows:
            cells = table.add_row().cells

            for cell, value in zip(
                cells, (right_text, left_text)
            ):
                cell.text = value

                for paragraph in cell.paragraphs:
                    paragraph.alignment = (
                        WD_ALIGN_PARAGRAPH.RIGHT
                    )
                    paragraph.paragraph_format.rtl = True

    for pdf in pdf_files:
        txt_path = OUTPUT_DIR / (pdf.stem + ".txt")

        if not valid_file(txt_path):
            continue

        content = txt_path.read_text(encoding="utf-8")

        if not content.strip():
            continue

        page = page_number(pdf)
        lines = content.splitlines()
        inside_table = False
        table_rows = []

        # Validate table formatting before modifying the document
        if page in (385, 614):
            if (
                content.count("[[TABLE_START]]") != 1
                or content.count("[[TABLE_END]]") != 1
            ):
                raise ValueError(
                    f"Missing or duplicated table markers: {pdf.name}"
                )

        if count > 0:
            doc.add_page_break()

        for line in lines:
            stripped = line.strip()

            if (
                page in (385, 614)
                and stripped == "[[TABLE_START]]"
            ):
                if inside_table:
                    raise ValueError(
                        f"Nested table in {pdf.name}"
                    )

                inside_table = True
                table_rows = []
                continue

            if (
                page in (385, 614)
                and stripped == "[[TABLE_END]]"
            ):
                if not inside_table:
                    raise ValueError(
                        f"Unexpected table end in {pdf.name}"
                    )

                add_special_table(table_rows, page)
                inside_table = False
                table_rows = []
                continue

            if inside_table:
                if line.count("|||") != 1:
                    raise ValueError(
                        f"Invalid table row in {pdf.name}"
                    )

                right, left = line.split("|||", 1)
                table_rows.append(
                    (right.strip(), left.strip())
                )
            else:
                add_regular_line(line)

        if inside_table:
            raise ValueError(
                f"Unclosed table in {pdf.name}"
            )

        count += 1

    if count == 0:
        raise ValueError(
            "No saved translations found; Word not replaced"
        )

    temp_file = WORD_FILE.with_name(
        WORD_FILE.stem + ".tmp.docx"
    )

    doc.save(temp_file)
    temp_file.replace(WORD_FILE)

    return count



def error_code(error):
    code = getattr(error, "code", None)
    if code is not None:
        return str(code)

    match = re.search(
        r"\b(402|429|500|502|503|504)\b",
        str(error)
    )

    return match.group(1) if match else None


def is_payment_error(error):
    message = str(error).lower()
    return (
        error_code(error) == "402"
        or "prepayment credits are depleted" in message
        or "credit balance" in message
    )


def is_retryable(error):
    if is_payment_error(error):
        return False

    code = error_code(error)

    if code in ("429", "500", "502", "503", "504"):
        return True

    message = str(error).lower()

    return any(term in message for term in (
        "timeout",
        "timed out",
        "connection reset",
        "connection error",
        "rate limit",
    ))


def process_page(pdf, instructions):
    output_path = OUTPUT_DIR / (pdf.stem + ".txt")

    if valid_file(output_path):
        return {
            "file": pdf.name,
            "status": "skipped",
        }

    if stop_event.is_set():
        return {
            "file": pdf.name,
            "status": "stopped",
        }

    start = time.perf_counter()

    for attempt in range(1, MAX_RETRIES + 1):
        if stop_event.is_set():
            return {
                "file": pdf.name,
                "status": "stopped",
            }

        try:
            # ORIGINAL translation function.
            # Gemini settings and thinking unchanged.
            translation = translate_scanned_pdf(
                pdf, instructions
            )

            if not translation or not translation.strip():
                raise ValueError("Empty translation")

            temporary = output_path.with_suffix(".tmp")

            temporary.write_text(
                translation,
                encoding="utf-8"
            )
            temporary.replace(output_path)

            elapsed = round(
                time.perf_counter() - start, 2
            )

            with print_lock:
                print(
                    f"[SUCCESS] {pdf.name} "
                    f"- {elapsed}s"
                )

            return {
                "file": pdf.name,
                "status": "success",
                "seconds": elapsed,
                "attempts": attempt,
            }

        except Exception as error:
            if is_payment_error(error):
                stop_event.set()

                with print_lock:
                    print(
                        "[STOP] Credits depleted. "
                        "No new pages will be started."
                    )

                return {
                    "file": pdf.name,
                    "status": "failed",
                    "error": str(error),
                }

            if (
                attempt == MAX_RETRIES
                or not is_retryable(error)
            ):
                with print_lock:
                    print(
                        f"[ERROR] {pdf.name}: {error}"
                    )

                return {
                    "file": pdf.name,
                    "status": "failed",
                    "error": str(error),
                }

            delay = min(
                60,
                2 ** attempt + random.uniform(0, 2)
            )

            with print_lock:
                print(
                    f"[RETRY] {pdf.name}, "
                    f"waiting {delay:.1f}s"
                )

            if stop_event.wait(delay):
                return {
                    "file": pdf.name,
                    "status": "stopped",
                }


def save_report(results, batch, seconds, word_count):
    report = {
        "date": datetime.now().isoformat(),
        "workers": MAX_WORKERS,
        "batch_size": len(batch),
        "seconds": round(seconds, 2),
        "word_source_pages": word_count,
        "results": sorted(
            results,
            key=lambda r: page_number(Path(r["file"]))
        ),
    }

    path = REPORT_DIR / datetime.now().strftime(
        "report_%Y%m%d_%H%M%S.json"
    )

    path.write_text(
        json.dumps(
            report,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    return path


def main():
    stop_event.clear()

    pdf_files = sorted(
        (
            pdf for pdf in PDF_DIR.glob("*.pdf")
            if page_number(pdf) >= 27
        ),
        key=page_number
    )

    if not pdf_files:
        print("[ERROR] No PDF files found.")
        return

    instructions = read_instructions(
        INSTRUCTIONS_FILE
    )

    imported = import_previous_results(pdf_files)

    pending = [
        pdf for pdf in pdf_files
        if not valid_file(
            OUTPUT_DIR / (pdf.stem + ".txt")
        )
    ]

    # Only the next 100 untranslated pages.
    batch = pending[:BATCH_SIZE]

    completed_before = len(pdf_files) - len(pending)

    print("=" * 55)
    print("BOOK TRANSLATION - NEXT BATCH")
    print(f"Total PDF files: {len(pdf_files)}")
    print(f"Imported old translations: {imported}")
    print(f"Already completed: {completed_before}")
    print(f"Selected for this batch: {len(batch)}")
    print(f"Workers: {MAX_WORKERS}")
    print("=" * 55)

    # Update the book using already saved translations.
    try:
        count = create_word(pdf_files)
        print(f"[WORD] Existing pages included: {count}")
    except Exception as error:
        print(f"[WORD ERROR] {error}")
        print("Close the Word document and run again.")
        return

    if not batch:
        print("[DONE] No untranslated PDF files remain.")
        return

    results = []
    new_pages = 0
    start = time.perf_counter()

    # Only MAX_WORKERS tasks are submitted at once.
    # New work stops if payment credits run out.
    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        iterator = iter(batch)

        futures = {}

        def submit_next():
            if stop_event.is_set():
                return

            pdf = next(iterator, None)

            if pdf is not None:
                future = executor.submit(
                    process_page,
                    pdf,
                    instructions
                )
                futures[future] = pdf

        for _ in range(min(MAX_WORKERS, len(batch))):
            submit_next()

        while futures:
            future = next(as_completed(futures))
            pdf = futures.pop(future)

            try:
                result = future.result()
            except Exception as error:
                result = {
                    "file": pdf.name,
                    "status": "failed",
                    "error": str(error),
                }

            results.append(result)

            if result["status"] == "success":
                new_pages += 1

                if new_pages % WORD_UPDATE_EVERY == 0:
                    try:
                        count = create_word(pdf_files)
                        print(
                            f"[WORD] Updated: "
                            f"{count} source pages"
                        )
                    except Exception as error:
                        print(f"[WORD ERROR] {error}")

            submit_next()

    seconds = time.perf_counter() - start

    try:
        count = create_word(pdf_files)
        print(f"[WORD] Final update: {count} pages")
    except Exception as error:
        count = None
        print(f"[WORD ERROR] {error}")

    report_path = save_report(
        results, batch, seconds, count
    )

    success = sum(
        r["status"] == "success" for r in results
    )

    failed = sum(
        r["status"] == "failed" for r in results
    )

    remaining = sum(
        not valid_file(
            OUTPUT_DIR / (pdf.stem + ".txt")
        )
        for pdf in pdf_files
    )

    print("\n" + "=" * 55)
    print("BATCH FINISHED")
    print(f"New successful pages: {success}")
    print(f"Failed pages: {failed}")
    print(f"Remaining in book: {remaining}")
    print(f"Time: {seconds:.1f} seconds")
    print(f"Word: {WORD_FILE}")
    print(f"Report: {report_path}")
    print("=" * 55)


if __name__ == "__main__":
    main()
