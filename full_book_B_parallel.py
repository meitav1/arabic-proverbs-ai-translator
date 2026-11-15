
"""Second-book translator. Keeps all first-book artifacts untouched."""
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types

from translate_and_merge import (
    translate_scanned_pdf,
    read_instructions,
    INSTRUCTIONS_FILE,
    client,
)

PDF_DIR = Path("אכוע חלק 2") / "אלאכוע חתוך 2"
OUTPUT_DIR = Path("translated_pages_B")
REPORT_DIR = Path("translation_reports_B")
WORD_FILE = Path("full_book_B_translation.docx")

MAX_WORKERS = 15
BATCH_SIZE = 150
MAX_RETRIES = 4
WORD_UPDATE_EVERY = 10

END_PAGES = {23, 24, 25, 150} | set(range(711, 723))

END_PROMPT = """
Translate ALL visible Arabic text on this scanned page into accurate Hebrew ONLY.
This page may contain prose, explanations, poetry, notes, or bibliography.
Translate all visible content into Hebrew, preserving its structure.
Do not require numbered proverbs.
Do not transcribe Arabic or use the three-line proverb format.
Do not omit any headings, notes, references, names, numbers, or text.
Do not invent text.

Preserve the original order and structure.

For any paired two-column poetry, use exactly this representation
(and repeat separately for multiple groups):

[[TABLE_START]]
Hebrew translation of right half ||| Hebrew translation of left half
[[TABLE_END]]

Each original paired verse must be on one row.
Keep prose and notes outside tables.

For bibliography, preserve one entry per source entry,
including all bibliographic details.
Do not join unrelated entries.

Use [עמוד N] with the page number from the filename.
Translate text cut off at the page boundary only as visible.
Do not invent continuations.
Do not include introductory remarks or Markdown code fences.
"""

OUTPUT_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)

print_lock = threading.Lock()
stop_event = threading.Event()


def number(pdf):
    match = re.fullmatch(
        r"ACWA_B(\d+)\.pdf",
        pdf.name,
        re.IGNORECASE,
    )
    if not match:
        raise ValueError(f"Unexpected filename: {pdf.name}")
    return int(match.group(1))


def validate_tables(text, filename):
    inside = False
    count = 0

    for line in text.splitlines():
        s = line.strip()

        if s == "[[TABLE_START]]":
            if inside:
                raise ValueError(f"Nested table: {filename}")
            inside = True
            count = 0

        elif s == "[[TABLE_END]]":
            if not inside or count == 0:
                raise ValueError(f"Invalid/empty table: {filename}")
            inside = False

        elif inside and s:
            if s.count("|||") != 1:
                raise ValueError(f"Invalid table row: {filename}")
            count += 1

    if inside:
        raise ValueError(f"Unclosed table: {filename}")

    if text.count("[[TABLE_START]]") != text.count("[[TABLE_END]]"):
        raise ValueError(f"Unbalanced tables: {filename}")


def translate_end_page(pdf):
    

    pdf_part = types.Part.from_bytes(
        data=pdf.read_bytes(),
        mime_type="application/pdf",
    )

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=[
            pdf_part,
            f"Filename page number: {number(pdf)}.\n" + END_PROMPT,
        ],
    )

    result = response.text.strip() if response.text else ""

    if not result:
        raise ValueError("Empty translation")

    if number(pdf) == 23:
        result = result.replace("[[TABLE_START]]", "")
        result = result.replace("[[TABLE_END]]", "")
        result = result.replace("|||", " — ")
        return result

    validate_tables(result, pdf.name)
    return result


def is_payment_error(error):
    msg = str(error).lower()

    return (
        "402" in msg
        or "prepayment credits are depleted" in msg
        or "credit balance" in msg
    )


def process(pdf, instructions):
    target = OUTPUT_DIR / (pdf.stem + ".txt")

    if target.exists() and target.stat().st_size:
        return ("cached", pdf.name, "")

    for attempt in range(1, MAX_RETRIES + 1):
        if stop_event.is_set():
            return ("stopped", pdf.name, "")

        try:
            if number(pdf) in END_PAGES:
                result = translate_end_page(pdf)
            else:
                # Original translation and Arabic verification.
                result = translate_scanned_pdf(
                    pdf,
                    instructions,
                )

            if not result or not result.strip():
                raise ValueError("Empty translation")

            temp = target.with_suffix(".tmp")
            temp.write_text(result, encoding="utf-8")
            temp.replace(target)

            return ("translated", pdf.name, "")

        except Exception as error:
            if is_payment_error(error):
                stop_event.set()
                return (
                    "failed",
                    pdf.name,
                    f"Credits/payment: {error}",
                )

            if attempt == MAX_RETRIES:
                return ("failed", pdf.name, str(error))

            with print_lock:
                print(
                    f"[RETRY] {pdf.name} "
                    f"attempt {attempt}/{MAX_RETRIES}: {error}"
                )

            time.sleep(
                min(60, 2 ** attempt + random.uniform(0, 2))
            )

    return ("failed", pdf.name, "Unknown error")


def add_paragraph(doc, text):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.rtl = True


def add_table(doc, rows):
    table = doc.add_table(rows=0, cols=2)

    for right, left in rows:
        cells = table.add_row().cells

        for cell, value in zip(
            cells,
            (right, left),
        ):
            cell.text = value

            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.rtl = True


def build_word(pdfs):
    doc = Document()
    included = 0

    for pdf in pdfs:
        txt = OUTPUT_DIR / (pdf.stem + ".txt")

        if not txt.exists() or not txt.stat().st_size:
            continue

        content = txt.read_text(encoding="utf-8")
        validate_tables(content, pdf.name)

        if included:
            doc.add_page_break()

        inside = False
        rows = []

        for line in content.splitlines():
            s = line.strip()

            if s == "[[TABLE_START]]":
                inside = True
                rows = []

            elif s == "[[TABLE_END]]":
                add_table(doc, rows)
                inside = False

            elif inside:
                if s:
                    right, left = line.split("|||", 1)
                    rows.append(
                        (right.strip(), left.strip())
                    )

            else:
                add_paragraph(doc, line)

        included += 1

    if included:
        temp = WORD_FILE.with_name(
            WORD_FILE.stem + ".tmp.docx"
        )
        doc.save(temp)
        temp.replace(WORD_FILE)

    return included


def main():
    pdfs = sorted(
    [
        p for p in PDF_DIR.glob("*.pdf")
        if re.fullmatch(
            r"ACWA_B\d+\.pdf",
            p.name,
            re.IGNORECASE,
        )
        and number(p) >= 5
    ],
    key=number,
)

    if not pdfs:
        print(
            f"[ERROR] No ACWA_B###.pdf files in "
            f"{PDF_DIR.resolve()}"
        )
        return

    found = {number(p) for p in pdfs}
    missing_end = sorted(END_PAGES - found)

    if missing_end:
        print(
            f"[WARNING] Missing final pages: {missing_end}"
        )

    instructions = read_instructions(INSTRUCTIONS_FILE)

    pending = [
        p for p in pdfs
        if not (
            (OUTPUT_DIR / (p.stem + ".txt")).exists()
            and (OUTPUT_DIR / (p.stem + ".txt")).stat().st_size > 0
        )
    ]

    batch = pending[:BATCH_SIZE]

    print(
        f"[START] B book: {len(pdfs)} PDFs, "
        f"{len(pending)} pending, batch {len(batch)}"
    )

    print(
        f"[WORD] Existing translations: "
        f"{build_word(pdfs)}"
    )

    results = []
    new = 0

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        iterator = iter(batch)
        futures = {}

        def submit_next():
            if stop_event.is_set():
                return

            p = next(iterator, None)

            if p is not None:
                futures[
                    executor.submit(
                        process,
                        p,
                        instructions,
                    )
                ] = p

        for _ in range(
            min(MAX_WORKERS, len(batch))
        ):
            submit_next()

        while futures:
            future = next(as_completed(futures))
            pdf = futures.pop(future)

            try:
                status, name, detail = future.result()

            except Exception as exc:
                status, name, detail = (
                    "failed",
                    pdf.name,
                    str(exc),
                )

            print(
                f"[{status.upper()}] {name}"
                + (f": {detail}" if detail else "")
            )

            results.append(
                (status, name, detail)
            )

            if status == "translated":
                new += 1

                if new % WORD_UPDATE_EVERY == 0:
                    print(
                        f"[WORD] Updated: "
                        f"{build_word(pdfs)} pages"
                    )

            submit_next()

    count = build_word(pdfs)

    remaining = sum(
        not (
            (OUTPUT_DIR / (p.stem + ".txt")).exists()
            and (OUTPUT_DIR / (p.stem + ".txt")).stat().st_size > 0
        )
        for p in pdfs
    )

    report = REPORT_DIR / (
        f"report_{datetime.now():%Y%m%d_%H%M%S}.txt"
    )

    report.write_text(
        "\n".join(
            f"{s}\t{n}\t{d}"
            for s, n, d in results
        ),
        encoding="utf-8",
    )

    print(
        f"[WORD] Final update: "
        f"{count}/{len(pdfs)} source pages"
    )

    print(
        f"[DONE] Remaining: {remaining}; "
        f"Word: {WORD_FILE}; report: {report}"
    )


if __name__ == "__main__":
    main()
