
import os
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types
import re
from dotenv import load_dotenv

# Configuration
PDF_FOLDER = "./pdf_files"
INSTRUCTIONS_FILE = "instructions.md"
OUTPUT_FILE = "merged_translation.docx"

# Gemini API Key
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise RuntimeError("GEMINI_API_KEY is missing")

client = genai.Client(api_key=API_KEY)


def read_instructions(md_path: str) -> str:
    path = Path(md_path)
    if not path.exists():
        print(
            f"[INFO] Instructions file '{md_path}' not found. "
            "Continuing without custom instructions."
        )
        return ""

    # קריאת קובץ טקסט/מארקדאון רגיל בקידוד UTF-8
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()


def translate_scanned_pdf(
    pdf_path: Path,
    instructions: str = ""
) -> str:
    instructions_block = ""
    if instructions:
        instructions_block = (
            f"\n--- Special Instructions ---\n"
            f"{instructions}\n"
            f"----------------------------\n"
        )

    prompt = f"""
    {instructions_block}

    [חוק עליון ומחייב לשריון התעתיק]:
    חל איסור מוחלט להוסיף שדה (ّ) על דעת עצמך או לנחש אותו מהתמונה.
    אסור בתכלית האיסור לערבב בין דגש עברי (ּ) לבין שדה ערבית (ّ).
    השתמש בשדה אך ורק במקום שהוא מופיע במפורש במקור הערבי,
    ואם הוא אינו קיים במקור – אל תוסיף אותו בשום אופן!
    """
    

    special_page_prompts = {
        "ACWA_A135.pdf": """
            This page continues commentary from the previous page.
            There is no new numbered proverb.
            Translate ALL visible Arabic text into Hebrew,
            including commentary and footnotes.
            Do not invent a proverb.
            Preserve paragraph order.
        """,

        "ACWA_A385.pdf": """
            This page contains commentary and paired poetry.
            There is no new numbered proverb.
            Translate ALL visible Arabic text into Hebrew,
            including commentary and footnotes.

            For paired poetry halves use:
            [[TABLE_START]]
            Hebrew right half ||| Hebrew left half
            Hebrew right half ||| Hebrew left half
            [[TABLE_END]]

            Preserve each original pair on one row.
            Keep prose outside the table.
            Do not invent or omit any text.
        """,

        "ACWA_A614.pdf": """
            This page contains a two-column comparison table.
            There is no new numbered proverb.
            Translate ALL visible Arabic text into Hebrew,
            including introduction and footnotes.

            For the comparison table use:
            [[TABLE_START]]
            Hebrew right cell ||| Hebrew left cell
            Hebrew right cell ||| Hebrew left cell
            [[TABLE_END]]

            Preserve every original row and its pairing.
            Keep prose outside the table.
            Do not invent or omit any text.
        """,
    }

    if pdf_path.name in special_page_prompts:
        prompt += (
            "\n\nIMPORTANT PAGE-SPECIFIC INSTRUCTIONS:\n"
            + special_page_prompts[pdf_path.name]
        )



    pdf_bytes = pdf_path.read_bytes()
    pdf_part = types.Part.from_bytes(
        data=pdf_bytes,
        mime_type="application/pdf"
    )

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=[pdf_part, prompt]
    )

    raw_text = response.text.strip() if response.text else ""


    if pdf_path.name in {
        "ACWA_A135.pdf",
        "ACWA_A385.pdf",
        "ACWA_A614.pdf",
    }:
        print(
            f"[DEBUG] {pdf_path.name}: "
            f"Gemini returned {len(raw_text)} characters"
        )
        print(f"[DEBUG] Preview: {raw_text[:500]!r}")
        
        if not raw_text:
            print(
                f"[DEBUG] Gemini candidates: "
                f"{response.candidates!r}"
            )
            raise ValueError(
                f"Gemini returned empty text: {pdf_path.name}"
            )



    # בדיקה נוספת מול הסריקה של שורות המקור הערביות בלבד
    
    special_pages = {
        "ACWA_A135.pdf",
        "ACWA_A385.pdf",
        "ACWA_A614.pdf",
    }

    if pdf_path.name in special_pages:
        if not raw_text.strip():
            raise ValueError("Empty special-page translation")

        if pdf_path.name in {
            "ACWA_A385.pdf",
            "ACWA_A614.pdf",
        }:
            if (
                raw_text.count("[[TABLE_START]]") != 1
                or raw_text.count("[[TABLE_END]]") != 1
            ):
                raise ValueError(
                    "Special-page table markers missing or duplicated"
                )

            table_body = raw_text.split("[[TABLE_START]]", 1)[1]
            table_body = table_body.split("[[TABLE_END]]", 1)[0]
            table_rows = [
                row.strip()
                for row in table_body.splitlines()
                if row.strip()
            ]

            if not table_rows or any(
                row.count("|||") != 1 for row in table_rows
            ):
                raise ValueError("Invalid special-page table rows")

        print(f"[CHECK] Special-page layout: {pdf_path.name}")
    else:
        raw_text = verify_arabic_source_lines(raw_text, pdf_part)


    return normalize_only_transcriptions(raw_text)


def _arabic_source_line(line: str):
    """
    Return (proverb number, Arabic content)
    for a numbered Arabic line.
    """
    if not re.search(r"[\u0621-\u064A]", line):
        return None

    m = re.match(r"^\s*(\d+)\.\s*(.+?)\s*$", line)
    if m:
        return int(m.group(1)), m.group(2)

    m = re.match(r"^\s*(.+?)\s+(\d+)\.\s*$", line)
    if m:
        return int(m.group(2)), m.group(1)

    return None


def verify_arabic_source_lines(text: str, pdf_part) -> str:
    """
    Second image-based review;
    replace ONLY numbered Arabic source lines.
    """
    lines = text.splitlines(keepends=True)
    originals = {}

    for i, line in enumerate(lines):
        item = _arabic_source_line(line.strip())

        if item:
            number, _ = item

            if number in originals:
                raise ValueError(
                    f"Duplicate Arabic proverb number: {number}"
                )

            originals[number] = i

    if not originals:
        raise ValueError(
            "No numbered Arabic proverb lines found for verification"
        )

    prompt = (
        "You are checking an OCR transcription against "
        "the ATTACHED SCANNED PDF. "
        "Independently inspect each printed Arabic proverb "
        "on the scan. "
        "Return exactly one line per requested proverb "
        "in the format NUMBER. ARABIC_TEXT. "
        "Copy ONLY what is visually printed, "
        "character by character, including diacritics. "
        "Do not normalize, guess, repair grammar, "
        "add shadda (U+0651), or omit printed shadda. "
        "Never output Hebrew letters. "
        "Do not translate or transliterate. "
        "If you cannot read a line confidently, "
        "output NUMBER. UNCLEAR instead. "
        "No explanation, headings, Markdown or code fences. "
        "Check especially the letter ta in proverb 4 "
        "and every other shadda.\n\n"
        "Lines from first pass to verify, NOT to blindly copy:\n"
        + "\n".join(
            lines[i].strip()
            for i in originals.values()
        )
    )

    for attempt in range(2):
        response = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=[pdf_part, prompt]
        )

        review = response.text.strip() if response.text else ""
        verified = {}
        valid = True

        for review_line in review.splitlines():
            match = re.fullmatch(
                r"\s*(\d+)\.\s*(.*?)\s*",
                review_line
            )

            if not match:
                valid = False
                break

            number = int(match.group(1))
            arabic = match.group(2)

            if (
                number not in originals
                or number in verified
                or arabic == "UNCLEAR"
                or not re.search(r"[\u0621-\u064A]", arabic)
                or re.search(r"[\u0590-\u05FF]", arabic)
            ):
                valid = False
                break

            verified[number] = arabic

        if valid and set(verified) == set(originals):
            for number, i in originals.items():
                # Preserve original newline convention
                newline = (
                    "\r\n"
                    if lines[i].endswith("\r\n")
                    else "\n"
                    if lines[i].endswith("\n")
                    else ""
                )

                lines[i] = (
                    f"{number}. {verified[number]}" + newline
                )

            print(
                f"  [CHECK] Second Arabic-source review "
                f"completed for {len(verified)} proverbs "
                f"(AI review, not guaranteed exact)."
            )

            return "".join(lines)

        print(
            f"  [WARNING] Arabic review response "
            f"invalid/incomplete "
            f"(attempt {attempt + 1}/2)."
        )

    raise ValueError(
        "Arabic source review failed: "
        "page not appended; check PDF manually"
    )


def normalize_transcription_accents(text: str) -> str:
    """
    תיקון ניקוד בתעתיק בלבד,
    ללא יצירת שדות חדשים.
    """

    # 1. החלפת תווי רפה בגרש תקני
    replacements = {
        'כֿ': 'כ׳',
        'כ\u05BF': 'כ׳',
        'גֿ': 'ג׳',
        'ג\u05BF': 'ג׳',
        'צֿ': 'צ׳',
        'צ\u05BF': 'צ׳',
        'דֿ': 'ד׳',
        'ד\u05BF': 'ד׳',
        'תֿ': 'ת׳',
        'ת\u05BF': 'ת׳',
        'זֿ': 'ז׳',
        'ז\u05BF': 'ז׳',
        '\u05BF': '׳',
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    # 2. העברת ניקוד שמופיע אחרי הגרש אל לפניו
    text = re.sub(
        r"([א-ת])׳([\u0591-\u05C7\u0651]+)",
        r"\1\2׳",
        text
    )

    # 3. תיקון דגשים בעברית ללא המצאת שדה ערבי
    pattern = r"([א-ת])([\u0591-\u05C7\u0651]*)([׳']?)"

    def fix_letter(match):
        letter, marks, geresh = match.groups()

        # הסרת סימני ניקוד כפולים
        marks = ''.join(dict.fromkeys(marks))

        has_shadda = '\u0651' in marks

        if geresh:
            # אות עם גרש: ללא דגש עברי פנימי
            marks = marks.replace('\u05BC', '')

        elif letter in 'בגדכת':
            # אותיות ב, ג, ד, כ, ת תמיד עם דגש עברי
            if '\u05BC' not in marks:
                marks += '\u05BC'

        elif has_shadda:
            # כל אות אחרת עם שדה מקבלת גם דגש עברי
            if '\u05BC' not in marks:
                marks += '\u05BC'

        return letter + marks + ('׳' if geresh else '')

    # 4. תיקון דגשים וניקוד
    text = re.sub(pattern, fix_letter, text)

    # 5. המרת אותיות בסוף מילה לאותיות סופיות
    final_letters = {
        'מ': 'ם',
        'נ': 'ן',
        'צ': 'ץ',
        'פ': 'ף',
        'כ': 'ך',
    }

    text = re.sub(
        r"([מנצפכ])([\u0591-\u05C7\u0651]*)(?![א-ת\u0591-\u05C7\u0651׳'])",
        lambda match: (
            final_letters[match.group(1)] + match.group(2)
        ),
        text
    )

    return text


def normalize_only_transcriptions(text: str) -> str:
    """
    תיקון שורות התעתיק בלבד,
    בלי לשנות עברית רגילה.
    """
    lines = text.splitlines()

    for i in range(len(lines) - 1):
        line = lines[i].strip()

        # זיהוי שורת מקור ערבית ממוספרת
        has_arabic = bool(
            re.search(r'[\u0621-\u064A]', line)
        )

        has_number = bool(
            re.match(r'^\d+[.)]\s*', line)
            or re.search(r'\s+\d+[.)]$', line)
        )

        if not (has_arabic and has_number):
            continue

        # אזהרה אם מופיעות אותיות עבריות במקור הערבי
        if re.search(r'[א-ת]', line):
            print(
                f"[WARNING] Hebrew letters in Arabic source: {line}"
            )

        # תיקון שורת התעתיק בלבד
        if re.search(r'[א-ת]', lines[i + 1]):
            lines[i + 1] = normalize_transcription_accents(
                lines[i + 1]
            )

    return '\n'.join(lines)


def main():
    folder = Path(PDF_FOLDER)

    if not folder.exists():
        print(
            f"[ERROR] Directory not found: "
            f"{folder.resolve()}"
        )
        return

    pdf_files = sorted([
        f for f in folder.iterdir()
        if f.suffix.lower() == ".pdf"
    ])

    print(
        f"Found {len(pdf_files)} PDF file(s) in directory."
    )

    if not pdf_files:
        print(
            f"[WARNING] No PDF files found in "
            f"'{folder.resolve()}'."
        )
        return

    custom_rules = read_instructions(INSTRUCTIONS_FILE)

    if custom_rules:
        print(
            f"[INFO] Instructions loaded successfully "
            f"from '{INSTRUCTIONS_FILE}'."
        )

    doc = Document()
    success_count = 0

    for idx, pdf_path in enumerate(pdf_files, 1):
        print(
            f"[{idx}/{len(pdf_files)}] Sending "
            f"'{pdf_path.name}' for translation..."
        )

        try:
            translation = translate_scanned_pdf(
                pdf_path,
                custom_rules
            )

            if not translation:
                print(
                    f"  [WARNING] No translation received "
                    f"for '{pdf_path.name}'."
                )
                continue

            # Add document title
            h = doc.add_heading(
                f"תרגום: {pdf_path.name}",
                level=1
            )

            h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            h.paragraph_format.rtl = True

            # Add paragraphs with RTL alignment
            for p_text in translation.split("\n\n"):
                if p_text.strip():
                    p = doc.add_paragraph(p_text.strip())
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    p.paragraph_format.rtl = True

            if idx < len(pdf_files):
                doc.add_page_break()

            success_count += 1

            print(
                f"  [SUCCESS] '{pdf_path.name}' "
                f"translated and appended successfully."
            )

        except Exception as err:
            print(
                f"  [ERROR] Failed to process "
                f"'{pdf_path.name}': {err}"
            )

    if success_count > 0:
        try:
            doc.save(OUTPUT_FILE)

            print(
                f"\n[DONE] Process complete! "
                f"{success_count} file(s) saved "
                f"to '{OUTPUT_FILE}'."
            )

        except PermissionError:
            print(
                f"\n[ERROR] Permission denied: "
                f"Please close '{OUTPUT_FILE}' "
                f"in Microsoft Word and run again."
            )
    else:
        print("\n[FAILED] No files were translated.")


if __name__ == "__main__":
    main()




