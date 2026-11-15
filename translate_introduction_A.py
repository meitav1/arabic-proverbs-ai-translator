
"""Translate ACWA_A001..A026 introduction scans into a separate Hebrew Word file.
Install: pip install google-genai python-docx
Set GEMINI_API_KEY in your environment before running.
"""
import os
import re
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from google import genai
from google.genai import types
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt
from dotenv import load_dotenv

PDF_DIR = Path('pdf_files_A_intro')
TEXT_DIR = Path('translated_intro_A')
OUTPUT = Path('book_A_introduction.docx')
MODEL = 'gemini-3.8-flash'
MAX_WORKERS = 15
MAX_RETRIES = 4
FILES = [f'ACWA_A{i:03d}.pdf' for i in range(1, 27)]

# Gemini API Key
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise RuntimeError("GEMINI_API_KEY is missing")

client = genai.Client(api_key=API_KEY)


# Verified visually blank page; don't ask Gemini to invent text.
BLANK_PAGES = {'ACWA_A024.pdf'}

PROMPT = '''Translate this scanned Arabic book introduction page into HEBREW ONLY.
Translate EVERY visible meaningful Arabic word faithfully, including headings, paragraphs,
footnotes, quoted sayings, bibliographic details, dates, and poetry.
No Arabic source lines, no Hebrew transliteration, and NO numbered-proverb three-line format.
Do not add interpretation, invented content, explanations or headings absent from the page.
Keep the original reading order, paragraph breaks and all footnotes. Preserve names and dates.
For paired half-lines of poetry printed in two columns, use this exact structure:
[[TABLE_START]]
translated right half ||| translated left half
[[TABLE_END]]
Each original paired verse must be exactly one row; put non-poetry prose outside the table.
For all other text use plain Hebrew paragraphs. Do not use markdown or code fences.
Start with a line exactly [עמוד N] where N is the PRINTED Arabic page number
converted to Western digits if clearly visible. If no page number is visible,
use [עמוד ללא מספר]. For a completely blank scan, output exactly [עמוד ריק].
A sentence cut off by the end of the page must remain cut off, not completed.
'''

def get_client():
    return genai.Client(api_key=API_KEY)

def translate_one(path):
    target = TEXT_DIR / f'{path.stem}.txt'
    if target.exists() and target.read_text(encoding='utf-8').strip():
        return path.name, 'cached'

    if path.name in BLANK_PAGES:
        target.write_text('[עמוד ריק]', encoding='utf-8')
        return path.name, 'blank'

    pdf = types.Part.from_bytes(
        data=path.read_bytes(),
        mime_type='application/pdf'
    )
    client = get_client()
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=[pdf, PROMPT]
            )
            content = (response.text or '').strip()

            if not content:
                raise ValueError('Empty Gemini response')

            if content.count('[[TABLE_START]]') != content.count('[[TABLE_END]]'):
                raise ValueError('Unbalanced poetry table markers')

            

            if '[[TABLE_START]]' in content:
                body = content.split('[[TABLE_START]]', 1)[1].split('[[TABLE_END]]', 1)[0]
                rows = [x.strip() for x in body.splitlines() if x.strip()]
                if not rows or any(x.count('|||') != 1 for x in rows):
                    raise ValueError('Malformed poetry table rows')

            target.write_text(content, encoding='utf-8')
            return path.name, 'translated'

        except Exception as exc:
            last_error = exc
            print(
                f'[RETRY] {path.name} attempt {attempt}/{MAX_RETRIES}: {exc}',
                flush=True
            )
            if attempt < MAX_RETRIES:
                time.sleep(min(2 ** attempt, 12))

    raise RuntimeError(f'{path.name}: {last_error}')

def add_paragraph(doc, content):
    if not content.strip():
        return

    para = doc.add_paragraph(content.strip())
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para.paragraph_format.rtl = True

def add_table(doc, content):
    rows = [r.strip() for r in content.splitlines() if r.strip()]
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for line in rows:
        right, left = [x.strip() for x in line.split('|||', 1)]
        cells = table.add_row().cells

        for cell, value in zip(cells, (right, left)):
            cell.text = value
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.rtl = True

def add_page(doc, text):
    # Retain order of prose and poetry tables.
    pattern = re.compile(
        r'\[\[TABLE_START\]\](.*?)\[\[TABLE_END\]\]',
        re.S
    )
    pos = 0

    for match in pattern.finditer(text):
        for part in re.split(r'\n\s*\n', text[pos:match.start()]):
            add_paragraph(doc, part)

        add_table(doc, match.group(1))
        pos = match.end()

    for part in re.split(r'\n\s*\n', text[pos:]):
        add_paragraph(doc, part)

def build_word():
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.size = Pt(11)
    included = 0

    for name in FILES:
        path = TEXT_DIR / (Path(name).stem + '.txt')

        if not path.exists() or not path.read_text(encoding='utf-8').strip():
            continue

        if included:
            doc.add_page_break()

        add_page(doc, path.read_text(encoding='utf-8'))
        included += 1

    if included:
        doc.save(OUTPUT)

    return included

def main():
    TEXT_DIR.mkdir(parents=True, exist_ok=True)

    missing = [
        name for name in FILES
        if not (PDF_DIR / name).exists()
    ]

    if missing:
        raise FileNotFoundError(
            f'Missing {len(missing)} PDFs in {PDF_DIR}: {missing}'
        )

    paths = [PDF_DIR / name for name in FILES]
    failures = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {
            pool.submit(translate_one, path): path
            for path in paths
        }

        for future in as_completed(futures):
            try:
                name, status = future.result()
                print(f'[OK] {name}: {status}', flush=True)
            except Exception as exc:
                failures.append(str(exc))
                print(f'[FAILED] {exc}', flush=True)

    count = build_word()

    print(
        f'[WORD] {OUTPUT}: {count}/{len(FILES)} pages included'
    )

    if failures:
        print('[FAILED PAGES] ' + ' | '.join(failures))
        print(
            'Run again to retry only failed pages; '
            'completed TXT files are retained.'
        )
    else:
        print(
            '[DONE] All introduction pages processed. '
            'Review translation against scans.'
        )

if __name__ == '__main__':
    main()
