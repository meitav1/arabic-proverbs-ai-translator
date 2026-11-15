# AI-Powered Arabic Proverbs Translation & Document Processing

An automated Python-based pipeline for processing scanned Arabic proverb books, generating Hebrew transliterations and translations, and producing structured Microsoft Word documents using Google's Gemini AI.

## Project Overview

This project was developed to automate the processing of large collections of scanned Arabic texts, with a particular focus on Yemenite-Arabic proverbs.

The system combines AI-powered text extraction, Hebrew transliteration, translation, and automated document generation.

The pipeline was designed to handle hundreds of pages while maintaining a strong emphasis on source-text accuracy, linguistic consistency, and processing efficiency.

## Key Features

- **AI-Powered OCR:** Extracts Arabic text from scanned PDF pages using Gemini's multimodal capabilities.
- **Hebrew Transliteration:** Generates Hebrew transliterations with vowel marks according to defined linguistic rules.
- **Automated Translation:** Produces Hebrew translations and explanations of Arabic proverbs.
- **Multi-Pass Validation:** Uses additional AI review passes to identify potential transcription and source-text errors.
- **Parallel Processing:** Processes multiple pages concurrently to improve throughput.
- **Word Document Generation:** Combines processed pages into structured `.docx` documents.
- **Source Preservation:** Applies strict rules to minimize unintended changes to the original Arabic text.

## Technologies

- **Language:** Python
- **AI:** Google Gemini API
- **PDF Processing:** PyMuPDF
- **Document Generation:** python-docx
- **Concurrency:** Python parallel processing
- **Configuration:** python-dotenv

## Project Structure

| File | Description |
|---|---|
| `translate_and_merge.py` | Core translation and document-processing pipeline |
| `full_book_parallel.py` | Parallel processing for Book A |
| `full_book_B_parallel.py` | Parallel processing for Book B |
| `parallel_translate.py` | Parallel translation utilities |
| `translate_introduction_A.py` | Processing the introduction of Book A |
| `measure_one_page.py` | Single-page processing and performance measurement |
| `instructions.md` | Translation and transliteration rules |
| `requirements.txt` | Python dependencies |
| `.env.example` | Example environment configuration |

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
cd YOUR_REPOSITORY
```

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

**3. Configure the Gemini API key**

Create a `.env` file in the project directory:

```env
GEMINI_API_KEY=your_api_key_here
```

Obtain your own API key from Google AI Studio.

**4. Prepare input files**

Place the source PDF files in the appropriate input directories, such as `pdf_files/`.

**5. Run the processing pipeline**

For example, to process Book B:

```bash
python full_book_B_parallel.py
```

Processing depends on the expected input files and directory structure.

## Accuracy and Validation

The pipeline emphasizes preserving the original Arabic source text, including letters, diacritics, and other linguistic details.

Dedicated transliteration rules support consistent Hebrew representations, including vowel marks, final letter forms, and Yemenite-Arabic phonetic conventions.

Additional AI review passes help identify potential errors. As with any AI-based OCR and translation system, results may still require manual verification.

## Performance Optimization

The project uses parallel processing to improve throughput across large collections of PDF pages.

Performance improvements were designed to preserve the existing AI model configuration, translation instructions, and validation workflow rather than reducing the quality of processing.

## Security and Data

API credentials are loaded from environment variables and are excluded from version control.

Original scanned books and generated translation documents are not included in this repository.

## Project Status

Developed as an independent software and AI document-processing project, with full-book batch processing implemented for large scanned PDF collections.

## Author

Meitav Bin Nun

Computer Science Student | Python & AI Development