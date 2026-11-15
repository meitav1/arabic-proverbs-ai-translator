
import json
import time
from pathlib import Path
from unittest.mock import patch

import translate_and_merge as translator

# First measurement: one page only
PDF_FILE = Path("pdf_files/ACWA_A033.pdf")
OUTPUT_FOLDER = Path("measurement_results")


def main():
    if not PDF_FILE.exists():
        print(f"[ERROR] File not found: {PDF_FILE}")
        return

    OUTPUT_FOLDER.mkdir(exist_ok=True)

    instructions = translator.read_instructions(
        translator.INSTRUCTIONS_FILE
    )

    usage_records = []
    original_method = type(
        translator.client.models
    ).generate_content

    def measured_generate_content(self, *args, **kwargs):
        call_number = len(usage_records) + 1
        start = time.perf_counter()

        response = original_method(self, *args, **kwargs)

        elapsed = time.perf_counter() - start
        usage = response.usage_metadata

        record = {
            "call": call_number,
            "seconds": round(elapsed, 2),
            "input_tokens": (
                usage.prompt_token_count if usage else None
            ),
            "output_tokens": (
                usage.candidates_token_count if usage else None
            ),
            "total_tokens": (
                usage.total_token_count if usage else None
            ),
            "thinking_tokens": (
                usage.thoughts_token_count if usage else None
            ),
        }

        usage_records.append(record)

        print(
            f"[API CALL {call_number}] "
            f"Input: {record['input_tokens']}, "
            f"Output: {record['output_tokens']}, "
            f"Total: {record['total_tokens']}, "
            f"Time: {elapsed:.1f}s"
        )

        return response

    start = time.perf_counter()

    try:
        with patch.object(
            type(translator.client.models),
            "generate_content",
            measured_generate_content
        ):
            translation = translator.translate_scanned_pdf(
                PDF_FILE,
                instructions
            )

        total_time = time.perf_counter() - start

        if not translation.strip():
            raise ValueError("Empty translation")

        output_file = OUTPUT_FOLDER / (
            PDF_FILE.stem + ".txt"
        )
        output_file.write_text(
            translation,
            encoding="utf-8"
        )

        print(f"\n[SUCCESS] Saved: {output_file}")

    except Exception as error:
        total_time = time.perf_counter() - start
        print(f"\n[ERROR] {error}")

    finally:
        report = {
            "file": PDF_FILE.name,
            "total_seconds": round(total_time, 2),
            "api_calls": usage_records,
        }

        report_file = OUTPUT_FOLDER / "usage_report.json"
        report_file.write_text(
            json.dumps(
                report,
                ensure_ascii=False,
                indent=2
            ),
            encoding="utf-8"
        )

        print(f"[REPORT] {report_file}")


if __name__ == "__main__":
    main()
