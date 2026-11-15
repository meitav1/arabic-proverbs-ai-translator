
import re
import time
import json
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from translate_and_merge import (
    translate_scanned_pdf,
    read_instructions,
    PDF_FOLDER,
    INSTRUCTIONS_FILE,
)

# ===== Experiment settings =====

MAX_PAGES = 10

# Start with 1, then test 2
MAX_WORKERS = 15

# Delay between starting jobs, in seconds
START_DELAY = 0

# Keep each experiment separate
OUTPUT_FOLDER = Path(
    f"parallel_results/workers_{MAX_WORKERS}"
)

# =================================

OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

print_lock = threading.Lock()


def page_number(path):
    match = re.search(r"(\d+)$", path.stem)
    if not match:
        raise ValueError(
            f"Cannot find page number: {path.name}"
        )
    return int(match.group(1))


def process_page(pdf_path, instructions):
    output_path = OUTPUT_FOLDER / (
        pdf_path.stem + ".txt"
    )

    # Do not repeat completed work
    if output_path.exists():
        return {
            "file": pdf_path.name,
            "status": "skipped",
            "seconds": 0,
        }

    start = time.perf_counter()

    try:
        translation = translate_scanned_pdf(
            pdf_path,
            instructions
        )

        if not translation.strip():
            raise ValueError("Empty translation")

        # Save only after successful processing
        temp_path = output_path.with_suffix(".tmp")
        temp_path.write_text(
            translation,
            encoding="utf-8"
        )
        temp_path.replace(output_path)

        duration = time.perf_counter() - start

        with print_lock:
            print(
                f"[SUCCESS] {pdf_path.name} "
                f"- {duration:.1f} seconds"
            )

        return {
            "file": pdf_path.name,
            "status": "success",
            "seconds": round(duration, 2),
        }

    except Exception as error:
        duration = time.perf_counter() - start

        with print_lock:
            print(
                f"[ERROR] {pdf_path.name}: {error}"
            )

        return {
            "file": pdf_path.name,
            "status": "failed",
            "seconds": round(duration, 2),
            "error": str(error),
        }


def main():
    folder = Path(PDF_FOLDER)

    files = [
        folder / f"ACWA_A{page:03d}.pdf"
        for page in range(95, 125)
    ]

    if not files:
        print("[ERROR] No PDF files found.")
        return

    instructions = read_instructions(
        INSTRUCTIONS_FILE
    )

    print(f"Pages selected: {len(files)}")
    print(f"Parallel workers: {MAX_WORKERS}")
    print(f"Output folder: {OUTPUT_FOLDER}")
    print("-" * 50)

    start = time.perf_counter()
    results = []

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        futures = []

        for pdf_path in files:
            future = executor.submit(
                process_page,
                pdf_path,
                instructions
            )
            futures.append(future)

            time.sleep(START_DELAY)

        for future in as_completed(futures):
            results.append(future.result())

    total_seconds = time.perf_counter() - start

    results.sort(
        key=lambda item: page_number(
            Path(item["file"])
        )
    )

    report = {
        "workers": MAX_WORKERS,
        "pages_selected": len(files),
        "total_seconds": round(total_seconds, 2),
        "results": results,
    }

    report_path = OUTPUT_FOLDER / "report.json"

    report_path.write_text(
        json.dumps(
            report,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    success = sum(
        item["status"] == "success"
        for item in results
    )

    skipped = sum(
        item["status"] == "skipped"
        for item in results
    )

    failed = sum(
        item["status"] == "failed"
        for item in results
    )

    print("\n" + "=" * 50)
    print("EXPERIMENT COMPLETED")
    print(f"Successful: {success}")
    print(f"Skipped: {skipped}")
    print(f"Failed: {failed}")
    print(f"Total time: {total_seconds:.1f} seconds")
    print(f"Report: {report_path}")
    print("=" * 50)


if __name__ == "__main__":
    main()
