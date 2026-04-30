"""
Split a PDF into per-participant files based on a name list from a CSV.

Each participant receives a fixed number of consecutive pages from the source PDF.
Output files are named:  01. Name.pdf,  02. Name.pdf,  etc.

Use cases:
- Training certificates (2 pages per person)
- Exam result sheets (1 page per person)
- Any batch PDF where pages are already in participant order

Usage:
    python split_pdf_by_participants.py --pdf certificates.pdf --names names.csv
    python split_pdf_by_participants.py --pdf certificates.pdf --names names.csv --pages 1
    python split_pdf_by_participants.py --pdf certificates.pdf --names names.csv --col 0 --skip-header
    python split_pdf_by_participants.py --pdf certificates.pdf --names names.csv --output ./hasil
"""

import csv
import sys
import argparse
from pathlib import Path
from pypdf import PdfReader, PdfWriter


def sanitize(name: str) -> str:
    """Remove characters that are illegal in filenames."""
    import re
    return re.sub(r'[\\/:*?"<>|]', "_", str(name)).strip()


def load_names(csv_path: str, name_col: int, skip_header: bool) -> list[str]:
    """Read participant names from a CSV file."""
    names = []
    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        if skip_header:
            next(reader, None)
        for row in reader:
            if len(row) > name_col:
                name = row[name_col].strip()
                if name:
                    names.append(name)
    return names


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Split a PDF into individual files, one per participant.\n"
            "Names come from a CSV file; pages are assigned in order."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--pdf", "-p",
        required=True,
        help="Path to the source PDF to split"
    )
    parser.add_argument(
        "--names", "-n",
        required=True,
        help="Path to the CSV file containing participant names"
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=2,
        help="Number of PDF pages per participant  (default: 2)"
    )
    parser.add_argument(
        "--col",
        type=int,
        default=1,
        help="Column index (0-based) in the CSV that contains names  (default: 1)"
    )
    parser.add_argument(
        "--skip-header",
        action="store_true",
        help="Skip the first row of the CSV (use if it contains column headers)"
    )
    parser.add_argument(
        "--output", "-o",
        default=".",
        help="Output folder for split PDF files  (default: current directory)"
    )
    args = parser.parse_args()

    # ── Validate inputs ───────────────────────────────────────────────────────
    pdf_path = Path(args.pdf)
    csv_path = Path(args.names)
    out_dir  = Path(args.output)

    if not pdf_path.exists():
        sys.exit(f"\n[ERROR] PDF not found: {pdf_path}\n")
    if not csv_path.exists():
        sys.exit(f"\n[ERROR] CSV not found: {csv_path}\n")

    out_dir.mkdir(parents=True, exist_ok=True)

    # ── Load names ────────────────────────────────────────────────────────────
    participants = load_names(str(csv_path), args.col, args.skip_header)
    if not participants:
        sys.exit("\n[ERROR] No names found in CSV. Check --col and --skip-header.\n")

    # ── Load PDF ──────────────────────────────────────────────────────────────
    reader      = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    pages_needed = len(participants) * args.pages

    print(f"\n[INFO] PDF        : {pdf_path.name}  ({total_pages} pages)")
    print(f"[INFO] Participants: {len(participants)}")
    print(f"[INFO] Pages/person: {args.pages}")
    print(f"[INFO] Pages needed: {pages_needed}")
    print(f"[INFO] Output dir  : {out_dir.resolve()}\n")

    if total_pages < pages_needed:
        print(
            f"[WARN] PDF has {total_pages} pages but {pages_needed} are needed "
            f"for {len(participants)} participants × {args.pages} pages.\n"
            f"       Last {len(participants) - total_pages // args.pages} participant(s) may be incomplete.\n"
        )

    # ── Split ─────────────────────────────────────────────────────────────────
    page_index  = 0
    success     = 0
    incomplete  = 0

    for i, name in enumerate(participants):
        if page_index >= total_pages:
            print(f"  ⚠  No pages left for: {name}")
            incomplete += 1
            continue

        writer = PdfWriter()
        pages_added = 0

        for _ in range(args.pages):
            if page_index < total_pages:
                writer.add_page(reader.pages[page_index])
                page_index += 1
                pages_added += 1

        filename = out_dir / f"{i + 1:02d}. {sanitize(name)}.pdf"
        with open(filename, "wb") as out_file:
            writer.write(out_file)

        status = "✓" if pages_added == args.pages else f"⚠ only {pages_added}/{args.pages} pages"
        print(f"  {status}  {filename.name}")

        if pages_added == args.pages:
            success += 1
        else:
            incomplete += 1

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n{'─' * 50}")
    print(f"  Done.  ✓ {success} complete   ⚠ {incomplete} incomplete")
    print(f"  Files saved to: {out_dir.resolve()}")
    print(f"{'─' * 50}\n")


if __name__ == "__main__":
    main()
