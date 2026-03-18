from __future__ import annotations

import argparse
import csv
from datetime import datetime
import os
from pathlib import Path
import re

from modules._shared.config import EXPORT_DIR, LOG_DIR_NAME, REPORT_NAME, TARGET_SOURCE_PERIOD_KEY


LOG_PREFIX = "automation_log"
EXPORT_PROGRESS_FILENAME = "export_progress.csv"
EXPORT_HEADER = [
    "report_name",
    "company_name",
    "source_period_key",
    "export_period_key",
    "month",
    "export_filename",
    "export_path",
]
REPORT_PREFIX = f"{REPORT_NAME}-"
YEAR_TEXT = "\u5e74"
MONTH_TEXT = "\u6708"
EXPORT_FILE_RE = re.compile(
    rf"^{re.escape(REPORT_NAME)}-(?P<company>.+?)\("
    rf"(?P<start_year>\d{{4}}){YEAR_TEXT}(?P<start_month>\d{{1,2}}){MONTH_TEXT}\s*-\s*"
    rf"(?P<end_year>\d{{4}}){YEAR_TEXT}(?P<end_month>\d{{1,2}}){MONTH_TEXT}\)$"
)
PERIOD_KEY_RE = re.compile(
    r"^(?P<start_year>\d{4})\.(?P<start_month>\d{2})-(?P<end_year>\d{4})\.(?P<end_month>\d{2})$"
)


def get_desktop_log_dir() -> Path:
    env_log_dir = os.environ.get("DXN_LOG_DIR", "").strip()
    if env_log_dir:
        log_dir = Path(env_log_dir)
    else:
        project_root = Path(__file__).resolve().parents[2]
        log_dir = project_root / "runtime" / LOG_DIR_NAME
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def get_log_file_path() -> Path:
    stamp = datetime.now().strftime("%Y%m%d")
    return get_desktop_log_dir() / f"{LOG_PREFIX}_{stamp}.txt"


def get_export_csv_path() -> Path:
    return get_desktop_log_dir() / EXPORT_PROGRESS_FILENAME


def write_log(step: str, status: str, message: str = "") -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}] [{status}] {step}"
    if message:
        line += f" | {message}"

    with open(get_log_file_path(), "a", encoding="utf-8") as file:
        file.write(line + "\n")

    print(line)


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_period_key(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""

    match = PERIOD_KEY_RE.match(text)
    if not match:
        return ""

    return (
        f"{int(match.group('start_year')):04d}.{int(match.group('start_month')):02d}-"
        f"{int(match.group('end_year')):04d}.{int(match.group('end_month')):02d}"
    )


def normalize_month(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    try:
        month = int(text)
    except Exception as exc:
        raise ValueError(f"Invalid month value: {value}") from exc
    if not 1 <= month <= 12:
        raise ValueError(f"Month out of range: {month}")
    return str(month)


def normalize_filename(raw_filename: str) -> str:
    filename = normalize_text(raw_filename)
    filename = re.sub(r"\.(xls|xlsx|csv)$", "", filename, flags=re.IGNORECASE)
    if not filename:
        raise ValueError("Export filename is empty after normalization")
    return filename


def parse_export_filename(filename: str, source_period_key: str | None = None) -> dict:
    filename = normalize_filename(filename)
    match = EXPORT_FILE_RE.match(filename)
    if not match:
        raise ValueError(f"Cannot parse export filename: {filename}")

    start_year = int(match.group("start_year"))
    start_month = int(match.group("start_month"))
    end_year = int(match.group("end_year"))
    end_month = int(match.group("end_month"))
    export_period_key = f"{start_year:04d}.{start_month:02d}-{end_year:04d}.{end_month:02d}"

    return {
        "report_name": REPORT_NAME,
        "company_name": normalize_text(match.group("company")),
        "source_period_key": normalize_period_key(source_period_key) or TARGET_SOURCE_PERIOD_KEY,
        "export_period_key": export_period_key,
        "month": str(end_month),
        "export_filename": filename,
    }


def _canonicalize_export_row(row: dict) -> tuple[str, str, str, str, str, str, str] | None:
    report_name = normalize_text(row.get("report_name", "")) or REPORT_NAME
    company_name = normalize_text(row.get("company_name", ""))
    source_period_key = normalize_period_key(row.get("source_period_key", "")) or TARGET_SOURCE_PERIOD_KEY
    export_period_key = normalize_period_key(row.get("export_period_key", ""))
    month = normalize_text(row.get("month", ""))
    export_filename = normalize_text(row.get("export_filename", ""))
    export_path = normalize_text(row.get("export_path", ""))

    if not export_filename:
        return None

    parsed = parse_export_filename(export_filename, source_period_key=source_period_key)
    company_name = company_name or parsed["company_name"]
    export_period_key = export_period_key or parsed["export_period_key"]
    month = normalize_month(month or parsed["month"])
    export_filename = parsed["export_filename"]

    if not company_name:
        return None

    return (
        report_name,
        company_name,
        source_period_key,
        export_period_key,
        month,
        export_filename,
        export_path,
    )


def _rewrite_export_csv_with_current_header(csv_file: Path) -> None:
    rows_to_keep: list[tuple[str, str, str, str, str, str, str]] = []
    seen: set[tuple[str, str, str, str, str, str, str]] = set()

    with open(csv_file, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            normalized_row = _canonicalize_export_row(row)
            if normalized_row is None or normalized_row in seen:
                continue
            seen.add(normalized_row)
            rows_to_keep.append(normalized_row)

    with open(csv_file, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(EXPORT_HEADER)
        writer.writerows(rows_to_keep)


def ensure_export_csv_header(csv_file: Path | None = None) -> None:
    csv_file = csv_file or get_export_csv_path()

    if csv_file.exists() and csv_file.stat().st_size > 0:
        with open(csv_file, "r", encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)
            fieldnames = reader.fieldnames or []

        if fieldnames != EXPORT_HEADER:
            _rewrite_export_csv_with_current_header(csv_file)
        return

    csv_file.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_file, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(EXPORT_HEADER)


def append_export_record(
    export_filename: str,
    *,
    export_path: str = "",
    source_period_key: str | None = None,
) -> None:
    csv_file = get_export_csv_path()
    ensure_export_csv_header(csv_file)
    parsed = parse_export_filename(export_filename, source_period_key=source_period_key)

    row = [
        parsed["report_name"],
        parsed["company_name"],
        parsed["source_period_key"],
        parsed["export_period_key"],
        parsed["month"],
        parsed["export_filename"],
        normalize_text(export_path),
    ]

    with open(csv_file, "a", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(row)

    print(
        "Appended export record | "
        f"{parsed['company_name']} | month={parsed['month']} | {parsed['export_filename']}"
    )


def write_export_success(
    step: str,
    export_filename: str,
    *,
    export_path: str = "",
    source_period_key: str | None = None,
    message: str = "",
) -> None:
    append_export_record(
        export_filename=export_filename,
        export_path=export_path,
        source_period_key=source_period_key,
    )
    log_message = message or export_filename
    if export_path:
        log_message = f"{log_message} | {export_path}"
    write_log(step, "OK", log_message)


def _iter_export_files(export_dir: Path):
    patterns = ("*.xlsx", "*.xls")
    seen: set[Path] = set()

    for pattern in patterns:
        for path in export_dir.rglob(pattern):
            if not path.is_file():
                continue
            if not normalize_text(path.stem).startswith(REPORT_PREFIX):
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            yield path


def rebuild_export_csv_from_export_dir(
    export_dir: Path | None = None,
    output_csv: Path | None = None,
    source_period_key: str | None = None,
) -> Path:
    export_dir = export_dir or EXPORT_DIR
    output_csv = output_csv or get_export_csv_path()

    if not export_dir.exists():
        raise FileNotFoundError(f"Export directory not found: {export_dir}")
    if not export_dir.is_dir():
        raise NotADirectoryError(f"Export path is not a directory: {export_dir}")

    rows_to_keep: list[tuple[str, str, str, str, str, str, str]] = []
    seen_rows: set[tuple[str, str, str, str, str, str, str]] = set()
    skipped_files: list[tuple[str, str]] = []

    for export_file in sorted(_iter_export_files(export_dir), key=lambda path: str(path).lower()):
        try:
            parsed = parse_export_filename(export_file.name, source_period_key=source_period_key)
        except Exception as exc:
            skipped_files.append((export_file.name, str(exc)))
            continue

        row = (
            parsed["report_name"],
            parsed["company_name"],
            parsed["source_period_key"],
            parsed["export_period_key"],
            parsed["month"],
            parsed["export_filename"],
            str(export_file),
        )

        if row in seen_rows:
            continue
        seen_rows.add(row)
        rows_to_keep.append(row)

    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(EXPORT_HEADER)
        writer.writerows(rows_to_keep)

    print(f"Scanned export dir: {export_dir}")
    print(f"Rebuilt export csv: {output_csv}")
    print(f"Rows kept: {len(rows_to_keep)}")
    if skipped_files:
        print(f"Skipped files: {len(skipped_files)}")
        for file_name, reason in skipped_files[:20]:
            print(f"  - {file_name} | {reason}")
        if len(skipped_files) > 20:
            print(f"  ... {len(skipped_files) - 20} more skipped")

    return output_csv


def _build_cli_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Rebuild the Xiangyuan export progress CSV from exported Excel files."
    )
    parser.add_argument(
        "--source-dir",
        default=str(EXPORT_DIR),
        help="Directory containing exported Excel files.",
    )
    parser.add_argument(
        "--output-csv",
        default="",
        help="Optional explicit output CSV path. Defaults to runtime export_progress.csv.",
    )
    parser.add_argument(
        "--source-period-key",
        default=TARGET_SOURCE_PERIOD_KEY,
        help="Source period key for rebuilt rows. Defaults to 2024.01-2024.12.",
    )
    return parser


def main() -> None:
    parser = _build_cli_parser()
    args = parser.parse_args()

    output_csv = Path(args.output_csv) if args.output_csv else None
    rebuild_export_csv_from_export_dir(
        export_dir=Path(args.source_dir),
        output_csv=output_csv,
        source_period_key=args.source_period_key,
    )


if __name__ == "__main__":
    main()
