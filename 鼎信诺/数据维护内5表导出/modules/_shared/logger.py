import argparse
import csv
from datetime import datetime
from pathlib import Path
import re


LOG_PREFIX = "automation_log"
EXPORT_PREFIX = "exported_tables"
ACTIVE_BATCH_MARKER_NAME = "active_batch_stamp.txt"
EXPORT_HEADER = ["table_name", "company_name", "year", "period_key", "year_node_text"]
BATCH_STAMP_RE = re.compile(r"^\d{8}$")
BASIC_TABLE_SUFFIX_RE = re.compile(r"[\(（]\s*基本表\s*[\)）]")
BASIC_TABLE_TRAILING_RE = re.compile(r"基本表$")
PERIOD_KEY_RE = re.compile(
    r"^(?P<start_year>\d{4})\.(?P<start_month>\d{2})-(?P<end_year>\d{4})\.(?P<end_month>\d{2})$"
)
PERIOD_SUFFIX_RE = re.compile(
    r"[\(（]?(?P<period>\d{4}\.\d{2}-\d{4}\.\d{2})[\)）]?$"
)


DEFAULT_EXPORT_SOURCE_DIR_NAME = "testing1"


def get_desktop_log_dir() -> Path:
    desktop = Path.home() / "Desktop"
    log_dir = desktop / "鼎信诺导账套日志"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def get_default_export_source_dir() -> Path:
    return Path.home() / "Desktop" / DEFAULT_EXPORT_SOURCE_DIR_NAME


def _get_active_batch_marker_path() -> Path:
    return get_desktop_log_dir() / ACTIVE_BATCH_MARKER_NAME


def _list_batch_files(prefix: str, suffix: str) -> list[Path]:
    return sorted(get_desktop_log_dir().glob(f"{prefix}_*{suffix}"))


def _extract_batch_stamp(path: Path, prefix: str, suffix: str) -> str | None:
    pattern = re.compile(rf"^{re.escape(prefix)}_(\d{{8}}){re.escape(suffix)}$")
    match = pattern.match(path.name)
    if not match:
        return None
    return match.group(1)


def _has_any_batch_files() -> bool:
    return bool(_list_batch_files(LOG_PREFIX, ".txt") or _list_batch_files(EXPORT_PREFIX, ".csv"))


def _batch_stamp_has_files(stamp: str) -> bool:
    log_dir = get_desktop_log_dir()
    log_path = log_dir / f"{LOG_PREFIX}_{stamp}.txt"
    export_path = log_dir / f"{EXPORT_PREFIX}_{stamp}.csv"
    return log_path.exists() or export_path.exists()


def _count_export_rows(csv_path: Path) -> int:
    if not csv_path.exists():
        return 0

    try:
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            return sum(1 for _ in reader)
    except Exception:
        try:
            with open(csv_path, "r", encoding="utf-8-sig") as f:
                return max(0, sum(1 for _ in f) - 1)
        except Exception:
            return 0


def _pick_best_existing_batch_stamp() -> str | None:
    export_candidates: list[tuple[int, int, float, str]] = []
    for csv_path in _list_batch_files(EXPORT_PREFIX, ".csv"):
        stamp = _extract_batch_stamp(csv_path, EXPORT_PREFIX, ".csv")
        if not stamp:
            continue

        stat = csv_path.stat()
        export_candidates.append((
            _count_export_rows(csv_path),
            stat.st_size,
            stat.st_mtime,
            stamp,
        ))

    if export_candidates:
        export_candidates.sort(reverse=True)
        return export_candidates[0][3]

    log_candidates: list[tuple[float, str]] = []
    for log_path in _list_batch_files(LOG_PREFIX, ".txt"):
        stamp = _extract_batch_stamp(log_path, LOG_PREFIX, ".txt")
        if not stamp:
            continue
        log_candidates.append((log_path.stat().st_mtime, stamp))

    if log_candidates:
        log_candidates.sort(reverse=True)
        return log_candidates[0][1]

    return None


def _read_active_batch_stamp() -> str | None:
    marker_path = _get_active_batch_marker_path()
    if not marker_path.exists():
        return None

    try:
        stamp = marker_path.read_text(encoding="utf-8").strip()
    except Exception:
        return None

    if not BATCH_STAMP_RE.match(stamp):
        return None

    return stamp


def _write_active_batch_stamp(stamp: str) -> None:
    _get_active_batch_marker_path().write_text(stamp, encoding="utf-8")


def get_active_batch_stamp() -> str:
    stamp = _read_active_batch_stamp()
    if stamp and (_batch_stamp_has_files(stamp) or not _has_any_batch_files()):
        return stamp

    stamp = _pick_best_existing_batch_stamp()
    if not stamp:
        stamp = datetime.now().strftime("%Y%m%d")

    _write_active_batch_stamp(stamp)
    return stamp


def get_log_file_path() -> Path:
    stamp = get_active_batch_stamp()
    return get_desktop_log_dir() / f"{LOG_PREFIX}_{stamp}.txt"


def normalize_period_key(raw_value: str | None) -> str:
    if raw_value is None:
        return ""

    text = str(raw_value).strip()
    if not text:
        return ""

    match = PERIOD_KEY_RE.match(text)
    if match:
        return (
            f"{match.group('start_year')}.{int(match.group('start_month')):02d}-"
            f"{match.group('end_year')}.{int(match.group('end_month')):02d}"
        )

    match = PERIOD_SUFFIX_RE.search(text)
    if not match:
        return ""

    return normalize_period_key(match.group("period"))


def format_period_key_as_year_node(period_key: str) -> str:
    normalized = normalize_period_key(period_key)
    if not normalized:
        return ""

    match = PERIOD_KEY_RE.match(normalized)
    if not match:
        return ""

    return (
        f"{match.group('start_year')}\u5e74{int(match.group('start_month'))}\u6708\u2014"
        f"{match.group('end_year')}\u5e74{int(match.group('end_month'))}\u6708"
    )


def _canonicalize_export_row(row: dict) -> tuple[str, str, str, str, str] | None:
    table_name = str(row.get("table_name", "")).strip()
    company_name = str(row.get("company_name", "")).strip()
    year = str(row.get("year", "")).strip()
    period_key = normalize_period_key(row.get("period_key", ""))
    year_node_text = str(row.get("year_node_text", "")).strip()

    if period_key and not year:
        year = period_key[:4]
    if period_key and not year_node_text:
        year_node_text = format_period_key_as_year_node(period_key)

    if not table_name or not company_name or not year:
        return None

    return (table_name, company_name, year, period_key, year_node_text)


def _rewrite_export_csv_with_current_header(csv_file: Path) -> None:
    rows_to_keep: list[tuple[str, str, str, str, str]] = []
    seen: set[tuple[str, str, str, str, str]] = set()

    with open(csv_file, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            normalized_row = _canonicalize_export_row(row)
            if normalized_row is None or normalized_row in seen:
                continue
            seen.add(normalized_row)
            rows_to_keep.append(normalized_row)

    with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(EXPORT_HEADER)
        writer.writerows(rows_to_keep)


def _iter_export_rows(csv_path: Path):
    if not csv_path.exists():
        return

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            normalized_row = _canonicalize_export_row(row)
            if normalized_row is None:
                continue
            yield normalized_row


def _merge_split_export_csvs(active_csv_path: Path) -> None:
    existing_rows = set(_iter_export_rows(active_csv_path) or [])
    rows_to_append: list[tuple[str, str, str, str, str]] = []

    other_files = [
        csv_path
        for csv_path in _list_batch_files(EXPORT_PREFIX, ".csv")
        if csv_path != active_csv_path
    ]
    other_files.sort(key=lambda path: path.stat().st_mtime)

    for csv_path in other_files:
        for row in _iter_export_rows(csv_path) or []:
            if row in existing_rows:
                continue
            existing_rows.add(row)
            rows_to_append.append(row)

    if not rows_to_append:
        return

    ensure_export_csv_header(active_csv_path)
    with open(active_csv_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerows(rows_to_append)

    print(f"Merged {len(rows_to_append)} export rows into active batch: {active_csv_path.name}")


def get_export_csv_path() -> Path:
    stamp = get_active_batch_stamp()
    csv_path = get_desktop_log_dir() / f"{EXPORT_PREFIX}_{stamp}.csv"
    _merge_split_export_csvs(csv_path)
    return csv_path


def write_log(step: str, status: str, message: str = "") -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}] [{status}] {step}"
    if message:
        line += f" | {message}"

    log_file = get_log_file_path()
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line + "\n")

    print(line)


def ensure_export_csv_header(csv_file: Path | None = None) -> None:
    if csv_file is None:
        csv_file = get_export_csv_path()

    if csv_file.exists() and csv_file.stat().st_size > 0:
        with open(csv_file, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or []

        if fieldnames != EXPORT_HEADER:
            _rewrite_export_csv_with_current_header(csv_file)
        return

    with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(EXPORT_HEADER)


def normalize_filename(raw_filename: str) -> str:
    if raw_filename is None:
        raise ValueError("Export filename is empty")

    filename = str(raw_filename).strip()
    filename = re.sub(r"\.(xls|xlsx|csv)$", "", filename, flags=re.IGNORECASE)

    if not filename:
        raise ValueError("Export filename is empty after normalization")

    return filename


def parse_export_filename(filename: str) -> tuple[str, str, str, str, str]:
    filename = normalize_filename(filename)

    period_match = PERIOD_SUFFIX_RE.search(filename)
    if not period_match:
        raise ValueError(f"Cannot parse export year from filename: {filename}")

    period_key = normalize_period_key(period_match.group("period"))
    if not period_key:
        raise ValueError(f"Cannot normalize export period from filename: {filename}")

    year = period_key[:4]
    year_node_text = format_period_key_as_year_node(period_key)
    prefix = filename[:period_match.start()].rstrip(" -")
    if not prefix:
        raise ValueError(f"Filename prefix is empty: {filename}")

    if "--" not in prefix:
        raise ValueError(f"Filename missing '--' separator: {filename}")

    table_name_raw, company_name = prefix.split("--", 1)
    table_name_raw = table_name_raw.strip()
    company_name = company_name.strip()

    if not table_name_raw:
        raise ValueError(f"Table name is empty: {filename}")
    if not company_name:
        raise ValueError(f"Company name is empty: {filename}")

    table_name = BASIC_TABLE_SUFFIX_RE.sub("", table_name_raw)
    table_name = BASIC_TABLE_TRAILING_RE.sub("", table_name)
    table_name = table_name.strip(" -_()??")

    if not table_name:
        raise ValueError(f"Table name is empty after cleanup: {filename}")

    return table_name, company_name, year, period_key, year_node_text


def append_export_record(export_filename: str) -> None:
    csv_file = get_export_csv_path()
    ensure_export_csv_header(csv_file)

    table_name, company_name, year, period_key, year_node_text = parse_export_filename(export_filename)

    with open(csv_file, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([table_name, company_name, year, period_key, year_node_text])

    print(
        f"已写入导出记录: "
        f"{table_name} | {company_name} | {year_node_text or year}"
    )


def write_export_success(step: str, export_filename: str, message: str = "") -> None:
    append_export_record(export_filename)
    log_message = f"{message} | {export_filename}" if message else export_filename
    write_log(step, "OK", log_message)



def _iter_export_files(export_dir: Path):
    patterns = ("*.xls", "*.xlsx", "*.csv")
    seen: set[Path] = set()

    for pattern in patterns:
        for path in export_dir.rglob(pattern):
            if not path.is_file():
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            yield path


def rebuild_export_csv_from_export_dir(
    export_dir: Path | None = None,
    output_csv: Path | None = None,
) -> Path:
    if export_dir is None:
        export_dir = get_default_export_source_dir()
    else:
        export_dir = Path(export_dir)

    if not export_dir.exists():
        raise FileNotFoundError(f"Export source directory not found: {export_dir}")
    if not export_dir.is_dir():
        raise NotADirectoryError(f"Export source path is not a directory: {export_dir}")

    if output_csv is None:
        output_csv = get_desktop_log_dir() / f"{EXPORT_PREFIX}_{get_active_batch_stamp()}.csv"
    else:
        output_csv = Path(output_csv)

    rows_to_keep: list[tuple[str, str, str, str, str]] = []
    seen_rows: set[tuple[str, str, str, str, str]] = set()
    skipped_files: list[tuple[str, str]] = []

    for export_file in sorted(_iter_export_files(export_dir), key=lambda path: str(path).lower()):
        try:
            row = parse_export_filename(export_file.name)
        except Exception as exc:
            skipped_files.append((export_file.name, str(exc)))
            continue

        if row in seen_rows:
            continue

        seen_rows.add(row)
        rows_to_keep.append(row)

    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
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
        description="Rebuild exported_tables CSV from exported account files."
    )
    parser.add_argument(
        "--source-dir",
        default=str(get_default_export_source_dir()),
        help="Directory containing exported account files. Defaults to Desktop/testing1.",
    )
    parser.add_argument(
        "--output-csv",
        default="",
        help="Optional explicit output CSV path. Defaults to current active exported_tables CSV.",
    )
    return parser


def main() -> None:
    parser = _build_cli_parser()
    args = parser.parse_args()

    output_csv = Path(args.output_csv) if args.output_csv else None
    rebuild_export_csv_from_export_dir(
        export_dir=Path(args.source_dir),
        output_csv=output_csv,
    )


if __name__ == "__main__":
    main()