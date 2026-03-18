from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import csv
import re
from typing import Iterable

from modules._shared.logger import get_export_csv_path


DEFAULT_REQUIRED_TABLES = [
    "\u79d1\u76ee\u4f59\u989d\u8868",
    "\u51ed\u8bc1\u8868",
    "\u5916\u5e01\u4f59\u989d\u8868",
    "\u6838\u7b97\u9879\u76ee\u4f59\u989d\u8868",
    "\u6838\u7b97\u9879\u76ee\u660e\u7ec6\u8868",
]
YEAR_NODE_PERIOD_RE = re.compile(
    r"(?P<start_year>\d{4})\s*\u5e74\s*(?P<start_month>\d{1,2})\s*\u6708\s*[\u2014\-~\u81f3]+\s*"
    r"(?P<end_year>\d{4})\s*\u5e74\s*(?P<end_month>\d{1,2})\s*\u6708"
)
PERIOD_KEY_RE = re.compile(
    r"^(?P<start_year>\d{4})\.(?P<start_month>\d{2})-(?P<end_year>\d{4})\.(?P<end_month>\d{2})$"
)


def normalize_text(value: str | None) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_year(value: str | int | None) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    if text.endswith("\u5e74"):
        text = text[:-1].strip()
    return text


def normalize_company_name(company_name: str | None) -> str:
    text = normalize_text(company_name)
    text = text.replace(chr(0xFF08), "(").replace(chr(0xFF09), ")")
    text = re.sub(r"[()]", "", text)
    text = re.sub(r"_+$", "", text).rstrip()
    return text


def normalize_table_name(table_name: str | None) -> str:
    text = normalize_text(table_name)
    text = text.replace("\uFF08\u57fa\u672c\u8868\uFF09", "")
    text = text.replace("(\u57fa\u672c\u8868)", "")
    text = text.replace("\u57fa\u672c\u8868", "")
    text = text.strip(" -_()\uFF08\uFF09")
    return text


def build_period_key(start_year: str | int, start_month: str | int, end_year: str | int, end_month: str | int) -> str:
    return f"{int(start_year):04d}.{int(start_month):02d}-{int(end_year):04d}.{int(end_month):02d}"


def normalize_period_key(value: str | None) -> str:
    text = normalize_text(value)
    if not text:
        return ""

    match = PERIOD_KEY_RE.match(text)
    if match:
        return build_period_key(
            match.group("start_year"),
            match.group("start_month"),
            match.group("end_year"),
            match.group("end_month"),
        )

    match = YEAR_NODE_PERIOD_RE.search(text)
    if match:
        return build_period_key(
            match.group("start_year"),
            match.group("start_month"),
            match.group("end_year"),
            match.group("end_month"),
        )

    return ""


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


def normalize_year_node_text(value: str | None) -> str:
    text = normalize_text(value)
    if not text:
        return ""

    period_key = normalize_period_key(text)
    if period_key:
        return format_period_key_as_year_node(period_key)

    return text


def load_export_records(csv_path: Path | None = None) -> list[dict]:
    if csv_path is None:
        csv_path = get_export_csv_path()

    if not csv_path.exists():
        return []

    rows: list[dict] = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            table_name = normalize_table_name(row.get("table_name", ""))
            company_name = normalize_company_name(row.get("company_name", ""))
            year = normalize_year(row.get("year", ""))
            period_key = normalize_period_key(row.get("period_key", ""))
            year_node_text = normalize_year_node_text(row.get("year_node_text", ""))

            if period_key and not year:
                year = period_key[:4]
            if year_node_text and not period_key:
                period_key = normalize_period_key(year_node_text)
            if period_key and not year_node_text:
                year_node_text = format_period_key_as_year_node(period_key)

            rows.append({
                "table_name": table_name,
                "company_name": company_name,
                "year": year,
                "period_key": period_key,
                "year_node_text": year_node_text,
            })
    return rows


def build_export_indexes(records: list[dict]) -> tuple[dict[tuple[str, str], set[str]], dict[tuple[str, str], set[str]]]:
    period_index: dict[tuple[str, str], set[str]] = defaultdict(set)
    year_index: dict[tuple[str, str], set[str]] = defaultdict(set)

    for row in records:
        company_name = normalize_company_name(row.get("company_name", ""))
        year = normalize_year(row.get("year", ""))
        period_key = normalize_period_key(row.get("period_key", ""))
        table_name = normalize_table_name(row.get("table_name", ""))

        if not company_name or not year or not table_name:
            continue

        year_index[(company_name, year)].add(table_name)
        if period_key:
            period_index[(company_name, period_key)].add(table_name)

    return period_index, year_index


def get_exported_tables_for_target(
    company_name: str,
    year: str | int,
    year_node_text: str | None = None,
    csv_path: Path | None = None,
) -> set[str]:
    records = load_export_records(csv_path)
    period_index, year_index = build_export_indexes(records)

    company_key = normalize_company_name(company_name)
    year_key = normalize_year(year)
    period_key = normalize_period_key(year_node_text)

    if period_key:
        return period_index.get((company_key, period_key), set())

    return year_index.get((company_key, year_key), set())


def is_target_fully_exported(
    company_name: str,
    year: str | int,
    year_node_text: str | None = None,
    required_tables: Iterable[str] | None = None,
    csv_path: Path | None = None,
) -> bool:
    if required_tables is None:
        required_tables = DEFAULT_REQUIRED_TABLES

    required = {normalize_table_name(x) for x in required_tables if normalize_table_name(x)}
    exported = get_exported_tables_for_target(company_name, year, year_node_text, csv_path)
    return required.issubset(exported)


def get_missing_tables_for_target(
    company_name: str,
    year: str | int,
    year_node_text: str | None = None,
    required_tables: Iterable[str] | None = None,
    csv_path: Path | None = None,
) -> list[str]:
    if required_tables is None:
        required_tables = DEFAULT_REQUIRED_TABLES

    required = {normalize_table_name(x) for x in required_tables if normalize_table_name(x)}
    exported = get_exported_tables_for_target(company_name, year, year_node_text, csv_path)
    return sorted(required - exported)


def should_skip_node_task(
    company_name: str,
    year: str | int,
    year_node_text: str | None = None,
    required_tables: Iterable[str] | None = None,
    csv_path: Path | None = None,
) -> bool:
    return is_target_fully_exported(
        company_name=company_name,
        year=year,
        year_node_text=year_node_text,
        required_tables=required_tables,
        csv_path=csv_path,
    )


def explain_skip_decision(
    company_name: str,
    year: str | int,
    year_node_text: str | None = None,
    required_tables: Iterable[str] | None = None,
    csv_path: Path | None = None,
) -> dict:
    if required_tables is None:
        required_tables = DEFAULT_REQUIRED_TABLES

    exported = sorted(get_exported_tables_for_target(company_name, year, year_node_text, csv_path))
    missing = get_missing_tables_for_target(company_name, year, year_node_text, required_tables, csv_path)
    should_skip = len(missing) == 0

    return {
        "company_name": normalize_company_name(company_name),
        "year": normalize_year(year),
        "year_node_text": normalize_year_node_text(year_node_text),
        "period_key": normalize_period_key(year_node_text),
        "should_skip": should_skip,
        "exported_tables": exported,
        "missing_tables": missing,
    }


def main():
    company_name = "\u676d\u5dde\u6c11\u7f6e\u6295\u8d44\u7ba1\u7406\u6709\u9650\u516c\u53f8"
    year = "2021"
    year_node_text = "\u0032\u0030\u0032\u0031\u5e74\u0031\u6708\u2014\u0032\u0030\u0032\u0031\u5e74\u0035\u6708"

    result = explain_skip_decision(company_name, year, year_node_text=year_node_text)
    print(result)


if __name__ == "__main__":
    main()
