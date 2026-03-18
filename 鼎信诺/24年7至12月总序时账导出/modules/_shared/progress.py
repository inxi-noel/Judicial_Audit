from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable

from modules._shared.config import TARGET_MONTHS, TARGET_SOURCE_PERIOD_KEY
from modules._shared.logger import get_export_csv_path


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_company_name(company_name) -> str:
    text = normalize_text(company_name)
    text = text.replace("\uff08", "(").replace("\uff09", ")")
    text = text.rstrip("_")
    return text


def normalize_period_key(value) -> str:
    return normalize_text(value)


def normalize_month(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    return str(int(text))


def load_export_records(csv_path: Path | None = None) -> list[dict]:
    csv_path = csv_path or get_export_csv_path()
    if not csv_path.exists():
        return []

    rows: list[dict] = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            rows.append(
                {
                    "report_name": normalize_text(row.get("report_name", "")),
                    "company_name": normalize_company_name(row.get("company_name", "")),
                    "source_period_key": normalize_period_key(row.get("source_period_key", "")),
                    "export_period_key": normalize_period_key(row.get("export_period_key", "")),
                    "month": normalize_month(row.get("month", "")),
                    "export_filename": normalize_text(row.get("export_filename", "")),
                    "export_path": normalize_text(row.get("export_path", "")),
                }
            )
    return rows


def get_exported_months_for_target(
    company_name: str,
    source_period_key: str,
    csv_path: Path | None = None,
) -> set[int]:
    company_key = normalize_company_name(company_name)
    period_key = normalize_period_key(source_period_key) or TARGET_SOURCE_PERIOD_KEY
    exported: set[int] = set()

    for row in load_export_records(csv_path):
        if row["company_name"] != company_key:
            continue
        if row["source_period_key"] != period_key:
            continue
        if not row["month"]:
            continue
        exported.add(int(row["month"]))

    return exported


def is_target_fully_exported(
    company_name: str,
    source_period_key: str,
    required_months: Iterable[int] | None = None,
    csv_path: Path | None = None,
) -> bool:
    required = set(required_months or TARGET_MONTHS)
    exported = get_exported_months_for_target(company_name, source_period_key, csv_path)
    return required.issubset(exported)


def get_missing_months_for_target(
    company_name: str,
    source_period_key: str,
    required_months: Iterable[int] | None = None,
    csv_path: Path | None = None,
) -> list[int]:
    required = set(required_months or TARGET_MONTHS)
    exported = get_exported_months_for_target(company_name, source_period_key, csv_path)
    return sorted(required - exported)


def explain_skip_decision(
    company_name: str,
    source_period_key: str,
    required_months: Iterable[int] | None = None,
    csv_path: Path | None = None,
) -> dict:
    required = sorted(set(required_months or TARGET_MONTHS))
    exported = sorted(get_exported_months_for_target(company_name, source_period_key, csv_path))
    missing = get_missing_months_for_target(company_name, source_period_key, required, csv_path)
    return {
        "company_name": normalize_company_name(company_name),
        "source_period_key": normalize_period_key(source_period_key) or TARGET_SOURCE_PERIOD_KEY,
        "should_skip": len(missing) == 0,
        "required_months": required,
        "exported_months": exported,
        "missing_months": missing,
    }
