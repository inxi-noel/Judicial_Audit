from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
import re
from typing import Any

from modules._shared.config import PROJECT_LIST_CLASS_NAME, PROJECT_LIST_TITLE_RE
from modules._shared.logger import get_desktop_log_dir
from modules._shared.ui_helpers import get_rect, get_text, rect_to_tuple, rect_valid
from modules._shared.window_retry import connect_uia_window


PERIOD_NODE_RE = re.compile(
    r"(?P<start_year>\d{4})\s*年\s*(?P<start_month>\d{1,2})\s*月\s*[\-—~至]+\s*"
    r"(?P<end_year>\d{4})\s*年\s*(?P<end_month>\d{1,2})\s*月"
)


def get_tree_observe_csv_path() -> Path:
    return get_desktop_log_dir() / "units_observe.csv"


def get_tree_observe_snapshot_csv_path() -> Path:
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    return get_desktop_log_dir() / f"units_observe_{now}.csv"


def normalize_period_key(raw_value: str | None) -> str:
    text = (raw_value or "").strip()
    if not text:
        return ""

    match = PERIOD_NODE_RE.search(text)
    if not match:
        return ""

    return (
        f"{int(match.group('start_year')):04d}.{int(match.group('start_month')):02d}-"
        f"{int(match.group('end_year')):04d}.{int(match.group('end_month')):02d}"
    )


def is_period_text(text: str) -> bool:
    return bool(normalize_period_key(text))


def safe_rect_tuple(ctrl: Any) -> tuple[int | str, int | str, int | str, int | str]:
    rect = get_rect(ctrl)
    if not rect_valid(rect):
        return "", "", "", ""
    return rect_to_tuple(rect)


def connect_project_list_window(
    title_re: str = PROJECT_LIST_TITLE_RE,
    class_name: str = PROJECT_LIST_CLASS_NAME,
):
    win = connect_uia_window(
        title_re=title_re,
        class_name=class_name,
        action_name="Connect project list window",
    )
    if win is None:
        raise RuntimeError("Project list window not found")
    return win


def scan_project_tree_items(win) -> list:
    return win.descendants(control_type="TreeItem")


def build_tree_observe_rows(items: list) -> list[dict]:
    rows: list[dict] = []

    project_name = None
    period_key = ""
    period_node_text = ""
    level3_name = ""
    sequence = 0

    for raw_index, ctrl in enumerate(items, 1):
        text = get_text(ctrl)
        if not text:
            continue

        left, top, right, bottom = safe_rect_tuple(ctrl)
        current_period_key = normalize_period_key(text)

        if project_name is None:
            project_name = text
            role_guess = "L1_ROOT"
            level = 1
            parent_name = ""
            path_key = project_name
            parent_path_key = ""
            company_name = ""
            level3_name = ""
            is_period_node = 0
            is_first_after_period = 0
        elif current_period_key:
            period_key = current_period_key
            period_node_text = text
            role_guess = "L2_PERIOD"
            level = 2
            parent_name = project_name
            path_key = f"{project_name}|{period_key}"
            parent_path_key = project_name
            company_name = ""
            level3_name = ""
            is_period_node = 1
            is_first_after_period = 0
        elif not period_key:
            role_guess = "UNRESOLVED_BEFORE_PERIOD"
            level = ""
            parent_name = ""
            path_key = f"{project_name}|UNRESOLVED|{text}"
            parent_path_key = ""
            company_name = ""
            is_period_node = 0
            is_first_after_period = 0
        elif not level3_name:
            level3_name = text
            role_guess = "L3_GUESS"
            level = 3
            parent_name = period_node_text
            path_key = f"{project_name}|{period_key}|{text}"
            parent_path_key = f"{project_name}|{period_key}"
            company_name = text
            is_period_node = 0
            is_first_after_period = 1
        else:
            role_guess = "L4_GUESS"
            level = 4
            parent_name = level3_name
            path_key = f"{project_name}|{period_key}|{level3_name}|{text}"
            parent_path_key = f"{project_name}|{period_key}|{level3_name}"
            company_name = text
            is_period_node = 0
            is_first_after_period = 0

        sequence += 1
        rows.append(
            {
                "sequence": sequence,
                "raw_index": raw_index,
                "project_name": project_name or "",
                "period_key": period_key,
                "period_node_text": period_node_text,
                "level3_name": level3_name,
                "node_text": text,
                "company_name": company_name,
                "parent_name": parent_name,
                "level": level,
                "role_guess": role_guess,
                "is_period_node": is_period_node,
                "is_first_after_period": is_first_after_period,
                "path_key": path_key,
                "parent_path_key": parent_path_key,
                "left": left,
                "top": top,
                "right": right,
                "bottom": bottom,
            }
        )

    return rows


def write_tree_observe_csv(rows: list[dict], output_csv: Path) -> Path:
    fieldnames = [
        "sequence",
        "raw_index",
        "project_name",
        "period_key",
        "period_node_text",
        "level3_name",
        "node_text",
        "company_name",
        "parent_name",
        "level",
        "role_guess",
        "is_period_node",
        "is_first_after_period",
        "path_key",
        "parent_path_key",
        "left",
        "top",
        "right",
        "bottom",
    ]

    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return output_csv


def export_project_tree_observe_csv(
    output_csv: Path | None = None,
    title_re: str = PROJECT_LIST_TITLE_RE,
    class_name: str = PROJECT_LIST_CLASS_NAME,
) -> Path:
    output_csv = output_csv or get_tree_observe_csv_path()
    win = connect_project_list_window(title_re=title_re, class_name=class_name)
    items = scan_project_tree_items(win)
    rows = build_tree_observe_rows(items)
    result_path = write_tree_observe_csv(rows, output_csv)

    print(f"Tree rows exported: {len(rows)}")
    print(f"Output csv: {result_path}")
    return result_path


if __name__ == "__main__":
    export_project_tree_observe_csv()
