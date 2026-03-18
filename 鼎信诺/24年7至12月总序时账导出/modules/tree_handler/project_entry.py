from __future__ import annotations

import csv
from pathlib import Path
import time

from pywinauto import Desktop
from pywinauto.keyboard import send_keys

from modules._shared.config import (
    ENTER_PROJECT_TEXT,
    MAIN_WINDOW_TITLE_RE,
    PROJECT_LIST_CLASS_NAME,
    PROJECT_LIST_TITLE_RE,
    TARGET_MONTHS,
    TARGET_SOURCE_PERIOD_KEY,
)
from modules._shared.logger import get_desktop_log_dir
from modules._shared.progress import explain_skip_decision
from modules._shared.ui_helpers import get_rect, get_text, is_rect_visible, rect_valid
from modules._shared.window_retry import connect_uia_window
from modules.tree_handler.tree_observer import get_tree_observe_csv_path, is_period_text, normalize_period_key


PAGE_DOWN_MAX = 200
SAME_PAGE_LIMIT = 3
ACTION_SLEEP = 0.6
DEBUG_PREVIEW_COUNT = 30
WAIT_PROJECT_LIST_CLOSE_SEC = 15
WAIT_MAIN_WINDOW_SEC = 20

COL_PROJECT_NAME = "project_name"
COL_PERIOD_KEY = "period_key"
COL_PERIOD_NODE_TEXT = "period_node_text"
COL_LEVEL3_NAME = "level3_name"
COL_COMPANY_NAME = "company_name"
COL_LEVEL = "level"
SKIP_TARGETS_CSV_NAME = "\u9700\u8981\u88ab\u8df3\u8fc7\u7684\u9879\u76ee.csv"
SKIP_COL_ROOT = "\u4e00\u7ea7"
SKIP_COL_PERIOD = "\u4e8c\u7ea7"
SKIP_COL_LEVEL3 = "\u4e09\u7ea7"
SKIP_COL_COMPANY = "\u56db\u7ea7"
ENTER_PROJECT_CONTROL_TYPES = ("Button", "Text")


_SKIP_TARGET_CACHE: dict[str, object] = {"mtime": None, "rows": []}


def get_skip_targets_csv_path() -> Path:
    return get_desktop_log_dir() / SKIP_TARGETS_CSV_NAME


def normalize_skip_period_value(value) -> str:
    text = normalize_text(value)
    text = text.replace("\u5e74", "")
    return text


def load_skip_targets() -> list[dict]:
    csv_path = get_skip_targets_csv_path()
    if not csv_path.exists():
        _SKIP_TARGET_CACHE["mtime"] = None
        _SKIP_TARGET_CACHE["rows"] = []
        return []

    mtime = csv_path.stat().st_mtime
    if _SKIP_TARGET_CACHE["mtime"] == mtime:
        return list(_SKIP_TARGET_CACHE["rows"])

    rows: list[dict] = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            rows.append(
                {
                    "root_name": normalize_text(row.get(SKIP_COL_ROOT, "")),
                    "period_value": normalize_skip_period_value(row.get(SKIP_COL_PERIOD, "")),
                    "level3_name": normalize_text(row.get(SKIP_COL_LEVEL3, "")),
                    "company_name": normalize_text(row.get(SKIP_COL_COMPANY, "")),
                }
            )

    _SKIP_TARGET_CACHE["mtime"] = mtime
    _SKIP_TARGET_CACHE["rows"] = rows
    return rows


def row_matches_skip_target(row: dict, skip_target: dict) -> bool:
    row_root = normalize_text(row.get(COL_PROJECT_NAME, ""))
    row_period_key = normalize_text(row.get(COL_PERIOD_KEY, ""))
    row_period_year = row_period_key.split(".", 1)[0] if row_period_key else ""
    row_period_node = normalize_text(row.get(COL_PERIOD_NODE_TEXT, ""))
    row_level3 = normalize_text(row.get(COL_LEVEL3_NAME, ""))
    row_company = normalize_text(row.get(COL_COMPANY_NAME, ""))

    skip_root = normalize_text(skip_target.get("root_name", ""))
    skip_period = normalize_skip_period_value(skip_target.get("period_value", ""))
    skip_level3 = normalize_text(skip_target.get("level3_name", ""))
    skip_company = normalize_text(skip_target.get("company_name", ""))

    if skip_root and skip_root != row_root:
        return False
    if skip_level3 and skip_level3 != row_level3:
        return False
    if skip_company and skip_company != row_company:
        return False
    if skip_period and skip_period not in {row_period_key, row_period_year, row_period_node}:
        return False

    return True


def should_skip_row_by_csv(row: dict) -> bool:
    for skip_target in load_skip_targets():
        if row_matches_skip_target(row, skip_target):
            print(
                f"Skip by CSV: {normalize_text(row.get(COL_COMPANY_NAME, ''))} | "
                f"period={normalize_text(row.get(COL_PERIOD_KEY, ''))} | "
                f"level3={normalize_text(row.get(COL_LEVEL3_NAME, ''))}"
            )
            return True
    return False


def get_desktop() -> Desktop:
    return Desktop(backend="uia")


def safe_text(ctrl) -> str:
    return get_text(ctrl)


def safe_rect(ctrl):
    return get_rect(ctrl)


def activate_win(win) -> bool:
    try:
        win.set_focus()
        time.sleep(0.3)
        return True
    except Exception:
        pass

    try:
        rect = win.rectangle()
        win.click_input(coords=(max(10, rect.width() // 2), 10))
        time.sleep(0.3)
        return True
    except Exception as exc:
        print("Activate window failed:", exc)
        return False


def connect_project_list_win():
    win = connect_uia_window(
        title_re=PROJECT_LIST_TITLE_RE,
        class_name=PROJECT_LIST_CLASS_NAME,
        action_name="Connect project list window",
    )
    if win is None:
        raise RuntimeError("Project list window not found")
    return win


def try_get_project_list_win():
    desktop = get_desktop()
    win = desktop.window(title_re=PROJECT_LIST_TITLE_RE, class_name=PROJECT_LIST_CLASS_NAME)
    try:
        if win.exists(timeout=1):
            return win
    except Exception:
        pass
    return None


def wait_main_window(timeout: int = WAIT_MAIN_WINDOW_SEC):
    _ = timeout
    return connect_uia_window(title_re=MAIN_WINDOW_TITLE_RE, action_name="Wait for main window")


def wait_project_list_closed(timeout: int = WAIT_PROJECT_LIST_CLOSE_SEC) -> bool:
    start = time.time()
    while time.time() - start < timeout:
        if try_get_project_list_win() is None:
            return True
        time.sleep(0.5)
    return False


def load_tree_rows(csv_path=None) -> list[dict]:
    csv_path = csv_path or get_tree_observe_csv_path()
    rows = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            rows.append(row)
    return rows


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def iter_target_rows(rows: list[dict]):
    seen = set()
    for row in rows:
        company_name = normalize_text(row.get(COL_COMPANY_NAME, ""))
        period_key = normalize_text(row.get(COL_PERIOD_KEY, ""))
        level3_name = normalize_text(row.get(COL_LEVEL3_NAME, ""))
        level = normalize_text(row.get(COL_LEVEL, ""))

        if level != "4":
            continue
        if not company_name or period_key != TARGET_SOURCE_PERIOD_KEY:
            continue

        key = (company_name, level3_name, period_key)
        if key in seen:
            continue

        seen.add(key)
        yield row


def pick_first_pending_row(rows: list[dict]) -> dict | None:
    for row in iter_target_rows(rows):
        company_name = normalize_text(row.get(COL_COMPANY_NAME, ""))

        if should_skip_row_by_csv(row):
            continue

        decision = explain_skip_decision(
            company_name=company_name,
            source_period_key=TARGET_SOURCE_PERIOD_KEY,
            required_months=TARGET_MONTHS,
        )

        if decision["should_skip"]:
            print(
                f"Skip: {company_name} | period={TARGET_SOURCE_PERIOD_KEY} | "
                f"exported={decision['exported_months']}"
            )
            continue

        print(
            f"Pending: {company_name} | period={TARGET_SOURCE_PERIOD_KEY} | "
            f"missing={decision['missing_months']}"
        )
        return row

    print("No pending task found: all targets for the current batch are complete")
    return None


def visible_items(win, debug: bool = False) -> list:
    nodes = win.descendants(control_type="TreeItem")
    visible = []

    for ctrl in nodes:
        text = safe_text(ctrl)
        if not text:
            continue
        if is_rect_visible(ctrl):
            visible.append(ctrl)

    visible.sort(
        key=lambda ctrl: (
            safe_rect(ctrl).top if safe_rect(ctrl) else 999999,
            safe_rect(ctrl).left if safe_rect(ctrl) else 999999,
        )
    )

    if debug:
        print(f"Visible tree items: {len(visible)}")
        for index, ctrl in enumerate(visible[:DEBUG_PREVIEW_COUNT], 1):
            print(f"  [{index:02d}] {safe_text(ctrl)}")

    return visible


def page_signature(win) -> tuple:
    signature = []
    for ctrl in visible_items(win, debug=False):
        text = safe_text(ctrl)
        rect = safe_rect(ctrl)
        top = rect.top if rect else ""
        signature.append((text, top))
    return tuple(signature)


def single_click(ctrl, label: str = "") -> bool:
    try:
        ctrl.click_input()
        time.sleep(ACTION_SLEEP)
        if label:
            print(f"Single click OK: {label}")
        return True
    except Exception as exc:
        if label:
            print(f"Single click failed: {label} | {exc}")
        return False


def verify_tree_item_selected(ctrl, label: str = "") -> bool:
    try:
        selected = bool(ctrl.is_selected())
    except Exception as exc:
        if label:
            print(f"Read selected state failed: {label} | {exc}")
        return False

    if selected and label:
        print(f"Selection verified: {label}")
    return selected


def ensure_tree_item_selected(ctrl, label: str = "") -> bool:
    try:
        ctrl.set_focus()
        time.sleep(0.2)
    except Exception:
        pass

    try:
        ctrl.select()
        time.sleep(ACTION_SLEEP)
    except Exception:
        pass

    if verify_tree_item_selected(ctrl, label):
        return True

    if not single_click(ctrl, label):
        return False

    return verify_tree_item_selected(ctrl, label)


def double_click(ctrl, label: str = "") -> bool:
    try:
        ctrl.click_input(double=True)
        time.sleep(ACTION_SLEEP)
        if label:
            print(f"Double click OK: {label}")
        return True
    except Exception as exc:
        if label:
            print(f"Double click failed: {label} | {exc}")
        return False


def send_pagedown(win) -> bool:
    if not activate_win(win):
        return False

    try:
        send_keys("{PGDN}")
        time.sleep(ACTION_SLEEP)
        return True
    except Exception as exc:
        print("Send PageDown failed:", exc)
        return False


def find_root_ctrl(win, project_name: str):
    for ctrl in visible_items(win, debug=True):
        if safe_text(ctrl) == project_name:
            return ctrl
    return None


def find_visible_period(win, period_key: str, period_node_text: str = "", debug: bool = False):
    visible = visible_items(win, debug=debug)
    if period_node_text:
        for ctrl in visible:
            if safe_text(ctrl) == period_node_text:
                return ctrl

    for ctrl in visible:
        text = safe_text(ctrl)
        if is_period_text(text) and normalize_period_key(text) == period_key:
            return ctrl
    return None


def is_collapsed_to_root_only(win, project_name: str) -> bool:
    visible = visible_items(win, debug=True)
    return len(visible) == 1 and safe_text(visible[0]) == project_name


def collapse_to_root_only(win, project_name: str, max_try: int = 3) -> bool:
    for _ in range(max_try):
        if is_collapsed_to_root_only(win, project_name):
            return True

        root = find_root_ctrl(win, project_name)
        if root is None:
            return False

        if not double_click(root, f"Collapse root: {project_name}"):
            return False

        time.sleep(0.8)
        if is_collapsed_to_root_only(win, project_name):
            return True

    return False


def ensure_period_expanded(win, project_name: str, period_key: str, period_node_text: str = ""):
    root = find_root_ctrl(win, project_name)
    if root is None:
        print("Failed: root project node not found")
        return None

    period_ctrl = find_visible_period(win, period_key, period_node_text=period_node_text, debug=True)
    if period_ctrl is None:
        if not double_click(root, f"Expand root: {project_name}"):
            return None
        period_ctrl = find_visible_period(win, period_key, period_node_text=period_node_text, debug=True)
        if period_ctrl is None:
            print(f"Failed: period node not found after expanding root -> {period_key}")
            return None

    if not double_click(period_ctrl, f"Expand period: {period_node_text or period_key}"):
        return None
    return period_ctrl


def find_level3_after_period(win, period_key: str, level3_name: str, period_node_text: str = ""):
    visible = visible_items(win, debug=True)
    period_index = -1

    for index, ctrl in enumerate(visible):
        text = safe_text(ctrl)
        if period_node_text and text == period_node_text:
            period_index = index
            break
        if is_period_text(text) and normalize_period_key(text) == period_key:
            period_index = index
            break

    if period_index == -1:
        print("Failed: target period node is not visible on current page")
        return None

    for index in range(period_index + 1, len(visible)):
        text = safe_text(visible[index])
        if is_period_text(text):
            break
        if text == level3_name:
            return visible[index]

    for index in range(period_index + 1, len(visible)):
        text = safe_text(visible[index])
        if is_period_text(text):
            break
        return visible[index]

    print("Failed: no level-3 candidate found after target period node")
    return None


def find_level4_by_pagedown(win, level4_name: str):
    visible = visible_items(win, debug=True)
    for ctrl in visible:
        if safe_text(ctrl) == level4_name:
            return ctrl

    previous_signature = page_signature(win)
    same_count = 0

    for _ in range(PAGE_DOWN_MAX):
        if not send_pagedown(win):
            return None

        visible = visible_items(win, debug=True)
        for ctrl in visible:
            if safe_text(ctrl) == level4_name:
                return ctrl

        current_signature = page_signature(win)
        if current_signature == previous_signature:
            same_count += 1
        else:
            same_count = 0

        if same_count >= SAME_PAGE_LIMIT:
            return None
        previous_signature = current_signature

    return None


def find_enter_project_control(win):
    for control_type in ENTER_PROJECT_CONTROL_TYPES:
        try:
            ctrl = win.child_window(title=ENTER_PROJECT_TEXT, control_type=control_type).wrapper_object()
        except Exception:
            ctrl = None

        if ctrl is None:
            continue

        try:
            rect = safe_rect(ctrl)
            if rect_valid(rect) and rect.width() > 3 and rect.height() > 3:
                return ctrl
        except Exception:
            continue

    candidates = []
    try:
        all_controls = win.descendants()
    except Exception as exc:
        print("Scan enter project controls failed:", exc)
        return None

    for ctrl in all_controls:
        try:
            if safe_text(ctrl) != ENTER_PROJECT_TEXT:
                continue
            rect = safe_rect(ctrl)
            if not rect_valid(rect):
                continue
            if rect.width() <= 3 or rect.height() <= 3:
                continue
            candidates.append(ctrl)
        except Exception:
            continue

    return candidates[0] if candidates else None


def click_enter_project(win) -> bool:
    ctrl = find_enter_project_control(win)
    if ctrl is None:
        return False

    try:
        ctrl.click_input()
        time.sleep(0.8)
        return True
    except Exception:
        pass

    try:
        rect = ctrl.rectangle()
        x = (rect.left + rect.right) // 2
        y = (rect.top + rect.bottom) // 2
        base_rect = win.rectangle()
        win.click_input(coords=(x - base_rect.left, y - base_rect.top))
        time.sleep(0.8)
        return True
    except Exception as exc:
        print("Coordinate click enter project failed:", exc)
        return False


def select_tree_target(win, row: dict) -> bool:
    project_name = normalize_text(row.get(COL_PROJECT_NAME, ""))
    period_key = normalize_text(row.get(COL_PERIOD_KEY, ""))
    period_node_text = normalize_text(row.get(COL_PERIOD_NODE_TEXT, ""))
    level3_name = normalize_text(row.get(COL_LEVEL3_NAME, ""))
    unit_name = normalize_text(row.get(COL_COMPANY_NAME, ""))

    try:
        level_value = int(normalize_text(row.get(COL_LEVEL, "")))
    except Exception:
        print("Failed: CSV level value is not a valid integer")
        return False

    if not activate_win(win):
        return False
    if not collapse_to_root_only(win, project_name):
        return False
    if ensure_period_expanded(win, project_name, period_key, period_node_text=period_node_text) is None:
        return False

    if level_value == 4:
        level3_ctrl = find_level3_after_period(win, period_key, level3_name, period_node_text=period_node_text)
        if level3_ctrl is None:
            return False
        if not double_click(level3_ctrl, f"Expand level-3: {level3_name}"):
            return False

        level4_ctrl = find_level4_by_pagedown(win, unit_name)
        if level4_ctrl is None:
            return False
        return ensure_tree_item_selected(level4_ctrl, f"Select level-4: {unit_name}")

    print(f"Unsupported level value: {level_value}; only level-4 targets are allowed")
    return False


def enter_first_pending_project() -> dict | None:
    rows = load_tree_rows()
    target_row = pick_first_pending_row(rows)
    if target_row is None:
        return None

    project_list_win = connect_project_list_win()
    if not select_tree_target(project_list_win, target_row):
        return None
    if not click_enter_project(project_list_win):
        return None
    if not wait_project_list_closed():
        return None

    main_win = wait_main_window()
    if main_win is None:
        return None

    print("Entered project main window ->", repr(main_win.window_text()))
    return target_row


if __name__ == "__main__":
    row = enter_first_pending_project()
    if row is None:
        print("No project entered")
    else:
        print("Project entered:", row)
