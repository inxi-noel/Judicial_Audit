from __future__ import annotations

import csv
import time
from pathlib import Path

from pywinauto.keyboard import send_keys

from modules._shared.config import (
    APP_MENU_BAR_TEXT,
    CHANGE_PROJECT_DOWN_COUNT,
    MAIN_WINDOW_TITLE_RE,
    PROJECT_MANAGEMENT_MENU_INDEX,
    PROJECT_MANAGEMENT_TEXT,
)
from modules._shared.window_retry import connect_uia_window, retry_window_operation
from modules.tree_handler.project_entry import (
    activate_win,
    connect_project_list_win,
    double_click,
    page_signature,
    safe_text,
    visible_items,
)
from modules.tree_handler.tree_observer import get_tree_observe_csv_path, is_period_text


PAGEUP_SLEEP = 0.8
MAX_PAGEUP_TIMES = 100
MAX_COLLAPSE_TRIES = 3


def load_root_project_name_from_csv(csv_path: Path | None = None) -> str:
    csv_path = csv_path or get_tree_observe_csv_path()
    if not csv_path.exists():
        raise FileNotFoundError(f"Tree observe CSV not found: {csv_path}")

    project_names = []
    seen = set()
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            name = str(row.get("project_name", "")).strip()
            if not name or name in seen:
                continue
            seen.add(name)
            project_names.append(name)

    if not project_names:
        raise RuntimeError("No root project name found in tree observe CSV")
    if len(project_names) > 1:
        raise RuntimeError(f"Multiple root project names found: {project_names}")
    return project_names[0]


def connect_main_window():
    win = connect_uia_window(title_re=MAIN_WINDOW_TITLE_RE, action_name="Connect main window")
    if win is None:
        raise RuntimeError("Main window not found")
    return win


def open_change_project_dialog() -> bool:
    try:
        main = connect_main_window()
    except Exception as exc:
        print("Connect main window failed:", exc)
        return False

    if not activate_win(main):
        return False

    def find_project_management_menu_item():
        try:
            ctrl = main.child_window(title=PROJECT_MANAGEMENT_TEXT, control_type="MenuItem").wrapper_object()
            if ctrl is not None:
                return ctrl
        except Exception:
            pass

        try:
            menu_items = main.descendants(control_type="MenuItem")
        except Exception as exc:
            print("Scan project management menu items failed:", exc)
            return None

        for menu_item in menu_items:
            try:
                if menu_item.window_text() == PROJECT_MANAGEMENT_TEXT:
                    return menu_item
            except Exception:
                continue
        return None

    def find_app_menu_bar():
        try:
            ctrl = main.child_window(title=APP_MENU_BAR_TEXT, control_type="MenuBar").wrapper_object()
            if ctrl is not None:
                return ctrl
        except Exception:
            pass

        try:
            menu_bars = main.descendants(control_type="MenuBar")
        except Exception as exc:
            print("Scan menu bars failed:", exc)
            return None

        for menu_bar in menu_bars:
            try:
                if menu_bar.window_text() == APP_MENU_BAR_TEXT:
                    return menu_bar
            except Exception:
                continue
        return None

    project_management_item = retry_window_operation(
        find_project_management_menu_item,
        action_name="Find project management menu item",
    )

    if project_management_item is not None:
        try:
            project_management_item.click_input()
        except Exception as exc:
            print("Click project management failed:", exc)
            return False
    else:
        app_menu_bar = retry_window_operation(find_app_menu_bar, action_name="Find application menu bar")
        if app_menu_bar is None:
            return False

        try:
            menu_items = app_menu_bar.children(control_type="MenuItem")
        except Exception as exc:
            print("Read menu items failed:", exc)
            return False

        if len(menu_items) <= PROJECT_MANAGEMENT_MENU_INDEX:
            print("Menu item count is abnormal")
            return False

        try:
            menu_items[PROJECT_MANAGEMENT_MENU_INDEX].click_input()
        except Exception as exc:
            print("Click project management failed:", exc)
            return False

    time.sleep(0.5)
    try:
        send_keys(f"{{DOWN {CHANGE_PROJECT_DOWN_COUNT}}}")
        time.sleep(0.2)
        send_keys("{ENTER}")
        time.sleep(1.0)
    except Exception as exc:
        print("Send change project shortcut failed:", exc)
        return False

    try:
        connect_project_list_win()
        return True
    except Exception:
        return False


def find_top_visible_period_index(visible: list) -> int:
    for index, ctrl in enumerate(visible):
        if is_period_text(safe_text(ctrl)):
            return index
    return -1


def find_root_ctrl_by_period_context(win, root_name: str, debug: bool = True):
    visible = visible_items(win, debug=debug)
    if not visible:
        return None

    if len(visible) == 1 and safe_text(visible[0]) == root_name:
        return visible[0]

    period_index = find_top_visible_period_index(visible)
    if period_index <= 0:
        return None

    candidate = visible[period_index - 1]
    if safe_text(candidate) != root_name:
        return None
    return candidate


def is_collapsed_to_root_only(win, root_name: str, debug: bool = True) -> bool:
    visible = visible_items(win, debug=debug)
    return len(visible) == 1 and safe_text(visible[0]) == root_name


def send_pageup(win) -> bool:
    if not activate_win(win):
        return False
    try:
        send_keys("{PGUP}")
        time.sleep(PAGEUP_SLEEP)
        return True
    except Exception:
        return False


def pageup_until_root_context_visible(win, root_name: str, max_times: int = MAX_PAGEUP_TIMES) -> bool:
    previous_signature = page_signature(win)

    for _ in range(max_times):
        ctrl = find_root_ctrl_by_period_context(win, root_name, debug=True)
        if ctrl is not None:
            return True

        if not send_pageup(win):
            return False

        current_signature = page_signature(win)
        previous_signature = current_signature

    return False


def collapse_tree_to_root_only(win, root_name: str, max_try: int = MAX_COLLAPSE_TRIES) -> bool:
    for _ in range(max_try):
        if is_collapsed_to_root_only(win, root_name=root_name, debug=True):
            return True

        root_ctrl = find_root_ctrl_by_period_context(win, root_name, debug=True)
        if root_ctrl is None:
            ok = pageup_until_root_context_visible(win, root_name)
            if not ok:
                return False
            root_ctrl = find_root_ctrl_by_period_context(win, root_name, debug=True)
            if root_ctrl is None:
                return False

        if not double_click(root_ctrl, f"Collapse root project {root_name}"):
            return False

        time.sleep(0.8)
        if is_collapsed_to_root_only(win, root_name=root_name, debug=True):
            return True

    return False


def focus_tree_by_single_click(win) -> bool:
    if not activate_win(win):
        return False
    visible = visible_items(win, debug=True)
    if not visible:
        return False
    try:
        visible[0].click_input()
        time.sleep(0.5)
        return True
    except Exception:
        return False


def reset_project_tree_for_next_round(
    root_name: str | None = None,
    csv_path: Path | None = None,
) -> bool:
    if root_name is None:
        try:
            root_name = load_root_project_name_from_csv(csv_path)
        except Exception as exc:
            print("Read root project name failed:", exc)
            return False

    if not open_change_project_dialog():
        return False

    try:
        win = connect_project_list_win()
    except Exception as exc:
        print("Connect project list window failed:", exc)
        return False

    if not activate_win(win):
        return False
    if not focus_tree_by_single_click(win):
        return False

    visible = visible_items(win, debug=True)
    if not visible:
        return False
    if is_collapsed_to_root_only(win, root_name=root_name, debug=True):
        return True

    return collapse_tree_to_root_only(win, root_name=root_name)


if __name__ == "__main__":
    ok = reset_project_tree_for_next_round()
    print("Reset result:", ok)
