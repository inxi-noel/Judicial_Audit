from __future__ import annotations

import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import win32api
import win32con
import win32gui
import win32process
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

from modules._shared.config import (
    COMPUTE_RECT,
    COMPUTE_RECT_TOLERANCE,
    CONFIRM_BUTTON_RECT,
    CONFIRM_BUTTON_TEXTS,
    CONFIRM_BUTTON_TYPES,
    DATA_ONLY_RADIO_RECT,
    DATA_ONLY_RADIO_TEXTS,
    DATA_ONLY_RADIO_TYPES,
    DIALOG_CLASS,
    EXPORT_BUTTON_RECT,
    EXPORT_BUTTON_TEXTS,
    EXPORT_BUTTON_TYPES,
    EXPORT_DIR,
    EXPAND_ALL_RECT,
    EXPAND_ALL_TEXTS,
    EXPAND_ALL_TYPES,
    LEDGER_WINDOW_REQUIRED_TEXT,
    MAIN_WINDOW_REQUIRED_TEXT,
    MAIN_WINDOW_TITLE_RE,
    MONTH_AUTO_ID,
    MONTH_COMBO_RECT,
    MONTH_ITEM_RECTS,
    MULTIYEAR_LEDGER_ENTRY_NAME,
    MULTIYEAR_LEDGER_ENTRY_TYPE,
    MULTIYEAR_LEDGER_REGION_NAME,
    POST_CONFIRM_WAIT,
    PROMPT_DIALOG_TITLE,
    QUERY_BUTTON_RECT,
    QUERY_BUTTON_TEXT,
    REPORT_NAME,
    SAVE_DIALOG_TITLE_KEYWORD,
    SUCCESS_TEXT_PART_1,
    SUCCESS_TEXT_PART_2,
    TARGET_LEDGER_YEAR,
    YEAR_AUTO_ID,
    YEAR_RECT,
    YES_BUTTON_RECT,
)
from modules._shared.logger import parse_export_filename
from modules._shared.window_retry import retry_window_operation


FOCUS_WAIT = 0.5
OPEN_WAIT = 0.45
AFTER_SELECT_WAIT = 0.70
AFTER_QUERY_CLICK_WAIT = 0.80
AFTER_OPEN_CLICK_WAIT = 1.20
POST_QUERY_WAIT_1 = 3.0
POST_QUERY_WAIT_2 = 3.0
POST_QUERY_SETTLE_WAIT = 5.0
PERIOD_CONTROL_READY_TIMEOUT = 30.0
PERIOD_CONTROL_READY_POLL_INTERVAL = 2.0

OPEN_READY_POLL_INTERVAL = 1.0
OPEN_READY_TIMEOUT = 60.0
QUERY_POLL_INTERVAL = 2.0
QUERY_TIMEOUT = 180.0
QUERY_STABLE_ROUNDS = 2

DIALOG_TIMEOUT = 8.0
CONFIG_TIMEOUT = 15.0
SAVE_DIALOG_SCAN_TIMEOUT = 10.0
SUCCESS_PROMPT_TIMEOUT = 30.0
SCAN_INTERVAL = 0.2
RECT_TOLERANCE = 35
SAVE_AFTER_ALT_S_WAIT = 0.5
SUCCESS_AFTER_ALT_N_WAIT = 0.5
EXPORT_FILE_TIMEOUT = 60.0
EXPORT_FILE_POLL_INTERVAL = 0.5
EXPORT_INITIAL_SAVE_DIALOG_WAIT = 40.0
EXPORT_RETRY_INTERVAL = 10.0
NON_SENSITIVE_RETRY_TIMES = 20
NON_SENSITIVE_RETRY_INTERVAL_SEC = 3.0
EXPORT_ENTRY_RETRY_TIMES = 5
EXPORT_ENTRY_RETRY_INTERVAL_SEC = 3.0
EXPAND_RETRY_WAIT = 1.0
POST_ALT_Y_WAIT = 1.0
POST_DATA_ONLY_CLICK_WAIT = 0.4


def log(msg: str) -> None:
    print(msg)


def get_desktop():
    return Desktop(backend="uia")


def safe_window_text(wrapper) -> str:
    try:
        return (wrapper.window_text() or "").strip()
    except Exception:
        return ""


def safe_control_type(wrapper) -> str:
    try:
        return str(getattr(wrapper.element_info, "control_type", "") or "")
    except Exception:
        return ""


def safe_friendly_class(wrapper) -> str:
    try:
        return str(wrapper.friendly_class_name() or "")
    except Exception:
        return ""


def safe_automation_id(wrapper) -> str:
    try:
        return str(getattr(wrapper.element_info, "automation_id", "") or "")
    except Exception:
        return ""


def get_rect(wrapper):
    try:
        rect = wrapper.rectangle()
        return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        return None


def rect_valid(rect: Optional[tuple[int, int, int, int]]) -> bool:
    return rect is not None and rect[0] < rect[2] and rect[1] < rect[3]


def rect_center(rect: tuple[int, int, int, int]) -> tuple[int, int]:
    return ((rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)


def rect_match(actual: Optional[tuple[int, int, int, int]], target: Optional[tuple[int, int, int, int]], tolerance: int = RECT_TOLERANCE) -> bool:
    if target is None:
        return True
    if actual is None:
        return False
    return all(abs(a - b) <= tolerance for a, b in zip(actual, target))


def rect_equal(actual: Optional[tuple[int, int, int, int]], target: tuple[int, int, int, int]) -> bool:
    return actual == target


def rect_close_enough(actual: Optional[tuple[int, int, int, int]], target: tuple[int, int, int, int], tol: int) -> bool:
    if actual is None:
        return False
    return all(abs(a - b) <= tol for a, b in zip(actual, target))


def rect_distance(actual: Optional[tuple[int, int, int, int]], target: Optional[tuple[int, int, int, int]]) -> int:
    if actual is None or target is None:
        return 10**18
    return sum(abs(a - b) for a, b in zip(actual, target))


def rect_intersects(a, b) -> bool:
    return not (
        a[2] <= b[0] or
        a[0] >= b[2] or
        a[3] <= b[1] or
        a[1] >= b[3]
    )


def click_rect(rect: tuple[int, int, int, int]) -> None:
    x, y = rect_center(rect)
    win32api.SetCursorPos((x, y))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.03)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def click_control(wrapper, fallback_rect: Optional[tuple[int, int, int, int]] = None, *, prefer_rect: bool = False) -> None:
    rect = get_rect(wrapper)

    if prefer_rect and rect_valid(rect):
        click_rect(rect)
        return
    if prefer_rect and fallback_rect is not None:
        click_rect(fallback_rect)
        return

    try:
        wrapper.click_input()
        return
    except Exception:
        pass

    if rect_valid(rect):
        click_rect(rect)
        return
    if fallback_rect is not None:
        click_rect(fallback_rect)
        return

    raise RuntimeError("Control cannot be clicked")


def set_focus_best(wrapper) -> None:
    try:
        wrapper.set_focus()
        time.sleep(FOCUS_WAIT)
        return
    except Exception:
        pass

    try:
        hwnd = wrapper.handle
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(FOCUS_WAIT)
        return
    except Exception:
        pass

    raise RuntimeError("Cannot activate target window")


def enum_visible_top_windows() -> list[dict]:
    result: list[dict] = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            result.append(
                {
                    "hwnd": hwnd,
                    "title": win32gui.GetWindowText(hwnd),
                    "class": win32gui.GetClassName(hwnd),
                    "rect": win32gui.GetWindowRect(hwnd),
                    "pid": win32process.GetWindowThreadProcessId(hwnd)[1],
                }
            )
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return result


def get_foreground_info() -> dict | None:
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return None
    try:
        return {
            "hwnd": hwnd,
            "title": win32gui.GetWindowText(hwnd),
            "class": win32gui.GetClassName(hwnd),
            "rect": win32gui.GetWindowRect(hwnd),
            "pid": win32process.GetWindowThreadProcessId(hwnd)[1],
        }
    except Exception:
        return None


def _find_top_window_by_text(required_text: str, timeout: float) -> object:
    desktop = get_desktop()
    deadline = time.time() + timeout

    while time.time() < deadline:
        candidates = []
        for item in enum_visible_top_windows():
            title = item["title"] or ""
            if required_text not in title:
                continue
            if title == "Ghost":
                continue
            try:
                wrapper = desktop.window(handle=item["hwnd"])
                candidates.append((wrapper, item))
            except Exception:
                pass

        if candidates:
            candidates.sort(
                key=lambda pair: (pair[1]["rect"][2] - pair[1]["rect"][0]) * (pair[1]["rect"][3] - pair[1]["rect"][1]),
                reverse=True,
            )
            return candidates[0][0]

        time.sleep(0.2)

    raise RuntimeError(f"Window not found | text={required_text}")


def get_main_screen_window(timeout: float = 10.0):
    return _find_top_window_by_text(MAIN_WINDOW_REQUIRED_TEXT, timeout)


def get_ledger_window(timeout: float = 10.0):
    return _find_top_window_by_text(LEDGER_WINDOW_REQUIRED_TEXT, timeout)


def try_get_ledger_window(timeout: float = 1.0):
    try:
        return get_ledger_window(timeout)
    except Exception:
        return None


def get_wrapper_pid(wrapper) -> int:
    try:
        return int(wrapper.process_id())
    except Exception:
        pass

    try:
        hwnd = int(wrapper.handle)
        return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception as exc:
        raise RuntimeError(f"Get window pid failed: {exc}") from exc


def prepare_export_ledger_window():
    win = get_ledger_window(timeout=15.0)
    set_focus_best(win)
    return win, get_wrapper_pid(win)


def safe_get_legacy_value(ctrl) -> str:
    try:
        props = ctrl.legacy_properties()
        return str(props.get("Value", "") or "").strip()
    except Exception:
        return ""


def safe_get_value_pattern(ctrl) -> str:
    try:
        return str(ctrl.iface_value.CurrentValue or "").strip()
    except Exception:
        return ""


def read_ctrl_value(ctrl) -> str:
    value = safe_get_value_pattern(ctrl)
    if value:
        return value
    value = safe_get_legacy_value(ctrl)
    if value:
        return value
    return safe_window_text(ctrl)


def locate_multiyear_entry(win):
    try:
        ctrl = win.child_window(
            title=MULTIYEAR_LEDGER_ENTRY_NAME,
            control_type=MULTIYEAR_LEDGER_ENTRY_TYPE,
        ).wrapper_object()
        rect = get_rect(ctrl)
        if rect_valid(rect):
            return ctrl, rect
    except Exception:
        pass

    region = win.child_window(title=MULTIYEAR_LEDGER_REGION_NAME, control_type="Image").wrapper_object()
    region_rect = get_rect(region)
    if not rect_valid(region_rect):
        raise RuntimeError("Ledger region anchor not found")

    candidates = []
    for ctrl in win.descendants():
        text = safe_window_text(ctrl)
        control_type = safe_control_type(ctrl)
        rect = get_rect(ctrl)

        if text != MULTIYEAR_LEDGER_ENTRY_NAME:
            continue
        if control_type != MULTIYEAR_LEDGER_ENTRY_TYPE:
            continue
        if not rect_valid(rect):
            continue
        if not rect_intersects(region_rect, rect):
            continue

        candidates.append((ctrl, rect))

    if not candidates:
        raise RuntimeError(
            f"Multiyear ledger entry not found: {MULTIYEAR_LEDGER_ENTRY_NAME!r} / {MULTIYEAR_LEDGER_ENTRY_TYPE!r}"
        )

    candidates.sort(key=lambda item: (item[1][1], item[1][0]))
    return candidates[0]


def find_compute_anchor(win):
    for ctrl in win.descendants():
        if safe_control_type(ctrl) != "Text":
            continue
        rect = get_rect(ctrl)
        if not rect_close_enough(rect, COMPUTE_RECT, COMPUTE_RECT_TOLERANCE):
            continue
        return ctrl
    return None


def read_compute_anchor(win) -> dict:
    ctrl = find_compute_anchor(win)
    if ctrl is None:
        return {
            "found": False,
            "type": "",
            "text": "",
            "value": "",
            "rect": "None",
            "non_empty": False,
        }

    rect = get_rect(ctrl)
    value = read_ctrl_value(ctrl)
    text = safe_window_text(ctrl)

    return {
        "found": True,
        "type": safe_control_type(ctrl),
        "text": text,
        "value": value,
        "rect": str(rect),
        "non_empty": bool(value or text),
    }


def wait_compute_anchor_ready(timeout: float = OPEN_READY_TIMEOUT) -> dict:
    deadline = time.time() + timeout
    round_index = 0

    while time.time() < deadline:
        round_index += 1
        try:
            win = get_ledger_window(timeout=8.0)
            state = read_compute_anchor(win)
        except Exception as exc:
            print(f"[{round_index:03d}] read compute_1 failed: {exc}")
            time.sleep(OPEN_READY_POLL_INTERVAL)
            continue

        print(
            f"[{round_index:03d}] compute_1: found={state['found']}, "
            f"text={state['text']!r}, value={state['value']!r}, non_empty={state['non_empty']}"
        )
        if state["found"] and state["non_empty"]:
            return state
        time.sleep(OPEN_READY_POLL_INTERVAL)

    raise RuntimeError("Timed out waiting for compute_1 to become non-empty")


def enter_multiyear_general_ledger() -> bool:
    clear_runtime_cache()
    win = get_main_screen_window(timeout=10.0)
    set_focus_best(win)

    ctrl, rect = locate_multiyear_entry(win)
    click_control(ctrl, fallback_rect=rect)
    time.sleep(AFTER_OPEN_CLICK_WAIT)

    state = wait_compute_anchor_ready(timeout=OPEN_READY_TIMEOUT)
    print(f"compute_1 ready -> {state['value']!r}")
    return True



@dataclass
class RuntimeCache:
    """Cache only relatively stable controls for month switching."""
    year_ctrl: Optional[object] = None
    month_ctrl: Optional[object] = None
    compute_ctrl: Optional[object] = None


RUNTIME = RuntimeCache()


def clear_runtime_cache() -> None:
    RUNTIME.year_ctrl = None
    RUNTIME.month_ctrl = None
    RUNTIME.compute_ctrl = None


def clear_compute_cache() -> None:
    RUNTIME.compute_ctrl = None


def get_query_window(timeout: float = 10.0):
    return get_ledger_window(timeout=timeout)


def focus_query_window(timeout: float = 10.0):
    win = get_query_window(timeout=timeout)
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(FOCUS_WAIT)
    return win


def read_combo_value(ctrl) -> str:
    value = safe_get_value_pattern(ctrl)
    if value:
        return value

    value = safe_get_legacy_value(ctrl)
    if value:
        return value

    value = safe_window_text(ctrl)
    if value:
        return value

    return ""


def combo_ctrl_alive(ctrl, rect_tuple: tuple[int, int, int, int]) -> bool:
    if ctrl is None:
        return False
    try:
        rect = get_rect(ctrl)
        return (
            safe_control_type(ctrl) == "ComboBox"
            and rect_valid(rect)
            and rect_close_enough(rect, rect_tuple, COMPUTE_RECT_TOLERANCE)
        )
    except Exception:
        return False


def bind_period_controls_direct(win):
    year_ctrl = win.child_window(auto_id=YEAR_AUTO_ID, control_type="ComboBox").wrapper_object()
    month_ctrl = win.child_window(auto_id=MONTH_AUTO_ID, control_type="ComboBox").wrapper_object()

    if not combo_ctrl_alive(year_ctrl, YEAR_RECT):
        raise RuntimeError(f"Year combo direct bind failed | rect={get_rect(year_ctrl)}")

    if not combo_ctrl_alive(month_ctrl, MONTH_COMBO_RECT):
        raise RuntimeError(f"Month combo direct bind failed | rect={get_rect(month_ctrl)}")

    return year_ctrl, month_ctrl


def bind_period_controls_fallback_scan(win):
    year_ctrl = None
    month_ctrl = None

    for ctrl in win.descendants():
        if safe_control_type(ctrl) != "ComboBox":
            continue

        auto_id = safe_automation_id(ctrl)
        rect = get_rect(ctrl)
        if not rect_valid(rect):
            continue

        if auto_id == YEAR_AUTO_ID and rect_equal(rect, YEAR_RECT):
            year_ctrl = ctrl
        elif auto_id == MONTH_AUTO_ID and rect_equal(rect, MONTH_COMBO_RECT):
            month_ctrl = ctrl

        if year_ctrl is not None and month_ctrl is not None:
            break

    if year_ctrl is None:
        raise RuntimeError(f"Year combo not found | auto_id={YEAR_AUTO_ID} | rect={YEAR_RECT}")
    if month_ctrl is None:
        raise RuntimeError(f"Month combo not found | auto_id={MONTH_AUTO_ID} | rect={MONTH_COMBO_RECT}")

    return year_ctrl, month_ctrl


def wait_period_controls_ready(timeout: float = PERIOD_CONTROL_READY_TIMEOUT):
    deadline = time.time() + timeout
    last_error = None

    while time.time() < deadline:
        try:
            win = focus_query_window(timeout=8.0)
            try:
                year_ctrl, month_ctrl = bind_period_controls_direct(win)
            except Exception:
                year_ctrl, month_ctrl = bind_period_controls_fallback_scan(win)
            RUNTIME.year_ctrl = year_ctrl
            RUNTIME.month_ctrl = month_ctrl
            return win, year_ctrl, month_ctrl
        except Exception as exc:
            last_error = exc
            time.sleep(PERIOD_CONTROL_READY_POLL_INTERVAL)

    if last_error is None:
        raise RuntimeError("Year/month controls not ready within timeout")
    raise RuntimeError(str(last_error))


def ensure_period_controls(win, force_rebind: bool = False):
    if not force_rebind:
        if combo_ctrl_alive(RUNTIME.year_ctrl, YEAR_RECT) and combo_ctrl_alive(RUNTIME.month_ctrl, MONTH_COMBO_RECT):
            return RUNTIME.year_ctrl, RUNTIME.month_ctrl

    try:
        year_ctrl, month_ctrl = bind_period_controls_direct(win)
    except Exception:
        try:
            year_ctrl, month_ctrl = bind_period_controls_fallback_scan(win)
        except Exception:
            rebound_win, year_ctrl, month_ctrl = wait_period_controls_ready(timeout=PERIOD_CONTROL_READY_TIMEOUT)
            _ = rebound_win

    RUNTIME.year_ctrl = year_ctrl
    RUNTIME.month_ctrl = month_ctrl
    return year_ctrl, month_ctrl


def read_period_anchor(year_ctrl, month_ctrl) -> tuple[str, str]:
    return read_combo_value(year_ctrl), read_combo_value(month_ctrl)


def compute_ctrl_alive(ctrl) -> bool:
    if ctrl is None:
        return False
    try:
        rect = get_rect(ctrl)
        return (
            safe_control_type(ctrl) == "Text"
            and rect_valid(rect)
            and rect_close_enough(rect, COMPUTE_RECT, COMPUTE_RECT_TOLERANCE)
        )
    except Exception:
        return False


def bind_compute_ctrl_once(win):
    candidates = []

    for ctrl in win.descendants():
        if safe_control_type(ctrl) != "Text":
            continue

        rect = get_rect(ctrl)
        if not rect_valid(rect):
            continue

        if rect_close_enough(rect, COMPUTE_RECT, COMPUTE_RECT_TOLERANCE):
            legacy_value = safe_get_legacy_value(ctrl)
            candidates.append((ctrl, legacy_value))

    if not candidates:
        raise RuntimeError(f"Compute control not found | rect={COMPUTE_RECT}")

    for ctrl, legacy_value in candidates:
        if legacy_value:
            return ctrl

    return candidates[0][0]


def ensure_compute_ctrl(win, force_rebind: bool = False):
    if not force_rebind and compute_ctrl_alive(RUNTIME.compute_ctrl):
        return RUNTIME.compute_ctrl

    ctrl = bind_compute_ctrl_once(win)
    RUNTIME.compute_ctrl = ctrl
    return ctrl


def read_compute_value_only(win, force_rebind: bool = False) -> str:
    ctrl = ensure_compute_ctrl(win, force_rebind=force_rebind)
    value = safe_get_legacy_value(ctrl)
    if value:
        return value
    raise RuntimeError("Compute control located but Legacy Value is empty")


def select_target_month(win, year_ctrl, month_ctrl, target_year: str, target_month: str):
    current_year, current_month = read_period_anchor(year_ctrl, month_ctrl)
    print(f"Current period anchor -> year={current_year!r}, month={current_month!r}")

    if current_year != target_year:
        raise RuntimeError(
            f"Unexpected year in ledger window | current={current_year!r} | target={target_year!r}"
        )

    if current_month == target_month:
        print(f"Month already selected -> {target_month}")
        return year_ctrl, month_ctrl

    click_rect(MONTH_COMBO_RECT)
    print(f"Month combo click -> rect={MONTH_COMBO_RECT}")
    time.sleep(OPEN_WAIT)

    click_rect(MONTH_ITEM_RECTS[target_month])
    print(f"Month item click -> rect={MONTH_ITEM_RECTS[target_month]}")
    time.sleep(AFTER_SELECT_WAIT)

    if not combo_ctrl_alive(year_ctrl, YEAR_RECT) or not combo_ctrl_alive(month_ctrl, MONTH_COMBO_RECT):
        year_ctrl, month_ctrl = ensure_period_controls(win, force_rebind=True)

    new_year, new_month = read_period_anchor(year_ctrl, month_ctrl)
    print(f"After month switch -> year={new_year!r}, month={new_month!r}")

    if new_year == target_year and new_month == target_month:
        return year_ctrl, month_ctrl

    raise RuntimeError(
        f"Period anchor verification failed | expected=({target_year},{target_month}) | actual=({new_year},{new_month})"
    )


def normalize_period_text(value: str) -> str:
    if not value:
        return ""
    value = re.sub(r"[\s\u3000]+", "", value)
    value = re.sub(r"[\uFF0D\u2014\u2013-]+", "\u2014", value)
    return value


def expected_compute_value(year: str, month: str) -> str:
    raw = f"{year}\u5e741\u6708\uff0d\uff0d{year}\u5e74{int(month)}\u6708"
    return normalize_period_text(raw)


def compute_matches_target(compute_value: str, year: str, month: str) -> bool:
    return normalize_period_text(compute_value) == expected_compute_value(year, month)


def click_query_once():
    clear_compute_cache()
    click_rect(QUERY_BUTTON_RECT)
    print(f"Clicked query button by rect -> rect={QUERY_BUTTON_RECT}")
    time.sleep(AFTER_QUERY_CLICK_WAIT)


def verify_query_result_for_month(year: str, month: str) -> dict:
    expected_norm = expected_compute_value(year, month)
    print(f"Expected compute_1 -> {expected_norm!r}")

    time.sleep(POST_QUERY_WAIT_1)
    win = get_query_window(timeout=3.0)

    value_1 = read_compute_value_only(win, force_rebind=True)
    norm_1 = normalize_period_text(value_1)
    matched_1 = norm_1 == expected_norm
    print(f"First compute read -> value={value_1!r} | matched={matched_1}")

    if matched_1:
        return {
            "found": True,
            "type": "Text",
            "value": value_1,
            "rect": str(get_rect(RUNTIME.compute_ctrl)),
        }

    time.sleep(POST_QUERY_WAIT_2)
    clear_compute_cache()
    win = get_query_window(timeout=3.0)

    value_2 = read_compute_value_only(win, force_rebind=True)
    norm_2 = normalize_period_text(value_2)
    matched_2 = norm_2 == expected_norm
    print(f"Second compute read -> value={value_2!r} | matched={matched_2}")

    if matched_2:
        return {
            "found": True,
            "type": "Text",
            "value": value_2,
            "rect": str(get_rect(RUNTIME.compute_ctrl)),
        }

    raise RuntimeError(
        f"Query result period mismatch | expected={expected_norm!r} | actual1={norm_1!r} | actual2={norm_2!r}"
    )


def select_month_and_query(month: int) -> bool:
    target_month = str(int(month))
    if target_month not in MONTH_ITEM_RECTS:
        raise RuntimeError(f"Target month rect is not configured: {target_month}")

    try:
        win = focus_query_window(timeout=8.0)
        year_ctrl, month_ctrl = ensure_period_controls(win, force_rebind=False)

        year_ctrl, month_ctrl = select_target_month(
            win=win,
            year_ctrl=year_ctrl,
            month_ctrl=month_ctrl,
            target_year=TARGET_LEDGER_YEAR,
            target_month=target_month,
        )

        win = focus_query_window(timeout=8.0)
        click_query_once()

        print(f"Post-query settle wait -> {POST_QUERY_SETTLE_WAIT:.1f}s")
        time.sleep(POST_QUERY_SETTLE_WAIT)

        print(f"Month query done | month={month}")
        return True
    except Exception:
        clear_runtime_cache()
        raise


def exact_name_match(actual: str, candidates: Iterable[str]) -> bool:
    actual_norm = (actual or "").strip().lower()
    return any(actual_norm == (candidate or "").strip().lower() for candidate in candidates)


def contains_name_match(actual: str, candidates: Iterable[str]) -> bool:
    actual_norm = (actual or "").strip().lower()
    return any((candidate or "").strip().lower() in actual_norm for candidate in candidates if candidate)


def type_match(wrapper, allowed_types: Iterable[str]) -> bool:
    allowed = {str(item).strip().lower() for item in allowed_types}
    actuals = set()
    control_type = safe_control_type(wrapper)
    if control_type:
        actuals.add(control_type.strip().lower())
    friendly_class = safe_friendly_class(wrapper)
    if friendly_class:
        actuals.add(friendly_class.strip().lower())
    return bool(actuals & allowed)


def find_control_by_rule(
    descendants,
    target_names: Iterable[str],
    target_types: Iterable[str],
    target_rect: Optional[tuple[int, int, int, int]] = None,
    *,
    allow_contains: bool = False,
):
    hits = []

    for wrapper in descendants:
        name = safe_window_text(wrapper)
        matched = contains_name_match(name, target_names) if allow_contains else exact_name_match(name, target_names)
        if not matched:
            continue
        if not type_match(wrapper, target_types):
            continue
        rect = get_rect(wrapper)
        if not rect_match(rect, target_rect):
            continue
        hits.append((wrapper, rect, name))

    if not hits:
        raise RuntimeError(
            f"Control not found | names={tuple(target_names)} | types={tuple(target_types)} | rect={target_rect}"
        )

    if target_rect is not None:
        hits.sort(key=lambda item: rect_distance(item[1], target_rect))
    return hits[0]


def try_find_control_by_rule(
    descendants,
    target_names: Iterable[str],
    target_types: Iterable[str],
    target_rect: Optional[tuple[int, int, int, int]] = None,
    *,
    allow_contains: bool = False,
):
    try:
        return find_control_by_rule(
            descendants,
            target_names,
            target_types,
            target_rect,
            allow_contains=allow_contains,
        )
    except Exception:
        return None


def collect_same_pid_top_wrappers(main_pid: int):
    desktop = get_desktop()
    wrappers = []

    for item in enum_visible_top_windows():
        if item["pid"] != main_pid:
            continue
        try:
            wrappers.append(desktop.window(handle=item["hwnd"]))
        except Exception:
            pass

    return wrappers


def click_expand_all(win) -> None:
    set_focus_best(win)
    click_rect(EXPAND_ALL_RECT)
    log(f"Clicked expand all by rect | rect={EXPAND_ALL_RECT}")


def wait_for_prompt_dialog(main_pid: int, timeout: float = DIALOG_TIMEOUT):
    desktop = get_desktop()
    deadline = time.time() + timeout

    while time.time() < deadline:
        fg = get_foreground_info()
        if fg and fg["pid"] == main_pid and fg["title"] == PROMPT_DIALOG_TITLE and fg["class"] == DIALOG_CLASS:
            try:
                return desktop.window(handle=fg["hwnd"])
            except Exception:
                pass
        time.sleep(SCAN_INTERVAL)

    raise RuntimeError("Timed out waiting for prompt dialog")


def try_wait_for_prompt_dialog(main_pid: int, timeout: float):
    try:
        return wait_for_prompt_dialog(main_pid=main_pid, timeout=timeout)
    except Exception:
        return None


def choose_yes_by_alt_y(dialog) -> None:
    set_focus_best(dialog)
    send_keys("%y")
    log("Sent Alt+Y")


def wait_for_expand_config_dialog(main_pid: int, timeout: float = CONFIG_TIMEOUT):
    desktop = get_desktop()
    deadline = time.time() + timeout

    while time.time() < deadline:
        fg = get_foreground_info()
        if fg and fg["pid"] == main_pid and fg["class"] == DIALOG_CLASS and fg["title"] != PROMPT_DIALOG_TITLE:
            try:
                return desktop.window(handle=fg["hwnd"])
            except Exception:
                pass
        time.sleep(SCAN_INTERVAL)

    return None


def wait_for_expand_config_controls(main_pid: int, timeout: float = CONFIG_TIMEOUT):
    deadline = time.time() + timeout
    last_scan = []

    while time.time() < deadline:
        wrappers = collect_same_pid_top_wrappers(main_pid)

        for host in wrappers:
            try:
                descendants = host.descendants()
            except Exception:
                continue

            radio_found = try_find_control_by_rule(
                descendants=descendants,
                target_names=DATA_ONLY_RADIO_TEXTS,
                target_types=DATA_ONLY_RADIO_TYPES,
                target_rect=DATA_ONLY_RADIO_RECT,
                allow_contains=True,
            )
            confirm_found = try_find_control_by_rule(
                descendants=descendants,
                target_names=CONFIRM_BUTTON_TEXTS,
                target_types=CONFIRM_BUTTON_TYPES,
                target_rect=CONFIRM_BUTTON_RECT,
                allow_contains=False,
            )

            last_scan.append((safe_window_text(host), bool(radio_found), bool(confirm_found)))
            if radio_found and confirm_found:
                return host, radio_found, confirm_found

        time.sleep(0.2)

    raise RuntimeError(f"Timed out waiting for expand config controls | last_scan={last_scan[-10:]}")


def select_data_only_and_confirm(main_pid: int) -> None:
    if DATA_ONLY_RADIO_RECT is not None and CONFIRM_BUTTON_RECT is not None:
        dialog = wait_for_expand_config_dialog(main_pid=main_pid, timeout=CONFIG_TIMEOUT)
        if dialog is not None:
            try:
                set_focus_best(dialog)
            except Exception:
                pass

        click_rect(DATA_ONLY_RADIO_RECT)
        log(f"Clicked data-only radio by rect | rect={DATA_ONLY_RADIO_RECT}")
        time.sleep(POST_DATA_ONLY_CLICK_WAIT)

        click_rect(CONFIRM_BUTTON_RECT)
        log(f"Clicked confirm by rect | rect={CONFIRM_BUTTON_RECT}")
        return

    host, radio_found, confirm_found = wait_for_expand_config_controls(main_pid=main_pid, timeout=CONFIG_TIMEOUT)
    radio_ctrl, radio_rect, radio_name = radio_found
    confirm_ctrl, confirm_rect, confirm_name = confirm_found

    set_focus_best(host)
    click_control(radio_ctrl, fallback_rect=DATA_ONLY_RADIO_RECT, prefer_rect=False)
    log(f"Clicked data-only radio | name={radio_name} | rect={radio_rect}")
    time.sleep(0.3)
    click_control(confirm_ctrl, fallback_rect=CONFIRM_BUTTON_RECT, prefer_rect=False)
    log(f"Clicked confirm | name={confirm_name} | rect={confirm_rect}")


def expand_all_with_data_only() -> bool:
    win = get_ledger_window(timeout=15.0)
    set_focus_best(win)

    main_pid = win.process_id()
    click_expand_all(win)
    time.sleep(0.3)

    dialog = try_wait_for_prompt_dialog(main_pid=main_pid, timeout=3.0)
    if dialog is None:
        time.sleep(EXPAND_RETRY_WAIT)
        win = get_ledger_window(timeout=8.0)
        set_focus_best(win)
        click_expand_all(win)
        time.sleep(0.3)
        dialog = wait_for_prompt_dialog(main_pid=main_pid, timeout=DIALOG_TIMEOUT)

    choose_yes_by_alt_y(dialog)
    time.sleep(POST_ALT_Y_WAIT)

    select_data_only_and_confirm(main_pid=main_pid)
    log(f"Fixed wait after confirm -> {POST_CONFIRM_WAIT:.0f}s")
    time.sleep(POST_CONFIRM_WAIT)
    return True


def click_export_button(win) -> None:
    click_rect(EXPORT_BUTTON_RECT)
    log(f"Clicked export button by rect | rect={EXPORT_BUTTON_RECT}")


def find_save_dialog(main_pid: Optional[int] = None):
    hits = []

    for item in enum_visible_top_windows():
        title = item["title"] or ""
        if SAVE_DIALOG_TITLE_KEYWORD in title and item["class"] == DIALOG_CLASS:
            hits.append(item)

    if not hits:
        return None

    if main_pid is not None:
        same_pid_hits = [item for item in hits if item["pid"] == main_pid]
        if same_pid_hits:
            return same_pid_hits[0]

    return hits[0]


def wait_for_save_dialog(timeout: float, main_pid: Optional[int] = None):
    deadline = time.time() + timeout

    while time.time() < deadline:
        hit = find_save_dialog(main_pid=main_pid)
        if hit:
            return hit
        time.sleep(SCAN_INTERVAL)
    return None


def save_file_by_alt_s(save_hwnd: int) -> None:
    dialog = get_desktop().window(handle=save_hwnd)
    set_focus_best(dialog)
    send_keys("%s")
    log("Sent Alt+S")


def window_contains_success_text(hwnd: int) -> bool:
    try:
        wrapper = get_desktop().window(handle=hwnd)
        descendants = wrapper.descendants()
    except Exception:
        return False

    texts = []
    for ctrl in descendants:
        text = safe_window_text(ctrl)
        if text:
            texts.append(text)

    merged = "\n".join(texts)
    return SUCCESS_TEXT_PART_1 in merged and SUCCESS_TEXT_PART_2 in merged


def find_success_prompt_dialog(main_pid: Optional[int] = None):
    for item in enum_visible_top_windows():
        if item["class"] != DIALOG_CLASS:
            continue
        if main_pid is not None and item["pid"] != main_pid:
            continue

        title = (item["title"] or "").strip()
        if title in {"提示信息", PROMPT_DIALOG_TITLE}:
            return item

    return None

def wait_for_success_prompt(main_pid: Optional[int] = None, timeout: float = SUCCESS_PROMPT_TIMEOUT):
    deadline = time.time() + timeout

    while time.time() < deadline:
        hit = find_success_prompt_dialog(main_pid=main_pid)
        if hit:
            return hit
        time.sleep(SCAN_INTERVAL)

    raise RuntimeError("Timed out waiting for export success prompt")


def find_success_prompt_dialog_once(main_pid: Optional[int] = None):
    hits = []

    for item in enum_visible_top_windows():
        if item["class"] != DIALOG_CLASS:
            continue
        if main_pid is not None and item["pid"] != main_pid:
            continue

        title = (item["title"] or "").strip()
        if title in {"提示信息", "鎻愮ず淇℃伅", PROMPT_DIALOG_TITLE}:
            if window_contains_success_text(item["hwnd"]):
                hits.append(item)
            continue

        if window_contains_success_text(item["hwnd"]):
            hits.append(item)

    if hits:
        return hits[0]

    return None


def find_success_prompt_dialog_resilient(main_pid: Optional[int] = None):
    title_hits = []

    for item in enum_visible_top_windows():
        if item["class"] != DIALOG_CLASS:
            continue
        if main_pid is not None and item["pid"] != main_pid:
            continue

        title = (item["title"] or "").strip()
        if window_contains_success_text(item["hwnd"]):
            return item

        # Some runs show the success prompt visually before its child texts are readable via UIA.
        if title in {"\u63d0\u793a\u4fe1\u606f", "鎻愮ず淇℃伅", PROMPT_DIALOG_TITLE}:
            title_hits.append(item)

    if title_hits:
        return title_hits[0]

    return None


def find_success_prompt_dialog_lightweight(main_pid: Optional[int] = None):
    for item in enum_visible_top_windows():
        if item["class"] != DIALOG_CLASS:
            continue
        if main_pid is not None and item["pid"] != main_pid:
            continue

        title = (item["title"] or "").strip()
        if title in {"提示信息", PROMPT_DIALOG_TITLE}:
            return item

    return find_success_prompt_dialog_resilient(main_pid=main_pid)


def close_success_prompt_by_alt_n(prompt_hwnd: int) -> None:
    dialog = get_desktop().window(handle=prompt_hwnd)
    set_focus_best(dialog)
    send_keys("%n")
    log("Sent Alt+N")


def close_success_prompt_by_alt_n_blind() -> None:
    send_keys("%n")
    log("Sent Alt+N blindly after prompt timeout")


def get_file_signature(path: Path) -> tuple[float, int] | None:
    if not path.exists():
        return None
    stat = path.stat()
    return stat.st_mtime, stat.st_size


def build_expected_export_filename(company_name: str, month: int) -> str:
    month = int(month)
    if not 1 <= month <= 12:
        raise ValueError(f"Month out of range: {month}")
    return f"{REPORT_NAME}-{company_name}(2024\u5e741\u6708 -2024\u5e74{month}\u6708).xlsx"


def find_matching_export_file(company_name: str, month: int, before_signature: tuple[float, int] | None) -> Path | None:
    month = int(month)
    candidates: list[tuple[float, Path]] = []

    for pattern in ("*.xlsx", "*.xls"):
        for path in EXPORT_DIR.rglob(pattern):
            try:
                parsed = parse_export_filename(path.name)
            except Exception:
                continue

            if parsed["company_name"] != company_name:
                continue
            if int(parsed["month"]) != month:
                continue

            signature = get_file_signature(path)
            if signature is None:
                continue
            if before_signature is not None and signature == before_signature:
                continue
            candidates.append((signature[0], path))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def wait_for_export_file(
    expected_path: Path,
    before_signature: tuple[float, int] | None,
    *,
    company_name: str,
    month: int,
) -> Path:
    deadline = time.time() + EXPORT_FILE_TIMEOUT

    while time.time() < deadline:
        current_signature = get_file_signature(expected_path)
        if current_signature is not None:
            if before_signature is None or current_signature != before_signature:
                return expected_path

        matched_path = find_matching_export_file(company_name, month, before_signature)
        if matched_path is not None:
            return matched_path

        time.sleep(EXPORT_FILE_POLL_INTERVAL)

    raise RuntimeError(f"Timed out waiting for export file: {expected_path}")


def find_export_file_once(
    expected_path: Path,
    before_signature: tuple[float, int] | None,
    *,
    company_name: str,
    month: int,
) -> Path | None:
    current_signature = get_file_signature(expected_path)
    if current_signature is not None:
        if before_signature is None or current_signature != before_signature:
            return expected_path

    return find_matching_export_file(company_name, month, before_signature)


def export_multiyear_general_ledger(company_name: str, month: int) -> tuple[str, Path]:
    export_filename = build_expected_export_filename(company_name=company_name, month=month)
    export_path = EXPORT_DIR / export_filename
    before_signature = get_file_signature(export_path)

    prepared = retry_window_operation(
        prepare_export_ledger_window,
        action_name="Prepare export ledger window",
        retry_times=EXPORT_ENTRY_RETRY_TIMES,
        retry_interval_sec=EXPORT_ENTRY_RETRY_INTERVAL_SEC,
    )
    if prepared is None:
        raise RuntimeError("Prepare export ledger window failed after retries")

    win, main_pid = prepared

    attempt = 0
    save_hit = None
    attempt += 1
    log(f"Export attempt {attempt}")
    click_export_button(win)

    log(f"Scan save dialog -> timeout={EXPORT_INITIAL_SAVE_DIALOG_WAIT:.0f}s")
    save_hit = wait_for_save_dialog(timeout=EXPORT_INITIAL_SAVE_DIALOG_WAIT, main_pid=main_pid)

    while save_hit is None:
        log("Save dialog not found yet, retry export after interval")

        ledger_win = try_get_ledger_window(timeout=1.0)
        if ledger_win is not None:
            set_focus_best(ledger_win)
            attempt += 1
            log(f"Export attempt {attempt}")
            click_export_button(ledger_win)
        else:
            log("Export retry skipped click | ledger window not ready")

        log(f"Scan save dialog -> timeout={EXPORT_RETRY_INTERVAL:.0f}s")
        save_hit = wait_for_save_dialog(timeout=EXPORT_RETRY_INTERVAL, main_pid=main_pid)

    log(f"Save dialog detected -> {save_hit}")

    save_file_by_alt_s(save_hit["hwnd"])
    time.sleep(SAVE_AFTER_ALT_S_WAIT)

    success_hit = retry_window_operation(
        lambda: find_success_prompt_dialog_lightweight(main_pid=main_pid),
        action_name="Wait for export success prompt",
        retry_times=NON_SENSITIVE_RETRY_TIMES,
        retry_interval_sec=NON_SENSITIVE_RETRY_INTERVAL_SEC,
    )
    blind_alt_n_fallback = False
    if success_hit is None:
        blind_alt_n_fallback = True
        log("Export success prompt not found after retries | fallback=blind Alt+N")
        close_success_prompt_by_alt_n_blind()
    else:
        log(f"Export success prompt detected -> {success_hit}")
        close_success_prompt_by_alt_n(success_hit["hwnd"])
    time.sleep(SUCCESS_AFTER_ALT_N_WAIT)

    final_export_path = retry_window_operation(
        lambda: find_export_file_once(
            expected_path=export_path,
            before_signature=before_signature,
            company_name=company_name,
            month=month,
        ),
        action_name="Wait for export file",
        retry_times=NON_SENSITIVE_RETRY_TIMES,
        retry_interval_sec=NON_SENSITIVE_RETRY_INTERVAL_SEC,
    )
    if final_export_path is None:
        if blind_alt_n_fallback:
            log(f"Export file not confirmed after blind Alt+N | fallback expected_path={export_path}")
            return export_filename, export_path
        raise RuntimeError(f"Export file not found after retries: {export_path}")

    return export_filename, final_export_path

