from __future__ import annotations

import time
from typing import Callable, TypeVar

from pywinauto import Desktop


T = TypeVar("T")

WINDOW_RETRY_TIMES = 5
WINDOW_RETRY_INTERVAL_SEC = 3.0
WINDOW_CONNECT_TIMEOUT_SEC = 1.0


def retry_window_operation(
    action: Callable[[], T],
    *,
    action_name: str,
    retry_times: int = WINDOW_RETRY_TIMES,
    retry_interval_sec: float = WINDOW_RETRY_INTERVAL_SEC,
) -> T | None:
    last_error: Exception | None = None

    for attempt in range(1, retry_times + 1):
        try:
            result = action()
            if result not in (None, False):
                if attempt > 1:
                    print(f"{action_name} succeeded on attempt {attempt}/{retry_times}")
                return result
        except Exception as exc:
            last_error = exc
            print(f"{action_name} failed on attempt {attempt}/{retry_times}: {exc}")

        if attempt < retry_times:
            print(
                f"{action_name} retrying after {retry_interval_sec:.0f}s "
                f"({attempt}/{retry_times})"
            )
            time.sleep(retry_interval_sec)

    if last_error is not None:
        print(f"{action_name} failed after {retry_times} attempts: {last_error}")
    else:
        print(f"{action_name} failed after {retry_times} attempts")
    return None


def connect_uia_window(
    *,
    action_name: str,
    title: str | None = None,
    title_re: str | None = None,
    class_name: str | None = None,
    timeout: float = WINDOW_CONNECT_TIMEOUT_SEC,
):
    def action():
        kwargs = {}
        if title is not None:
            kwargs["title"] = title
        if title_re is not None:
            kwargs["title_re"] = title_re
        if class_name is not None:
            kwargs["class_name"] = class_name

        win = Desktop(backend="uia").window(**kwargs)
        win.wait("exists enabled visible", timeout=timeout)
        return win

    return retry_window_operation(action, action_name=action_name)


def connect_win32_window(
    *,
    title: str,
    action_name: str,
):
    def action():
        return Desktop(backend="win32").window(title=title).wrapper_object()

    return retry_window_operation(action, action_name=action_name)


def find_child_window_by_text(
    parent,
    target_text: str,
    *,
    action_name: str,
):
    def action():
        for child in parent.children():
            try:
                if child.window_text() == target_text:
                    return child
            except Exception:
                continue
        return None

    return retry_window_operation(action, action_name=action_name)
