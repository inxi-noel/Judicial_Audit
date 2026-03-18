from __future__ import annotations

import time

from pywinauto import Desktop

from modules._shared.config import MAIN_WINDOW_REQUIRED_TEXT, MAIN_WINDOW_TITLE_RE
from modules._shared.ui_helpers import get_rect, rect_valid


POLL_INTERVAL = 0.2


def _window_area(win) -> int:
    rect = get_rect(win)
    if not rect_valid(rect):
        return -1
    return max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)


def _iter_main_window_candidates(required_texts: tuple[str, ...], forbidden_texts: tuple[str, ...]):
    desktop = Desktop(backend="uia")

    try:
        windows = desktop.windows(title_re=MAIN_WINDOW_TITLE_RE, top_level_only=True, visible_only=True)
    except Exception:
        windows = []

    for win in windows:
        try:
            title = (win.window_text() or "").strip()
            class_name = getattr(win.element_info, "class_name", "") or ""
            handle = win.handle
        except Exception:
            continue

        if required_texts and not all(text in title for text in required_texts if text):
            continue
        if forbidden_texts and any(text in title for text in forbidden_texts if text):
            continue
        if class_name == "Ghost":
            continue

        yield handle, _window_area(win)


def get_main_window(
    timeout: float = 10.0,
    *,
    required_texts: tuple[str, ...] | None = None,
    forbidden_texts: tuple[str, ...] | None = None,
):
    deadline = time.time() + timeout
    desktop = Desktop(backend="uia")
    required_texts = required_texts if required_texts is not None else ((MAIN_WINDOW_REQUIRED_TEXT,) if MAIN_WINDOW_REQUIRED_TEXT else ())
    forbidden_texts = forbidden_texts or ()

    while time.time() < deadline:
        candidates = list(_iter_main_window_candidates(required_texts, forbidden_texts))
        if candidates:
            candidates.sort(key=lambda item: item[1], reverse=True)
            handle = candidates[0][0]
            win = desktop.window(handle=handle)
            win.wait("visible", timeout=1.0)
            return win
        time.sleep(POLL_INTERVAL)

    raise RuntimeError("Main window not found within timeout")