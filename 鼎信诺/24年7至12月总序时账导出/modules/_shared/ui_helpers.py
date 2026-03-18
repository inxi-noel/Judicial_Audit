from __future__ import annotations

from typing import Iterable, Tuple

from pywinauto.mouse import click


def get_text(ctrl) -> str:
    try:
        return (ctrl.window_text() or "").strip()
    except Exception:
        return ""


def get_type(ctrl) -> str:
    try:
        return ctrl.element_info.control_type or ""
    except Exception:
        return ""


def get_rect(ctrl):
    try:
        return ctrl.rectangle()
    except Exception:
        return None


def rect_valid(rect) -> bool:
    return rect is not None and rect.left < rect.right and rect.top < rect.bottom


def rect_equal(rect, rect_tuple: Tuple[int, int, int, int]) -> bool:
    return (
        rect.left == rect_tuple[0]
        and rect.top == rect_tuple[1]
        and rect.right == rect_tuple[2]
        and rect.bottom == rect_tuple[3]
    )


def rect_intersects(a, b) -> bool:
    return not (
        a.right <= b.left or
        a.left >= b.right or
        a.bottom <= b.top or
        a.top >= b.bottom
    )


def rect_center(rect) -> tuple[int, int]:
    return (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2


def rect_center_tuple(rect_tuple: Tuple[int, int, int, int]) -> tuple[int, int]:
    left, top, right, bottom = rect_tuple
    return (left + right) // 2, (top + bottom) // 2


def rect_to_tuple(rect) -> tuple[int, int, int, int]:
    return rect.left, rect.top, rect.right, rect.bottom


def is_rect_visible(ctrl) -> bool:
    rect = get_rect(ctrl)
    if rect is None:
        return False
    try:
        return rect.width() > 5 and rect.height() > 5
    except Exception:
        return False


def click_control(ctrl, fallback_rect: Tuple[int, int, int, int] | None = None) -> str:
    try:
        ctrl.click_input()
        return "click_input"
    except Exception:
        if fallback_rect is None:
            raise
        click(coords=rect_center_tuple(fallback_rect))
        return "mouse_click"


def find_control_by_rule(
    controls: Iterable,
    *,
    target_name: str,
    target_type: str,
    target_rect: Tuple[int, int, int, int] | None = None,
):
    for ctrl in controls:
        if get_type(ctrl) != target_type:
            continue
        if get_text(ctrl) != target_name:
            continue

        rect = get_rect(ctrl)
        if not rect_valid(rect):
            continue
        if target_rect is not None and not rect_equal(rect, target_rect):
            continue

        return ctrl, rect

    raise RuntimeError(
        f"Control not found | name={target_name!r} | type={target_type!r} | rect={target_rect}"
    )


def try_find_control_by_rule(
    controls: Iterable,
    *,
    target_name: str,
    target_type: str,
    target_rect: Tuple[int, int, int, int] | None = None,
):
    try:
        return find_control_by_rule(
            controls,
            target_name=target_name,
            target_type=target_type,
            target_rect=target_rect,
        )
    except Exception:
        return None
