from __future__ import annotations

import csv
import re
import time
from pathlib import Path
from typing import Any

from pywinauto import Desktop
from pywinauto.keyboard import send_keys

from modules._shared.window_retry import connect_uia_window, retry_window_operation


MAIN_WINDOW_TITLE = "鼎信诺审计系统V7.0 7100系列(单机版) - [主界面]"
PROJECT_LIST_TITLE_RE = ".*项目列表.*"
PROJECT_LIST_CLASS_NAME = "FNWND3105"

YEAR_RE = re.compile(r"(20\d{2})年")

ACTION_SLEEP = 0.6
PAGEUP_SLEEP = 0.8
MAX_PAGEUP_TIMES = 100
MAX_COLLAPSE_TRIES = 3
DEBUG_PREVIEW_COUNT = 50


# =========================================================
# 路径工具
# =========================================================
def get_desktop_log_dir() -> Path:
    desktop = Path.home() / "Desktop"
    log_dir = desktop / "鼎信诺导账套日志"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def get_tree_observe_csv_path() -> Path:
    return get_desktop_log_dir() / "units_observe.csv"


# =========================================================
# 基础工具
# =========================================================
def safe_text(ctrl: Any) -> str:
    try:
        return ctrl.window_text().strip()
    except Exception:
        return ""


def safe_rect(ctrl: Any):
    try:
        return ctrl.rectangle()
    except Exception:
        return None


def is_rect_visible(ctrl: Any) -> bool:
    rect = safe_rect(ctrl)
    if rect is None:
        return False
    try:
        return rect.width() > 5 and rect.height() > 5
    except Exception:
        return False


def is_year_text(text: str) -> bool:
    return bool(YEAR_RE.search(text))


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
    except Exception as e:
        print("窗口激活失败：", e)
        return False


# =========================================================
# CSV：读取唯一一级项目名称
# =========================================================
def load_root_project_name_from_csv(csv_path: Path | None = None) -> str:
    if csv_path is None:
        csv_path = get_tree_observe_csv_path()

    if not csv_path.exists():
        raise FileNotFoundError(f"未找到树观测 CSV：{csv_path}")

    project_names = []
    seen = set()

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = str(row.get("项目名称", "")).strip()
            if not name:
                continue
            if name in seen:
                continue
            seen.add(name)
            project_names.append(name)

    if not project_names:
        raise RuntimeError("CSV 中未读取到任何【项目名称】")

    if len(project_names) > 1:
        raise RuntimeError(f"CSV 中存在多个一级项目名称：{project_names}")

    return project_names[0]


# =========================================================
# 窗口连接
# =========================================================
def find_main_window():
    return connect_uia_window(
        title=MAIN_WINDOW_TITLE,
        action_name="Find main window",
    )


def connect_main_window():
    win = connect_uia_window(
        title=MAIN_WINDOW_TITLE,
        action_name="Connect main window",
    )
    if win is None:
        raise RuntimeError("Main window not found")
    return win


def connect_project_list_win(timeout: int = 10):
    _ = timeout
    win = connect_uia_window(
        title_re=PROJECT_LIST_TITLE_RE,
        class_name=PROJECT_LIST_CLASS_NAME,
        action_name="Connect project list window",
    )
    if win is None:
        raise RuntimeError("Project list window not found")
    return win


# =========================================================
# Open: project management -> change project
# =========================================================
def open_change_project_dialog() -> bool:
    try:
        main = connect_main_window()
    except Exception as e:
        print("连接主窗口失败：", e)
        return False

    if not activate_win(main):
        return False

    try:
        menu_bars = main.descendants(control_type="MenuBar")
    except Exception as e:
        print("扫描菜单栏失败：", e)
        return False

    def find_app_menu_bar():
        for mb in menu_bars:
            try:
                if mb.window_text() == "\u5e94\u7528\u7a0b\u5e8f":
                    return mb
            except Exception:
                continue
        return None

    app_menu_bar = retry_window_operation(
        find_app_menu_bar,
        action_name="Find application menu bar",
    )
    if app_menu_bar is None:
        print("Application menu bar not found")
        return False

    try:
        menu_items = app_menu_bar.children(control_type="MenuItem")
    except Exception as e:
        print("读取菜单项失败：", e)
        return False

    if len(menu_items) <= 1:
        print("菜单项数量异常，无法定位【项目管理】")
        return False

    try:
        menu_items[1].click_input()
        print("已点击【项目管理】主菜单（Item 1）")
    except Exception as e:
        print("点击【项目管理】失败：", e)
        return False

    time.sleep(0.5)

    try:
        send_keys("{DOWN 4}")
        time.sleep(0.2)
        send_keys("{ENTER}")
        time.sleep(1.0)
        print("已发送【DOWN 4 + ENTER】打开更换项目")
    except Exception as e:
        print("通过键盘选择【更换项目】失败：", e)
        return False

    try:
        connect_project_list_win(timeout=10)
        print("【项目列表】窗口已出现")
        return True
    except Exception as e:
        print("等待【项目列表】窗口出现失败：", e)
        return False


# =========================================================
# 树可见节点
# =========================================================
def visible_items(win, debug: bool = False) -> list:
    nodes = win.descendants(control_type="TreeItem")
    visible = []

    for ctrl in nodes:
        text = safe_text(ctrl)
        if not text:
            continue
        if is_rect_visible(ctrl):
            visible.append(ctrl)

    visible.sort(key=lambda c: (
        safe_rect(c).top if safe_rect(c) else 999999,
        safe_rect(c).left if safe_rect(c) else 999999,
    ))

    if debug:
        print(f"当前可见节点数量: {len(visible)}")
        for i, ctrl in enumerate(visible[:DEBUG_PREVIEW_COUNT], 1):
            rect = safe_rect(ctrl)
            print(f"  [{i:02d}] {safe_text(ctrl)} | rect={rect}")

    return visible


def page_signature(win) -> tuple:
    sig = []
    for ctrl in visible_items(win, debug=False):
        rect = safe_rect(ctrl)
        sig.append((
            safe_text(ctrl),
            rect.top if rect else None,
            rect.left if rect else None,
        ))
    return tuple(sig)


# =========================================================
# 树焦点初始化
# =========================================================
def focus_tree_by_single_click(win) -> bool:
    if not activate_win(win):
        return False

    vis = visible_items(win, debug=True)
    if not vis:
        print("失败：树中没有任何可单击的可见节点")
        return False

    target = vis[0]
    text = safe_text(target)

    try:
        target.click_input()
        time.sleep(0.5)
        print(f"已单击树节点，使树获得焦点：{text!r}")
        return True
    except Exception as e:
        print(f"单击树节点失败：{text!r} | {e}")
        return False


# =========================================================
# 一级判定
# =========================================================
def find_top_visible_year_index(vis: list) -> int:
    for i, ctrl in enumerate(vis):
        if is_year_text(safe_text(ctrl)):
            return i
    return -1


def find_root_ctrl_by_year_context(win, root_name: str, debug: bool = True):
    vis = visible_items(win, debug=debug)
    if not vis:
        return None

    if len(vis) == 1 and safe_text(vis[0]) == root_name:
        print("判定：当前页仅一个节点，且等于一级项目名称 -> 已收起一级")
        return vis[0]

    year_idx = find_top_visible_year_index(vis)
    if year_idx == -1:
        print("当前页未发现任何可见年份节点，无法通过上下文稳定判定一级")
        return None

    if year_idx == 0:
        print("最上面的可见节点就是年份，说明一级当前不在可见区")
        return None

    candidate = vis[year_idx - 1]
    candidate_text = safe_text(candidate)

    if candidate_text != root_name:
        print(
            f"年份节点上方节点文本不匹配一级名称："
            f"上方={candidate_text!r} | 目标一级={root_name!r}"
        )
        return None

    print(f"判定成功：最上方年份节点前一个节点为一级 -> {candidate_text!r}")
    return candidate


def is_collapsed_to_root_only(win, root_name: str, debug: bool = True) -> bool:
    vis = visible_items(win, debug=debug)
    return len(vis) == 1 and safe_text(vis[0]) == root_name


# =========================================================
# 操作封装
# =========================================================
def double_click(ctrl, label: str = "") -> bool:
    try:
        ctrl.click_input(double=True)
        time.sleep(ACTION_SLEEP)
        if label:
            print(f"双击成功: {label}")
        return True
    except Exception as e:
        if label:
            print(f"双击失败: {label} | {e}")
        return False


def send_pageup(win) -> bool:
    if not activate_win(win):
        return False

    try:
        send_keys("{PGUP}")
        time.sleep(PAGEUP_SLEEP)
        print("已发送 PageUp")
        return True
    except Exception as e:
        print("发送 PageUp 失败：", e)
        return False


# =========================================================
# PageUp 寻找一级上下文
# =========================================================
def pageup_until_root_context_visible(
    win,
    root_name: str,
    max_times: int = MAX_PAGEUP_TIMES,
) -> bool:
    print(f"\n--- 开始 PageUp，直到能通过年份上下文定位一级：{root_name} ---")

    prev_sig = page_signature(win)

    for i in range(1, max_times + 1):
        ctrl = find_root_ctrl_by_year_context(win, root_name, debug=True)
        if ctrl is not None:
            print("已成功在当前页建立一级判定上下文")
            return True

        print(f"[PageUp {i}/{max_times}] 当前页无法稳定判定一级，继续上翻")

        if not send_pageup(win):
            return False

        curr_sig = page_signature(win)
        if curr_sig == prev_sig:
            print("页面签名未变化，可能已到顶部或树未响应")
        prev_sig = curr_sig

    print("失败：达到最大 PageUp 次数后，仍无法建立一级判定上下文")
    return False


# =========================================================
# 树复位核心
# =========================================================
def collapse_tree_to_root_only(
    win,
    root_name: str,
    max_try: int = MAX_COLLAPSE_TRIES,
) -> bool:
    print("\n--- 开始执行树复位 ---")
    print(f"目标一级项目名称：{root_name}")

    for i in range(1, max_try + 1):
        print(f"\n[收起尝试 {i}/{max_try}]")

        if is_collapsed_to_root_only(win, root_name=root_name, debug=True):
            print("树已处于只剩一级状态")
            return True

        root_ctrl = find_root_ctrl_by_year_context(win, root_name, debug=True)
        if root_ctrl is None:
            print("当前页无法稳定识别一级，准备 PageUp 寻找上下文")

            ok = pageup_until_root_context_visible(win, root_name)
            if not ok:
                print("失败：无法将一级对应上下文翻回可见区域")
                return False

            root_ctrl = find_root_ctrl_by_year_context(win, root_name, debug=True)
            if root_ctrl is None:
                print("失败：理论上应已可判定一级，但再次读取仍失败")
                return False

        if not double_click(root_ctrl, f"收起一级项目: {root_name}"):
            return False

        time.sleep(0.8)

        if is_collapsed_to_root_only(win, root_name=root_name, debug=True):
            print("成功：树已恢复为只剩一级项目状态")
            return True

    print("失败：多次尝试后，树仍未恢复到只剩一级项目状态")
    return False


# =========================================================
# 对外入口
# =========================================================
def reset_project_tree_for_next_round(
    root_name: str | None = None,
    csv_path: Path | None = None,
) -> bool:
    """
    正式入口：
    1. 打开【项目管理 -> 更换项目】
    2. 单击树节点，让树获得焦点
    3. 将树重置到“只剩一级项目”的初始状态

    参数：
        root_name:
            可直接传入一级项目名称。
            若不传，则自动从 units_observe.csv 读取。
        csv_path:
            可选，自定义树观测 CSV 路径。

    返回：
        True  -> 树已成功复位
        False -> 失败
    """
    if root_name is None:
        try:
            root_name = load_root_project_name_from_csv(csv_path)
            print(f"自动读取一级项目名称：{root_name}")
        except Exception as e:
            print("读取 CSV 中的一级项目名称失败：", e)
            return False

    if not open_change_project_dialog():
        print("失败：未能打开【项目列表】窗口")
        return False

    try:
        win = connect_project_list_win(timeout=10)
    except Exception as e:
        print("连接【项目列表】窗口失败：", e)
        return False

    if not activate_win(win):
        return False

    print("\n--- 先单击一个树节点，让树获得焦点 ---")
    if not focus_tree_by_single_click(win):
        return False

    print("\n--- 单击后重新读取当前树状态 ---")
    vis = visible_items(win, debug=True)
    if not vis:
        print("失败：当前树区域没有读取到任何可见节点")
        return False

    if is_collapsed_to_root_only(win, root_name=root_name, debug=True):
        print("当前已是初始状态，无需处理")
        return True

    return collapse_tree_to_root_only(win, root_name=root_name)


def main():
    ok = reset_project_tree_for_next_round()
    print("\n最终结果：", "成功" if ok else "失败")


if __name__ == "__main__":
    main()