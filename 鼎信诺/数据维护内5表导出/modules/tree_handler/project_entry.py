from __future__ import annotations

import csv
import re
import time
from typing import Any

from pywinauto import Desktop

from modules._shared.window_retry import connect_uia_window

from modules.tree_handler.tree_observer import get_tree_observe_csv_path
from modules.tree_handler.export_progress_checker import explain_skip_decision


YEAR_RE = re.compile(r"(20\d{2})年")

PROJECT_LIST_TITLE_RE = ".*项目列表.*"
PROJECT_LIST_CLASS_NAME = "FNWND3105"
MAIN_WINDOW_TITLE = "鼎信诺审计系统V7.0 7100系列(单机版) - [主界面]"
ENTER_PROJECT_TEXT = "进入项目"

PAGE_DOWN_MAX = 200
SAME_PAGE_LIMIT = 3
ACTION_SLEEP = 0.6
DEBUG_PREVIEW_COUNT = 30
WAIT_PROJECT_LIST_CLOSE_SEC = 15
WAIT_MAIN_WINDOW_SEC = 20

COL_PROJECT_NAME = "\u9879\u76ee\u540d\u79f0"
COL_YEAR = "\u5e74\u5ea6"
COL_YEAR_NODE_TEXT = "\u5e74\u4efd\u8282\u70b9\u6587\u672c"
COL_LEVEL3_NAME = "\u4e09\u7ea7\u540d\u79f0"
COL_COMPANY_NAME = "\u5355\u4f4d\u540d\u79f0"
COL_LEVEL = "\u7ea7"
YEAR_SUFFIX = "\u5e74"



# =========================================================
# 基础工具
# =========================================================
def get_desktop() -> Desktop:
    return Desktop(backend="uia")


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


def is_year_text(text: str) -> bool:
    return bool(YEAR_RE.search(text))


# =========================================================
# 窗口连接
# =========================================================
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
    win = desktop.window(
        title_re=PROJECT_LIST_TITLE_RE,
        class_name=PROJECT_LIST_CLASS_NAME,
    )
    try:
        if win.exists(timeout=1):
            return win
    except Exception:
        pass
    return None


def find_main_window():
    return connect_uia_window(
        title=MAIN_WINDOW_TITLE,
        action_name="Find main window",
    )


def wait_main_window(timeout: int = WAIT_MAIN_WINDOW_SEC):
    _ = timeout
    return connect_uia_window(
        title=MAIN_WINDOW_TITLE,
        action_name="Wait for main window",
    )


def wait_project_list_closed(timeout: int = WAIT_PROJECT_LIST_CLOSE_SEC) -> bool:
    print("等待【项目列表】窗口关闭...")
    start = time.time()

    while time.time() - start < timeout:
        win = try_get_project_list_win()
        if win is None:
            print("【项目列表】窗口已关闭")
            return True
        time.sleep(0.5)

    print("超时：项目列表窗口仍未关闭")
    return False


# =========================================================
# CSV 读取与任务筛选
# =========================================================
def load_tree_rows(csv_path=None) -> list[dict]:
    if csv_path is None:
        csv_path = get_tree_observe_csv_path()

    rows = []
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_year(value: Any) -> str:
    text = normalize_text(value)
    if text.endswith("年"):
        text = text[:-1].strip()
    return text


def normalize_year_node_text(value: Any) -> str:
    return normalize_text(value)


def iter_unique_task_rows(rows: list[dict]):
    """
    Preserve CSV order, but dedupe on (company, year, exact year-node text)
    so split-year branches inside the same numeric year remain distinct tasks.
    """
    seen = set()

    for row in rows:
        company_name = normalize_text(row.get(COL_COMPANY_NAME, ""))
        year = normalize_year(row.get(COL_YEAR, ""))
        year_node_text = normalize_year_node_text(row.get(COL_YEAR_NODE_TEXT, ""))
        level = normalize_text(row.get(COL_LEVEL, ""))

        if level not in ("3", "4"):
            continue

        if not company_name or not year:
            continue

        key = (company_name, year, year_node_text)
        if key in seen:
            continue

        seen.add(key)
        yield row

def pick_first_pending_row(rows: list[dict]) -> dict | None:
    """
    Pick the first task row that is still incomplete for the current batch.
    """
    for row in iter_unique_task_rows(rows):
        company_name = normalize_text(row.get(COL_COMPANY_NAME, ""))
        year = normalize_year(row.get(COL_YEAR, ""))
        year_node_text = normalize_year_node_text(row.get(COL_YEAR_NODE_TEXT, ""))
        level = normalize_text(row.get(COL_LEVEL, ""))

        if level not in ("3", "4"):
            continue

        decision = explain_skip_decision(company_name, year, year_node_text=year_node_text)
        year_display = year_node_text or year

        if decision["should_skip"]:
            print(
                f"Skip: {company_name} | {year_display} | completed | "
                f"exported={decision['exported_tables']}"
            )
            continue

        print(
            f"Pending: {company_name} | {year_display} | missing={decision['missing_tables']}"
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

    visible.sort(key=lambda c: (
        safe_rect(c).top if safe_rect(c) else 999999,
        safe_rect(c).left if safe_rect(c) else 999999,
    ))

    if debug:
        print(f"Visible tree items: {len(visible)}")
        for i, ctrl in enumerate(visible[:DEBUG_PREVIEW_COUNT], 1):
            print(f"  [{i:02d}] {safe_text(ctrl)}")

    return visible


def page_signature(win) -> tuple:
    sig = []
    for ctrl in visible_items(win, debug=False):
        text = safe_text(ctrl)
        rect = safe_rect(ctrl)
        top = rect.top if rect else ""
        sig.append((text, top))
    return tuple(sig)


def single_click(ctrl, label: str = "") -> bool:
    try:
        ctrl.click_input()
        time.sleep(ACTION_SLEEP)
        if label:
            print(f"Single click OK: {label}")
        return True
    except Exception as e:
        if label:
            print(f"Single click failed: {label} | {e}")
        return False


def focus_tree_by_single_click(win) -> bool:
    if not activate_win(win):
        return False

    vis = visible_items(win, debug=False)
    if not vis:
        print("No visible tree item to focus")
        return False

    target = vis[0]
    text = safe_text(target)

    try:
        target.click_input()
        time.sleep(ACTION_SLEEP)
        print(f"Tree focused by click: {text!r}")
        return True
    except Exception as e:
        print(f"Tree focus click failed: {text!r} | {e}")
        return False


def verify_tree_item_selected(ctrl, label: str = "") -> bool:
    try:
        selected = bool(ctrl.is_selected())
    except Exception as e:
        if label:
            print(f"Read selected state failed: {label} | {e}")
        return False

    if selected:
        if label:
            print(f"Selection verified: {label}")
        return True

    if label:
        print(f"Selection verify failed: {label}")
    return False


def ensure_tree_item_selected(ctrl, label: str = "") -> bool:
    try:
        ctrl.set_focus()
        time.sleep(0.2)
    except Exception as e:
        if label:
            print(f"Set focus failed: {label} | {e}")

    try:
        ctrl.select()
        time.sleep(ACTION_SLEEP)
        if label:
            print(f"select() called: {label}")
    except Exception as e:
        if label:
            print(f"select() failed: {label} | {e}")

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
            print(f"双击成功: {label}")
        return True
    except Exception as e:
        if label:
            print(f"双击失败: {label} | {e}")
        return False


def send_pagedown(win) -> bool:
    if not activate_win(win):
        return False

    try:
        from pywinauto.keyboard import send_keys
        send_keys("{PGDN}")
        time.sleep(ACTION_SLEEP)
        print("Sent PageDown")
        return True
    except Exception as e:
        print("Send PageDown failed:", e)
        return False


def find_root_ctrl(win, project_name: str):
    vis = visible_items(win, debug=True)
    for ctrl in vis:
        if safe_text(ctrl) == project_name:
            return ctrl
    return None


def find_visible_year(win, year_value: str, year_node_text: str = "", debug: bool = False):
    exact_year_text = normalize_year_node_text(year_node_text)
    prefix = normalize_year(year_value) + YEAR_SUFFIX
    vis = visible_items(win, debug=debug)

    if exact_year_text:
        for ctrl in vis:
            if safe_text(ctrl) == exact_year_text:
                return ctrl

    for ctrl in vis:
        text = safe_text(ctrl)
        if is_year_text(text) and text.startswith(prefix):
            return ctrl
    return None

def is_collapsed_to_root_only(win, project_name: str) -> bool:
    vis = visible_items(win, debug=True)
    return len(vis) == 1 and safe_text(vis[0]) == project_name


def collapse_to_root_only(win, project_name: str, max_try: int = 3) -> bool:
    print("\n--- 开始重置树到只剩一级 ---")

    for i in range(max_try):
        if is_collapsed_to_root_only(win, project_name):
            print("树已处于只剩一级状态")
            return True

        root = find_root_ctrl(win, project_name)
        if root is None:
            print("重置失败：当前找不到一级项目节点")
            return False

        if not double_click(root, f"收起一级(第{i+1}次): {project_name}"):
            return False

        time.sleep(0.8)

        if is_collapsed_to_root_only(win, project_name):
            print("重置成功：已回到只剩一级状态")
            return True

    print("重置失败：多次尝试后仍未回到只剩一级状态")
    return False


def ensure_year_expanded(win, project_name: str, year_value: str, year_node_text: str = ""):
    root = find_root_ctrl(win, project_name)
    if root is None:
        print("Failed: root project node not found")
        return None

    target_year = normalize_year_node_text(year_node_text) or year_value
    year_ctrl = find_visible_year(win, year_value, year_node_text=year_node_text, debug=True)
    if year_ctrl is None:
        if not double_click(root, f"Expand root: {project_name}"):
            return None

        year_ctrl = find_visible_year(win, year_value, year_node_text=year_node_text, debug=True)
        if year_ctrl is None:
            print(f"Failed: year node not found after expanding root -> {target_year}")
            return None

    if not double_click(year_ctrl, f"Expand year: {target_year}"):
        return None

    return year_ctrl

def find_level3_after_year(win, year_value: str, level3_name: str, year_node_text: str = ""):
    target_year = normalize_year_node_text(year_node_text) or normalize_year(year_value)
    print(f"\n--- Find level-3 after year node: {level3_name} | year-node={target_year} ---")
    vis = visible_items(win, debug=True)

    exact_year_text = normalize_year_node_text(year_node_text)
    prefix = normalize_year(year_value) + YEAR_SUFFIX
    year_idx = -1

    if exact_year_text:
        for i, ctrl in enumerate(vis):
            if safe_text(ctrl) == exact_year_text:
                year_idx = i
                break

    if year_idx == -1:
        for i, ctrl in enumerate(vis):
            text = safe_text(ctrl)
            if is_year_text(text) and text.startswith(prefix):
                year_idx = i
                break

    if year_idx == -1:
        print("Failed: target year node is not visible on current page")
        return None

    for i in range(year_idx + 1, len(vis)):
        text = safe_text(vis[i])
        if is_year_text(text):
            break
        if text == level3_name:
            print(f"Found level-3 in year context, index={i + 1}")
            return vis[i]

    for i in range(year_idx + 1, len(vis)):
        text = safe_text(vis[i])
        if is_year_text(text):
            break
        print(f"Fallback to first non-year node after year context: {text}")
        return vis[i]

    print("Failed: no level-3 candidate found after target year node")
    return None

def find_level4_by_pagedown(win, level4_name: str):
    print(f"\n--- Start scanning level-4 target: {level4_name} ---")

    vis = visible_items(win, debug=True)
    for ctrl in vis:
        if safe_text(ctrl) == level4_name:
            print("Level-4 target already visible")
            return ctrl

    prev_sig = page_signature(win)
    same_count = 0

    for step in range(1, PAGE_DOWN_MAX + 1):
        print(f"\n[PageDown step {step}]")

        if not send_pagedown(win):
            print("PageDown could not be sent")
            return None

        vis = visible_items(win, debug=True)
        for ctrl in vis:
            if safe_text(ctrl) == level4_name:
                print("Level-4 target is now visible")
                return ctrl

        curr_sig = page_signature(win)

        if curr_sig == prev_sig:
            same_count += 1
            print(f"Page signature unchanged: {same_count}")
        else:
            same_count = 0
            print("Page signature changed; continue scanning")

        if same_count >= SAME_PAGE_LIMIT:
            print("Page signature unchanged repeatedly; stop paging")
            return None

        prev_sig = curr_sig

    print("Reached max PageDown attempts")
    return None


# =========================================================
# Enter project
# =========================================================
def find_enter_project_control(win):
    candidates = []

    try:
        all_ctrls = win.descendants()
    except Exception as e:
        print("扫描进入项目控件失败：", e)
        return None

    for ctrl in all_ctrls:
        try:
            text = safe_text(ctrl)
            if text != ENTER_PROJECT_TEXT:
                continue

            rect = safe_rect(ctrl)
            if rect is None:
                continue

            if rect.width() <= 3 or rect.height() <= 3:
                continue

            candidates.append(ctrl)
        except Exception:
            continue

    if not candidates:
        print("未找到【进入项目】控件")
        return None

    ctrl = candidates[0]
    rect = safe_rect(ctrl)
    print(f"找到【进入项目】控件，矩形={rect}")
    return ctrl


def click_enter_project(win) -> bool:
    ctrl = find_enter_project_control(win)
    if ctrl is None:
        return False

    try:
        ctrl.click_input()
        time.sleep(0.8)
        print("已点击【进入项目】")
        return True
    except Exception as e:
        print("点击【进入项目】失败：", e)

    try:
        rect = ctrl.rectangle()
        x = (rect.left + rect.right) // 2
        y = (rect.top + rect.bottom) // 2
        base_rect = win.rectangle()
        win.click_input(coords=(x - base_rect.left, y - base_rect.top))
        time.sleep(0.8)
        print("已通过坐标点击【进入项目】")
        return True
    except Exception as e:
        print("坐标点击【进入项目】失败：", e)
        return False


# =========================================================
# 主流程
# =========================================================
def select_tree_target(win, row: dict) -> bool:
    project_name = normalize_text(row.get(COL_PROJECT_NAME, ""))
    year_value = normalize_year(row.get(COL_YEAR, ""))
    year_node_text = normalize_year_node_text(row.get(COL_YEAR_NODE_TEXT, ""))
    level3_name = normalize_text(row.get(COL_LEVEL3_NAME, ""))
    unit_name = normalize_text(row.get(COL_COMPANY_NAME, ""))

    try:
        level_value = int(normalize_text(row.get(COL_LEVEL, "")))
    except Exception:
        print("Failed: CSV level value is not a valid integer")
        return False

    print("=" * 80)
    print("Prepare to enter project")
    print(f"project={project_name}")
    print(f"year={year_value}")
    print(f"year_node={year_node_text or year_value}")
    print(f"level3={level3_name}")
    print(f"company={unit_name}")
    print(f"level={level_value}")
    print("=" * 80)

    if not activate_win(win):
        print("Failed: window activation failed")
        return False

    if not collapse_to_root_only(win, project_name):
        print("Failed: tree could not be reset to root-only state")
        return False

    if ensure_year_expanded(win, project_name, year_value, year_node_text=year_node_text) is None:
        print("Failed: target year node expansion failed")
        return False

    if level_value == 3:
        level3_ctrl = find_level3_after_year(
            win,
            year_value,
            unit_name,
            year_node_text=year_node_text,
        )
        if level3_ctrl is None:
            return False

        if not ensure_tree_item_selected(level3_ctrl, f"Select level-3: {unit_name}"):
            print("Level-3 selection verification failed")
            return False

        print("Level-3 target selected")
        return True

    if level_value == 4:
        level3_ctrl = find_level3_after_year(
            win,
            year_value,
            level3_name,
            year_node_text=year_node_text,
        )
        if level3_ctrl is None:
            print("Level-3 parent not found on current page")
            return False

        if not double_click(level3_ctrl, f"Expand level-3: {level3_name}"):
            print("Expand level-3 failed")
            return False

        level4_ctrl = find_level4_by_pagedown(win, unit_name)
        if level4_ctrl is None:
            print("Level-4 target not found")
            return False

        if not ensure_tree_item_selected(level4_ctrl, f"Select level-4: {unit_name}"):
            print("Level-4 selection verification failed")
            return False

        print("Level-4 target selected")
        return True

    print(f"Unsupported level value: {level_value}")
    return False

def enter_first_pending_project() -> dict | None:
    """
    正式入口：
    1. 读取 units_observe.csv
    2. 找到第一个今天未完成导表的目标
    3. 在项目树中定位并进入项目
    4. 等待主界面打开
    5. 返回目标行 dict，供后续导表模块继续使用

    返回：
        None  -> 没有可处理项目，或进入失败
        dict  -> 已成功进入项目，对应的树行信息
    """
    rows = load_tree_rows()
    target_row = pick_first_pending_row(rows)
    if target_row is None:
        return None

    project_list_win = connect_project_list_win()

    if not select_tree_target(project_list_win, target_row):
        print("失败：目标选中阶段失败")
        return None

    if not click_enter_project(project_list_win):
        print("失败：点击进入项目失败")
        return None

    if not wait_project_list_closed():
        print("失败：项目列表未关闭")
        return None

    main_win = wait_main_window()
    if main_win is None:
        print("失败：未等到主界面出现")
        return None

    print("成功：已进入项目主界面 ->", repr(main_win.window_text()))
    return target_row


def main():
    row = enter_first_pending_project()
    if row is None:
        print("\n最终结果：没有进入任何项目")
        return

    print("\n最终结果：成功进入项目")
    print(
        f"当前目标: 项目={row.get('项目名称','')} | 年度={row.get('年度','')} | "
        f"单位={row.get('单位名称','')} | 级={row.get('级','')}"
    )


if __name__ == "__main__":
    main()
