from __future__ import annotations

from pathlib import Path
from datetime import datetime
import csv
import re
import sys
from typing import Any

# Allow direct execution of this file while keeping package imports unchanged
# when imported via the normal application entrypoints.
if __package__ in (None, ""):
    project_root = Path(__file__).resolve().parents[2]
    project_root_str = str(project_root)
    if project_root_str not in sys.path:
        sys.path.insert(0, project_root_str)

from modules._shared.window_retry import connect_uia_window


YEAR_RE = re.compile(r"(20\d{2})年")

#导出当前项目下的所有一二三四级树目录及相关数据到日志路径
# =========================================================
# 路径工具
# =========================================================
def get_desktop_log_dir() -> Path:
    desktop = Path.home() / "Desktop"
    log_dir = desktop / "鼎信诺导账套日志"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def get_tree_observe_csv_path() -> Path:
    """
    项目树观测结果输出路径
    固定文件名，便于后续其它模块直接读取
    """
    return get_desktop_log_dir() / "units_observe.csv"


def get_tree_observe_snapshot_csv_path() -> Path:
    """
    如果你后面想保留历史快照，可以用这个
    目前不是主输出，仅预留
    """
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    return get_desktop_log_dir() / f"units_observe_{now}.csv"


# =========================================================
# 基础工具
# =========================================================
def is_year(text: str) -> bool:
    return bool(YEAR_RE.search(text))


def get_year(text: str) -> str:
    m = YEAR_RE.search(text)
    return m.group(1) if m else ""


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


def rect_val(rect: Any, attr: str, default: Any = "") -> Any:
    if rect is None:
        return default
    return getattr(rect, attr, default)


def write_tree_observe_csv(rows: list[dict], output_csv: Path) -> Path:
    fieldnames = [
        "执行序号",
        "原始顺序",
        "项目名称",
        "年度",
        "年份节点文本",
        "三级名称",
        "节点文本",
        "单位名称",
        "父级名称",
        "级",
        "角色猜测",
        "判级依据",
        "是否年份",
        "是否年份后首节点",
        "路径键",
        "父路径键",
        "上一节点文本",
        "上一节点级别猜测",
        "left",
        "top",
        "right",
        "bottom",
        "width",
        "height",
        "left差值",
        "top差值",
        "缩进猜测",
        "页码猜测",
    ]

    output_csv.parent.mkdir(parents=True, exist_ok=True)

    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    return output_csv


# =========================================================
# 窗口与树扫描
# =========================================================
def connect_project_list_window(
    title_re: str = ".*项目列表.*",
    class_name: str = "FNWND3105",
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
    """
    扫描当前项目列表窗口中的 TreeItem
    注意：这里依赖 UIA 当前可暴露的顺序，不做排序
    """
    return win.descendants(control_type="TreeItem")


# =========================================================
# 核心解析
# =========================================================
def build_tree_observe_rows(items: list) -> list[dict]:
    rows: list[dict] = []

    project = None
    year = None
    year_node_text = None
    level3 = None
    seq = 0

    prev_text = ""
    prev_left = None
    prev_top = None
    prev_level_guess = ""
    page_index_guess = 1

    for raw_idx, ctrl in enumerate(items, 1):
        text = safe_text(ctrl)
        if not text:
            continue

        rect = safe_rect(ctrl)
        left = rect_val(rect, "left")
        top = rect_val(rect, "top")
        right = rect_val(rect, "right")
        bottom = rect_val(rect, "bottom")

        width = ""
        height = ""
        if rect is not None:
            try:
                width = rect.width()
                height = rect.height()
            except Exception:
                pass

        top_delta = ""
        if prev_top not in (None, "") and top not in ("", None):
            top_delta = top - prev_top

        left_delta = ""
        if prev_left not in (None, "") and left not in ("", None):
            left_delta = left - prev_left

        is_year_flag = 1 if is_year(text) else 0
        year_value = get_year(text) if is_year_flag else ""

        if project is None:
            project = text
            role_guess = "L1"
            level_guess = 1
            parent_name = ""
            parent_path_key = ""
            current_path_key = f"{project}"
            level3 = None
            year = None
            year_node_text = None
            is_first_after_year = 0
            judge_basis = "首个非空节点视为一级"

        elif is_year_flag:
            year = year_value
            year_node_text = text
            level3 = None
            role_guess = "L2_YEAR"
            level_guess = 2
            parent_name = project
            parent_path_key = f"{project}"
            current_path_key = f"{project}|{year}"
            is_first_after_year = 0
            judge_basis = "命中年份正则，视为二级年份节点"

        elif not year:
            role_guess = "UNKNOWN_BEFORE_YEAR"
            level_guess = ""
            parent_name = ""
            parent_path_key = ""
            current_path_key = f"{project}|UNRESOLVED|{text}"
            is_first_after_year = 0
            judge_basis = "尚未进入年份上下文，无法稳定判级"

        elif level3 is None:
            level3 = text
            role_guess = "L3_GUESS"
            level_guess = 3
            parent_name = f"{year}年"
            parent_path_key = f"{project}|{year}"
            current_path_key = f"{project}|{year}|{text}"
            is_first_after_year = 1
            judge_basis = "年份后的首个非年份节点，暂判三级"

        else:
            role_guess = "L4_GUESS"
            level_guess = 4
            parent_name = level3
            parent_path_key = f"{project}|{year}|{level3}"
            current_path_key = f"{project}|{year}|{level3}|{text}"
            is_first_after_year = 0
            judge_basis = "当前三级上下文后的后续节点，暂判四级"

        indent_guess = ""
        if left not in ("", None):
            if project == text and level_guess == 1:
                indent_guess = "ROOT"
            elif level_guess == 2:
                indent_guess = "YEAR_INDENT"
            elif level_guess == 3:
                indent_guess = "L3_INDENT"
            elif level_guess == 4:
                indent_guess = "L4_INDENT"

        seq += 1
        rows.append({
            "执行序号": seq,
            "原始顺序": raw_idx,
            "项目名称": project,
            "年度": year or "",
            "年份节点文本": year_node_text or "",
            "三级名称": level3 or "",
            "节点文本": text,
            "单位名称": text if level_guess in (3, 4) else "",
            "父级名称": parent_name,
            "级": level_guess,
            "角色猜测": role_guess,
            "判级依据": judge_basis,
            "是否年份": is_year_flag,
            "是否年份后首节点": is_first_after_year,
            "路径键": current_path_key,
            "父路径键": parent_path_key,
            "上一节点文本": prev_text,
            "上一节点级别猜测": prev_level_guess,
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
            "width": width,
            "height": height,
            "left差值": left_delta,
            "top差值": top_delta,
            "缩进猜测": indent_guess,
            "页码猜测": page_index_guess,
        })

        prev_text = text
        prev_left = left
        prev_top = top
        prev_level_guess = level_guess

    return rows


# =========================================================
# 对外主入口
# =========================================================
def export_project_tree_observe_csv(
    output_csv: Path | None = None,
    title_re: str = ".*项目列表.*",
    class_name: str = "FNWND3105",
) -> Path:
    """
    导出当前“项目列表”树结构观测 CSV

    返回：
        输出文件路径 Path
    """
    if output_csv is None:
        output_csv = get_tree_observe_csv_path()

    win = connect_project_list_window(title_re=title_re, class_name=class_name)
    items = scan_project_tree_items(win)
    rows = build_tree_observe_rows(items)
    result_path = write_tree_observe_csv(rows, output_csv)

    print(f"生成记录: {len(rows)}")
    print(f"输出: {result_path}")

    return result_path


def main():
    export_project_tree_observe_csv()


if __name__ == "__main__":
    main()
