import time
from modules._shared.config import EXPORT_DIR, TARGET_MONTHS, TARGET_SOURCE_PERIOD_KEY
from modules._shared.logger import (
    get_export_csv_path,
    rebuild_export_csv_from_export_dir,
    write_export_success,
    write_log,
)
from modules._shared.progress import explain_skip_decision
from modules.general_ledger.expand_all_rows import expand_all_with_data_only
from modules.general_ledger.export_excel import export_multiyear_general_ledger
from modules.general_ledger.open_multiyear_ledger import enter_multiyear_general_ledger
from modules.general_ledger.select_month_and_query import select_month_and_query
from modules.tree_handler.project_entry import (
    activate_win,
    connect_project_list_win,
    enter_first_pending_project,
    load_tree_rows,
    pick_first_pending_row,
)
from modules.tree_handler.reset_project_tree import reset_project_tree_for_next_round
from modules.tree_handler.tree_observer import export_project_tree_observe_csv


ROW_PROJECT_KEY = "project_name"
ROW_COMPANY_KEY = "company_name"
ROW_PERIOD_KEY = "period_key"
ROW_LEVEL_KEY = "level"

INTER_MONTH_COOLDOWN_SECONDS = 5.0


def rebuild_progress_index() -> bool:
    write_log("PROGRESS_INDEX", "RUN", f"Rebuild export progress from {EXPORT_DIR}")
    try:
        output_csv = rebuild_export_csv_from_export_dir(
            export_dir=EXPORT_DIR,
            output_csv=get_export_csv_path(),
            source_period_key=TARGET_SOURCE_PERIOD_KEY,
        )
    except Exception as exc:
        write_log("PROGRESS_INDEX", "FAIL", f"Rebuild export progress failed | {exc}")
        return False

    write_log("PROGRESS_INDEX", "OK", f"Export progress ready | {output_csv}")
    return True


def scan_project_tree() -> bool:
    write_log("TREE_SCAN", "RUN", "Activate project list window and export tree csv")

    try:
        project_list_win = connect_project_list_win()
    except Exception as exc:
        write_log("TREE_SCAN", "FAIL", f"Connect project list window failed | {exc}")
        return False

    if not activate_win(project_list_win):
        write_log("TREE_SCAN", "FAIL", "Activate project list window failed")
        return False

    try:
        output_csv = export_project_tree_observe_csv()
    except Exception as exc:
        write_log("TREE_SCAN", "FAIL", f"Export tree csv failed | {exc}")
        return False

    write_log("TREE_SCAN", "OK", f"Tree csv ready | {output_csv}")
    return True


def export_current_project_months(target_row: dict) -> bool:
    company_name = str(target_row.get(ROW_COMPANY_KEY, "")).strip()

    decision = explain_skip_decision(
        company_name=company_name,
        source_period_key=TARGET_SOURCE_PERIOD_KEY,
        required_months=TARGET_MONTHS,
    )
    months_to_export = decision["missing_months"]
    if not months_to_export:
        write_log("EXPORT_LEDGER", "OK", f"All target months already complete | company={company_name}")
        return True

    write_log("OPEN_LEDGER", "RUN", f"Open multiyear general ledger | company={company_name}")
    try:
        opened = enter_multiyear_general_ledger()
    except Exception as exc:
        write_log("OPEN_LEDGER", "FAIL", f"Open multiyear general ledger failed | company={company_name} | {exc}")
        return False
    if not opened:
        write_log("OPEN_LEDGER", "FAIL", f"Open multiyear general ledger returned false | company={company_name}")
        return False
    write_log("OPEN_LEDGER", "OK", f"Multiyear general ledger opened | company={company_name}")

    for month_index, month in enumerate(months_to_export):
        write_log("MONTH_QUERY", "RUN", f"Select month and query | company={company_name} | month={month}")
        try:
            query_ok = select_month_and_query(month)
        except Exception as exc:
            write_log("MONTH_QUERY", "FAIL", f"Select month and query failed | company={company_name} | month={month} | {exc}")
            return False
        if not query_ok:
            write_log("MONTH_QUERY", "FAIL", f"Select month and query returned false | company={company_name} | month={month}")
            return False
        write_log("MONTH_QUERY", "OK", f"Month query completed | company={company_name} | month={month}")

        write_log("EXPAND_ROWS", "RUN", f"Expand ledger rows | company={company_name} | month={month}")
        try:
            expand_ok = expand_all_with_data_only()
        except Exception as exc:
            write_log("EXPAND_ROWS", "FAIL", f"Expand ledger rows failed | company={company_name} | month={month} | {exc}")
            return False
        if not expand_ok:
            write_log("EXPAND_ROWS", "FAIL", f"Expand ledger rows returned false | company={company_name} | month={month}")
            return False
        write_log("EXPAND_ROWS", "OK", f"Ledger rows expanded | company={company_name} | month={month}")

        write_log("EXPORT_LEDGER", "RUN", f"Export ledger to Excel | company={company_name} | month={month}")
        try:
            export_filename, export_path = export_multiyear_general_ledger(company_name=company_name, month=month)
        except Exception as exc:
            write_log("EXPORT_LEDGER", "FAIL", f"Export ledger failed | company={company_name} | month={month} | {exc}")
            return False

        write_export_success(
            step="EXPORT_LEDGER",
            export_filename=export_filename,
            export_path=str(export_path),
            source_period_key=TARGET_SOURCE_PERIOD_KEY,
            message=f"Month {month} export completed | company={company_name}",
        )

        if month_index < len(months_to_export) - 1:
            next_month = months_to_export[month_index + 1]
            write_log(
                "MONTH_COOLDOWN",
                "RUN",
                f"Wait {INTER_MONTH_COOLDOWN_SECONDS:.0f}s before next month query | company={company_name} | current_month={month} | next_month={next_month}",
            )
            time.sleep(INTER_MONTH_COOLDOWN_SECONDS)
            write_log(
                "MONTH_COOLDOWN",
                "OK",
                f"Cooldown finished | company={company_name} | next_month={next_month}",
            )

    return True


def reset_project_tree_after_export(project_name: str) -> bool:
    write_log("RESET_PROJECT_TREE", "RUN", f"Open change project dialog and reset tree | project={project_name}")

    if not reset_project_tree_for_next_round(root_name=project_name or None):
        write_log("RESET_PROJECT_TREE", "FAIL", "Open change project dialog or reset tree failed")
        return False

    write_log("RESET_PROJECT_TREE", "OK", "Project list reopened and tree reset to root-only state")
    return True


def find_next_pending_target() -> dict | None:
    rows = load_tree_rows()
    return pick_first_pending_row(rows)


def main():
    write_log("MAIN", "START", "Start Xiangyuan multiyear general ledger export flow")

    if not rebuild_progress_index():
        return

    if not scan_project_tree():
        return

    round_index = 1
    completed_units = 0

    while True:
        write_log("ROUND", "RUN", f"Start round {round_index}")

        pending_row = find_next_pending_target()
        if pending_row is None:
            write_log("MAIN", "OK", f"All pending units finished | completed_units={completed_units}")
            return

        target_row = enter_first_pending_project()
        if target_row is None:
            write_log("PROJECT_ENTRY", "FAIL", f"Pending unit exists but project entry failed | round={round_index}")
            return

        project_name = str(target_row.get(ROW_PROJECT_KEY, "")).strip()
        company_name = str(target_row.get(ROW_COMPANY_KEY, "")).strip()
        period_key = str(target_row.get(ROW_PERIOD_KEY, "")).strip()
        level = str(target_row.get(ROW_LEVEL_KEY, "")).strip()
        write_log(
            "PROJECT_ENTRY",
            "OK",
            f"Project entered | company={company_name} | period={period_key} | level={level} | round={round_index}",
        )

        if not export_current_project_months(target_row):
            return

        if not reset_project_tree_after_export(project_name):
            return

        completed_units += 1
        write_log(
            "ROUND",
            "OK",
            f"Round {round_index} completed | project={project_name} | company={company_name} | period={period_key}",
        )
        round_index += 1


if __name__ == "__main__":
    main()