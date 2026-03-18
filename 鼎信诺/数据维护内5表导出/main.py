from modules._shared.logger import write_log, write_export_success
from modules._shared.close_data_maintenance_window import close_data_maintenance_window

from modules.tree_handler.tree_observer import export_project_tree_observe_csv
from modules.tree_handler.project_entry import (
    activate_win,
    connect_project_list_win,
    enter_first_pending_project,
    load_tree_rows,
    pick_first_pending_row,
)
from modules.tree_handler.reset_project_tree import reset_project_tree_for_next_round

from modules.export_subject_balance.open_data_maintenance import open_data_maintenance
from modules.export_subject_balance.select_subject_balance import select_subject_balance
from modules.export_subject_balance.trigger_export_menu import export_subject_balance

from modules.voucher_table.open_voucher_selector import open_voucher_selector
from modules.voucher_table.select_voucher_table import select_voucher_table
from modules.voucher_table.trigger_export_menu import export_voucher_table

from modules.foreign_currency_balance.open_foreign_currency_selector import open_foreign_currency_selector
from modules.foreign_currency_balance.select_foreign_currency_balance import select_foreign_currency_balance
from modules.foreign_currency_balance.trigger_export_menu import export_foreign_currency_balance

from modules.project_balance.open_project_balance_selector import open_project_balance_selector
from modules.project_balance.select_project_balance import select_project_balance
from modules.project_balance.trigger_export_menu import export_project_balance

from modules.project_detail.open_project_detail_selector import open_project_detail_selector
from modules.project_detail.select_project_detail import select_project_detail
from modules.project_detail.trigger_export_menu import export_project_detail


ROW_PROJECT_KEY = "\u9879\u76ee\u540d\u79f0"
ROW_COMPANY_KEY = "\u5355\u4f4d\u540d\u79f0"
ROW_YEAR_KEY = "\u5e74\u5ea6"
ROW_LEVEL_KEY = "\u7ea7"


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


def export_current_project_tables() -> bool:
    write_log("DATA_MAINTENANCE", "RUN", "Open data maintenance")
    if not open_data_maintenance():
        write_log("DATA_MAINTENANCE", "FAIL", "Open data maintenance failed")
        return False
    write_log("DATA_MAINTENANCE", "OK", "Data maintenance opened")

    write_log("EXPORT_SUBJECT_BALANCE", "RUN", "Select subject balance table")
    if not select_subject_balance():
        write_log("EXPORT_SUBJECT_BALANCE", "FAIL", "Select subject balance table failed")
        return False
    write_log("EXPORT_SUBJECT_BALANCE", "OK", "Subject balance table selected")

    try:
        filename = export_subject_balance()
        write_export_success("EXPORT_SUBJECT_BALANCE", filename, "Subject balance export completed")
    except Exception as exc:
        write_log("EXPORT_SUBJECT_BALANCE", "FAIL", f"Subject balance export failed | {exc}")
        return False

    write_log("EXPORT_VOUCHER", "RUN", "Open voucher selector")
    if not open_voucher_selector():
        write_log("EXPORT_VOUCHER", "FAIL", "Open voucher selector failed")
        return False

    if not select_voucher_table():
        write_log("EXPORT_VOUCHER", "FAIL", "Select voucher table failed")
        return False
    write_log("EXPORT_VOUCHER", "OK", "Voucher table selected")

    try:
        filename = export_voucher_table()
        write_export_success("EXPORT_VOUCHER", filename, "Voucher export completed")
    except Exception as exc:
        write_log("EXPORT_VOUCHER", "FAIL", f"Voucher export failed | {exc}")
        return False

    write_log("EXPORT_FOREIGN_CURRENCY", "RUN", "Open foreign currency selector")
    if not open_foreign_currency_selector():
        write_log("EXPORT_FOREIGN_CURRENCY", "FAIL", "Open foreign currency selector failed")
        return False

    if not select_foreign_currency_balance():
        write_log("EXPORT_FOREIGN_CURRENCY", "FAIL", "Select foreign currency balance failed")
        return False
    write_log("EXPORT_FOREIGN_CURRENCY", "OK", "Foreign currency balance selected")

    try:
        filename = export_foreign_currency_balance()
        write_export_success("EXPORT_FOREIGN_CURRENCY", filename, "Foreign currency balance export completed")
    except Exception as exc:
        write_log("EXPORT_FOREIGN_CURRENCY", "FAIL", f"Foreign currency balance export failed | {exc}")
        return False

    write_log("EXPORT_PROJECT_BALANCE", "RUN", "Open project balance selector")
    if not open_project_balance_selector():
        write_log("EXPORT_PROJECT_BALANCE", "FAIL", "Open project balance selector failed")
        return False

    if not select_project_balance():
        write_log("EXPORT_PROJECT_BALANCE", "FAIL", "Select project balance failed")
        return False
    write_log("EXPORT_PROJECT_BALANCE", "OK", "Project balance selected")

    try:
        filename = export_project_balance()
        write_export_success("EXPORT_PROJECT_BALANCE", filename, "Project balance export completed")
    except Exception as exc:
        write_log("EXPORT_PROJECT_BALANCE", "FAIL", f"Project balance export failed | {exc}")
        return False

    write_log("EXPORT_PROJECT_DETAIL", "RUN", "Open project detail selector")
    if not open_project_detail_selector():
        write_log("EXPORT_PROJECT_DETAIL", "FAIL", "Open project detail selector failed")
        return False

    if not select_project_detail():
        write_log("EXPORT_PROJECT_DETAIL", "FAIL", "Select project detail failed")
        return False
    write_log("EXPORT_PROJECT_DETAIL", "OK", "Project detail selected")

    try:
        filename = export_project_detail()
        write_export_success("EXPORT_PROJECT_DETAIL", filename, "Project detail export completed")
    except Exception as exc:
        write_log("EXPORT_PROJECT_DETAIL", "FAIL", f"Project detail export failed | {exc}")
        return False

    write_log("CLOSE_WINDOWS", "RUN", "Close data maintenance related windows")
    if not close_data_maintenance_window():
        write_log("CLOSE_WINDOWS", "FAIL", "Close data maintenance related windows failed")
        return False

    write_log("CLOSE_WINDOWS", "OK", "Data maintenance related windows closed")
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
    write_log("MAIN", "START", "Start multi-unit export flow")

    if not scan_project_tree():
        return

    round_index = 1
    completed_units = 0

    while True:
        write_log("ROUND", "RUN", f"Start round {round_index}")

        pending_row = find_next_pending_target()
        if pending_row is None:
            write_log(
                "MAIN",
                "OK",
                f"All pending units finished | completed_units={completed_units}",
            )
            return

        write_log("PROJECT_ENTRY", "RUN", f"Enter next pending unit | round={round_index}")
        target_row = enter_first_pending_project()
        if target_row is None:
            write_log(
                "PROJECT_ENTRY",
                "FAIL",
                f"Pending unit exists but project entry failed | round={round_index}",
            )
            return

        project_name = str(target_row.get(ROW_PROJECT_KEY, "")).strip()
        company_name = str(target_row.get(ROW_COMPANY_KEY, "")).strip()
        year = str(target_row.get(ROW_YEAR_KEY, "")).strip()
        level = str(target_row.get(ROW_LEVEL_KEY, "")).strip()
        write_log(
            "PROJECT_ENTRY",
            "OK",
            f"Project entered | company={company_name} | year={year} | level={level} | round={round_index}",
        )

        if not export_current_project_tables():
            return

        if not reset_project_tree_after_export(project_name):
            return

        completed_units += 1
        write_log(
            "ROUND",
            "OK",
            f"Round {round_index} completed | project={project_name} | company={company_name} | year={year}",
        )
        round_index += 1


if __name__ == "__main__":
    main()
