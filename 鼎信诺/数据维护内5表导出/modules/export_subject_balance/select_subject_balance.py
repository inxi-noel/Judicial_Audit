from modules._shared.select_table_by_index import select_table_by_index


def select_subject_balance():
    return select_table_by_index(
        target_index=0,
        expected_text="科目余额表(基本表)"
    )