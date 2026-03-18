import time

from modules._shared.window_retry import connect_uia_window


DATA_WINDOW_TITLE = "数据维护"


def select_table_by_index(target_index: int, expected_text: str):
    data_window = connect_uia_window(
        title=DATA_WINDOW_TITLE,
        action_name="Connect data maintenance window",
    )
    if data_window is None:
        print("Data maintenance window not found")
        return False

    edits = data_window.descendants(control_type="Edit")
    if target_index >= len(edits):
        print(f"Edit index out of range: target_index={target_index}, actual={len(edits)}")
        return False

    target_edit = edits[target_index]

    try:
        actual_text = target_edit.get_value()
    except Exception as e:
        print("Read target edit text failed:", e)
        return False

    print(f"Target edit[{target_index}] text: {repr(actual_text)}")

    if actual_text != expected_text:
        print(f"Target edit text mismatch: expected={repr(expected_text)}, actual={repr(actual_text)}")
        return False

    rect = target_edit.rectangle()
    x = rect.left - 25
    y = rect.top + rect.height() // 2

    try:
        data_window.click_input(coords=(x, y))
        print("Checked target table:", actual_text)
    except Exception as e:
        print("Check target table failed:", e)
        return False

    try:
        data_window.child_window(auto_id="1004", control_type="Button").click_input()
    except Exception as e:
        print("Click confirm button failed:", e)
        return False

    time.sleep(1)

    buttons = data_window.descendants(control_type="Button")
    for b in buttons:
        try:
            txt = b.window_text()
            if "是" in txt:
                b.invoke()
                print("Selection confirm finished")
                return True
        except Exception:
            continue

    print("Confirm button not found, continue by default")
    return True
