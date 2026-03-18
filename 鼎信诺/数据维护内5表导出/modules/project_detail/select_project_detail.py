from pywinauto.keyboard import send_keys
import time

from modules._shared.window_retry import connect_win32_window, find_child_window_by_text


DATA_WINDOW_TITLE = "\u6570\u636e\u7ef4\u62a4"
SELECT_WINDOW_TITLE = "\u9009\u62e9\u8868"
ALL_BUTTON_TEXT = "\u5168\u90e8"
OK_BUTTON_TEXTS = ("\u786e\u5b9a(O&K)", "\u786e\u5b9a")
TABLE_LABEL = "核算项目明细表(基本表)"
SUCCESS_LABEL = "Project detail selection finished"


def select_project_detail():
    data_window = connect_win32_window(
        title=DATA_WINDOW_TITLE,
        action_name="Connect data maintenance window",
    )
    if data_window is None:
        print("Data maintenance window not found")
        return False

    try:
        rect = data_window.rectangle()
        data_window.click_input(coords=(max(20, rect.width() // 2), 10))
        print("Data maintenance window activated")
    except Exception as e:
        print("Activate data maintenance window failed:", e)
        return False

    time.sleep(0.8)

    select_win = find_child_window_by_text(
        data_window,
        SELECT_WINDOW_TITLE,
        action_name="Find select-table window",
    )
    if select_win is None:
        print("Select-table window not found")
        return False

    print(f"Select-table window found, handle={select_win.handle}")

    try:
        select_children = select_win.children()
    except Exception as e:
        print("Read select-table children failed:", e)
        return False

    all_btn = None
    ok_btn = None
    for c in select_children:
        try:
            txt = c.window_text()
            cls = c.class_name()
            if txt == ALL_BUTTON_TEXT and cls == "Button":
                all_btn = c
            if txt in OK_BUTTON_TEXTS and cls == "Button":
                ok_btn = c
        except Exception:
            continue

    if all_btn is None:
        print("All button not found")
        return False
    if ok_btn is None:
        print("Confirm button not found")
        return False

    try:
        all_btn.click_input()
        print("Clicked all button")
    except Exception as e:
        print("Click all button failed:", e)
        return False

    time.sleep(0.6)
    for i in range(10):
        send_keys("{DOWN}")
        print(f"Sent DOWN {i + 1}")
        time.sleep(0.15)

    time.sleep(0.3)
    try:
        send_keys("{SPACE}")
        print(f"Sent SPACE to check {TABLE_LABEL}")
    except Exception as e:
        print("Send SPACE failed:", e)
        return False

    time.sleep(0.4)
    try:
        ok_btn.click_input()
        print("Clicked confirm button")
    except Exception as e:
        print("Click confirm button failed:", e)
        return False

    time.sleep(0.8)
    try:
        send_keys("y")
        print("Sent y confirm")
    except Exception as e:
        print("Send y failed:", e)
        return False

    time.sleep(1.2)
    print(SUCCESS_LABEL)
    return True
