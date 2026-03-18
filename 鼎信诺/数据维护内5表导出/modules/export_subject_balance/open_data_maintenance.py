from pywinauto.keyboard import send_keys
import time

from modules._shared.window_retry import connect_uia_window, retry_window_operation
from modules.tree_handler.project_entry import MAIN_WINDOW_TITLE


APP_MENU_BAR_TEXT = "应用程序"


def open_data_maintenance():
    main = connect_uia_window(
        title=MAIN_WINDOW_TITLE,
        action_name="Connect main window",
    )
    if main is None:
        print("Main window not found")
        return False

    try:
        rect = main.rectangle()
        main.click_input(coords=(rect.width() // 2, 10))
        print("Main window activated:", repr(main.window_text()))
    except Exception as e:
        print("Activate main window failed:", e)
        return False

    time.sleep(0.8)

    def find_app_menu_bar():
        for mb in main.descendants(control_type="MenuBar"):
            try:
                if mb.window_text() == APP_MENU_BAR_TEXT:
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

    time.sleep(0.3)

    try:
        menu_items = app_menu_bar.children(control_type="MenuItem")
    except Exception as e:
        print("Read menu items failed:", e)
        return False

    print("Application menu count:", len(menu_items))

    if len(menu_items) <= 2:
        print("Menu count abnormal, cannot locate financial data menu")
        for i, item in enumerate(menu_items):
            try:
                print(i, repr(item.window_text()), item.rectangle())
            except Exception:
                print(i, "<read failed>")
        return False

    try:
        menu_items[2].click_input()
        print("Clicked financial data menu (Item 2)")
    except Exception as e:
        print("Click financial data menu failed:", e)
        return False

    time.sleep(0.5)
    send_keys("{DOWN 6}")
    time.sleep(0.2)
    send_keys("{ENTER}")
    time.sleep(0.2)
    send_keys("{ENTER}")
    time.sleep(0.5)

    print("Entered data maintenance window")
    return True