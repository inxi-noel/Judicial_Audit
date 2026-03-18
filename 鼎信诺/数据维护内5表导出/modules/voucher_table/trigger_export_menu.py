from pywinauto.keyboard import send_keys
import time
import re
import pyperclip

from modules._shared.window_retry import connect_uia_window


WINDOW_TITLE_RE = ".*凭证表.*"
WINDOW_CLASS_NAME = "FNWND3105"
WINDOW_LABEL = "voucher"
DATA_MAINTENANCE_PREFIX_RE = r"^\s*\u6570\u636e\u7ef4\u62a4\s*-\s*"


def sanitize_filename(filename: str) -> str:
    if not filename:
        raise ValueError("Filename is empty")

    filename = filename.strip()
    filename = re.sub(r'[\/:*?"<>|]', "_", filename)
    filename = filename.strip(" .")

    if not filename:
        raise ValueError("Filename is empty after sanitize")

    return filename


def get_export_filename_from_window_title(window_text: str) -> str:
    if not window_text:
        raise ValueError("Window title is empty")

    window_text = window_text.strip()
    m = re.search(r"\[(.*?)\]", window_text)
    if m:
        filename = m.group(1).strip()
    else:
        filename = re.sub(DATA_MAINTENANCE_PREFIX_RE, "", window_text).strip()

    if not filename:
        raise ValueError(f"Cannot parse export filename from title: {window_text}")

    return sanitize_filename(filename)


def find_voucher_window():
    data_window = connect_uia_window(
        title_re=WINDOW_TITLE_RE,
        class_name=WINDOW_CLASS_NAME,
        action_name=f"Find {WINDOW_LABEL} window",
    )
    if data_window is None:
        raise RuntimeError(f"{WINDOW_LABEL} window not found")
    return data_window


def activate_window_center(window):
    rect = window.rectangle()
    rel_x = rect.width() // 2
    rel_y = rect.height() // 2
    window.click_input(coords=(rel_x, rel_y))


def trigger_export_menu():
    data_window = find_voucher_window()

    full_title = data_window.window_text()
    export_filename = get_export_filename_from_window_title(full_title)

    print("Found window:", full_title)
    print("Export filename:", export_filename)

    activate_window_center(data_window)
    time.sleep(0.5)
    send_keys("%e")
    time.sleep(0.3)
    send_keys("l")
    time.sleep(0.3)

    print("Triggered Edit -> Export Excel")
    return export_filename


def wait_and_save_by_hotkeys(export_filename: str, wait_seconds: float = 1.5):
    if not export_filename:
        raise ValueError("export_filename is empty")

    export_filename = sanitize_filename(export_filename)

    time.sleep(wait_seconds)
    send_keys("^a")
    time.sleep(0.1)
    send_keys("{BACKSPACE}")
    time.sleep(0.1)
    pyperclip.copy(export_filename)
    time.sleep(0.1)
    send_keys("^v")
    time.sleep(0.3)
    send_keys("%s")
    time.sleep(0.8)
    send_keys("{ENTER}")
    time.sleep(0.5)

    print("Entered filename, saved, and confirmed")


def export_voucher_table():
    export_filename = trigger_export_menu()
    wait_and_save_by_hotkeys(export_filename)
    return export_filename


if __name__ == "__main__":
    export_filename = export_voucher_table()
    print("Export filename:", export_filename)
