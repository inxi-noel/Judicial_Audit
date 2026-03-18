from pywinauto.keyboard import send_keys
import time


def open_foreign_currency_selector(wait_before_f3: float = 0.8, wait_after_f3: float = 1.0):
    time.sleep(wait_before_f3)

    send_keys("{F3}")
    time.sleep(wait_after_f3)

    print("已直接按下 F3，等待外币余额表选择窗口")
    return True