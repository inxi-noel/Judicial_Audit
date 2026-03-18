from pywinauto.keyboard import send_keys
import time


def open_project_detail_selector(wait_before_f3: float = 0.8, wait_after_f3: float = 1.0):
    time.sleep(wait_before_f3)

    send_keys("{F3}")
    time.sleep(wait_after_f3)

    print("已直接按下 F3，等待核算项目明细表选择窗口")
    return True