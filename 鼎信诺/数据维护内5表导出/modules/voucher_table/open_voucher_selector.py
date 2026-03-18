from pywinauto.keyboard import send_keys
import time


def open_voucher_selector():
    time.sleep(0.3)

    # 直接发送 F3，不再定位“数据维护”窗口
    send_keys("{F3}")
    time.sleep(1)

    print("已直接按下 F3，等待选择表窗口")
    return True