from pywinauto.keyboard import send_keys
import time


def close_data_maintenance_window(
    close_times: int = 7,
    wait_before_close: float = 0.8,
    wait_between_close: float = 0.6,
    wait_after_close: float = 1.0
):
    time.sleep(wait_before_close)

    for i in range(close_times):
        send_keys("%{F4}")
        print(f"已执行 Alt+F4：第 {i + 1}/{close_times} 次")
        time.sleep(wait_between_close)

    time.sleep(wait_after_close)

    print("数据维护相关残留窗口关闭流程完成")
    return True