import os
import time
import pyperclip
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


def upload_directory(folder_path, wait_show=120, wait_close=180, win_title="打开", win_class="#32770"):
    # 检查目录是否存在
    if not os.path.isdir(folder_path):
        raise Exception(f"目录不存在: {folder_path}")

    # 检查目录下是否有文件
    files = [
        x for x in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, x))
    ]
    if not files:
        raise Exception("目录没有文件")

    # 等待上传窗口出现
    start = time.time()
    dlg = None
    while time.time() - start < wait_show:
        try:
            dlg = Desktop(backend="win32").window(
                class_name=win_class,
                title=win_title
            )
            if dlg.exists() and dlg.is_visible():
                dlg.set_focus()
                break
        except Exception:
            pass
        time.sleep(0.5)

    if not dlg:
        print("找不到上传窗口")
        return False

    hwnd = dlg.handle

    # 通过地址栏跳转到目标文件夹
    pyperclip.copy(folder_path)
    dlg["Toolbar3"].click() # send_keys("%d")
    time.sleep(0.2)
    send_keys("^a")
    time.sleep(0.1)
    send_keys("^v")
    time.sleep(0.2)
    send_keys("{ENTER}")

    # 等待目录加载完成
    time.sleep(2)

    # 查找文件视图控件
    file_view = None
    for cls in ["SysListView32", "DirectUIHWND", "SHELLDLL_DefView"]:
        try:
            obj = dlg.child_window(class_name=cls)
            if obj.exists():
                file_view = obj
                break
        except Exception:
            pass

    if not file_view:
        print("找不到文件列表")
        return False

    # 聚焦文件区域 + 全选文件
    file_view.click_input()
    time.sleep(0.5)
    send_keys("^a")
    time.sleep(0.5)

    # 确认上传
    send_keys("{ENTER}")

    # 等待窗口关闭
    close_start = time.time()
    while time.time() - close_start < wait_close:
        try:
            check = Desktop(backend="win32").window(handle=hwnd)
            if not check.exists():
                return True
        except Exception:
            return True
        time.sleep(1)

    print("关闭超时")
    return False