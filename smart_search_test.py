# -*- coding: utf-8 -*-
"""
智能搜索测试 - 找到正确的企业微信搜索方式
"""

import time
import win32gui
import pyautogui
import pyperclip
import logging
from wxwork_sender import WXWorkSenderRobust

def test_search_methods():
    """测试不同的搜索方法"""

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

    sender = WXWorkSenderRobust()

    # 查找企业微信窗口
    hwnd = sender.find_wxwork_window()
    if not hwnd:
        print("❌ 找不到企业微信窗口")
        return

    # 激活窗口
    if not sender.activate_window(hwnd):
        print("❌ 激活企业微信窗口失败")
        return

    target_group = "蓝光统计"

    print("\n🧪 开始测试不同的搜索方法")
    print("请观察企业微信窗口的变化，选择有效的方法")

    # 方法1: Ctrl+F
    print("\n方法1: 使用 Ctrl+F")
    input("按回车开始测试 Ctrl+F...")

    sender.activate_window(hwnd)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(1)
    pyperclip.copy(target_group)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    effective = input("这个方法有效吗？(y/n): ").lower() == 'y'
    if effective:
        print("✅ Ctrl+F 方法有效")
        return "ctrl+f"

    # 方法2: 点击搜索框（顶部中央）
    print("\n方法2: 点击搜索框（顶部中央）")
    input("按回车开始测试点击搜索框...")

    sender.activate_window(hwnd)
    rect = win32gui.GetWindowRect(hwnd)
    search_x = rect[0] + (rect[2] - rect[0]) // 2
    search_y = rect[1] + 50

    print(f"点击位置: ({search_x}, {search_y})")
    pyautogui.click(search_x, search_y)
    time.sleep(1)
    pyperclip.copy(target_group)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    effective = input("这个方法有效吗？(y/n): ").lower() == 'y'
    if effective:
        print("✅ 点击搜索框方法有效")
        return "click_search"

    # 方法3: 点击搜索框（左上角）
    print("\n方法3: 点击搜索框（左上角）")
    input("按回车开始测试点击搜索框（左上角）...")

    sender.activate_window(hwnd)
    search_x = rect[0] + 100
    search_y = rect[1] + 50

    print(f"点击位置: ({search_x}, {search_y})")
    pyautogui.click(search_x, search_y)
    time.sleep(1)
    pyperclip.copy(target_group)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    effective = input("这个方法有效吗？(y/n): ").lower() == 'y'
    if effective:
        print("✅ 点击搜索框（左上角）方法有效")
        return "click_search_left"

    # 方法4: 手动指导
    print("\n方法4: 手动指导模式")
    print("请手动点击企业微信的搜索框，然后按回车...")
    input("点击搜索框后按回车...")

    pyperclip.copy(target_group)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)

    effective = input("这个方法有效吗？(y/n): ").lower() == 'y'
    if effective:
        # 获取当前鼠标位置作为搜索框位置
        pos = pyautogui.position()
        print(f"✅ 手动方法有效，搜索框位置: {pos}")
        return f"manual_position_{pos.x}_{pos.y}"

    print("❌ 所有方法都无效，需要进一步调试")
    return None

if __name__ == "__main__":
    result = test_search_methods()
    if result:
        print(f"\n✅ 找到有效的搜索方法: {result}")
        print("现在可以更新主程序使用这个方法")
    else:
        print("\n❌ 未找到有效的搜索方法")
        print("可能需要使用其他策略，如OCR识别或图像匹配")