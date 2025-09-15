# -*- coding: utf-8 -*-
"""
测试企业微信点击操作
诊断为什么自动化脚本检测到窗口但不点击
"""

import time
import logging
import win32gui
import win32con
import pyautogui
import sys
from simple_wxwork_fix import find_wxwork_main_window

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_window_activation(hwnd):
    """测试窗口激活"""
    logger.info("🔧 测试窗口激活...")

    try:
        # 检查窗口是否存在
        if not win32gui.IsWindow(hwnd):
            logger.error("❌ 窗口句柄无效")
            return False

        # 获取窗口信息
        title = win32gui.GetWindowText(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        logger.info(f"  窗口标题: {title}")
        logger.info(f"  窗口位置: {rect}")

        # 检查窗口是否可见
        if not win32gui.IsWindowVisible(hwnd):
            logger.warning("⚠️ 窗口不可见，尝试显示...")
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            time.sleep(1)

        # 检查窗口是否最小化
        if win32gui.IsIconic(hwnd):
            logger.info("📱 窗口最小化，正在还原...")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(1)

        # 激活窗口
        logger.info("🎯 激活企业微信窗口...")
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(2)

        # 验证激活是否成功
        foreground_hwnd = win32gui.GetForegroundWindow()
        if foreground_hwnd == hwnd:
            logger.info("✅ 窗口激活成功")
            return True
        else:
            logger.warning(f"⚠️ 窗口激活可能失败，前台窗口: {win32gui.GetWindowText(foreground_hwnd)}")
            return False

    except Exception as e:
        logger.error(f"❌ 窗口激活失败: {e}")
        return False

def test_mouse_click():
    """测试鼠标点击功能"""
    logger.info("🖱️ 测试鼠标点击功能...")

    try:
        # 设置pyautogui
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5

        # 获取当前鼠标位置
        current_pos = pyautogui.position()
        logger.info(f"  当前鼠标位置: {current_pos}")

        # 测试移动鼠标
        logger.info("  测试鼠标移动...")
        pyautogui.moveTo(current_pos.x + 50, current_pos.y + 50, duration=1)
        new_pos = pyautogui.position()
        logger.info(f"  移动后位置: {new_pos}")

        # 移回原位置
        pyautogui.moveTo(current_pos.x, current_pos.y, duration=1)

        logger.info("✅ 鼠标功能正常")
        return True

    except Exception as e:
        logger.error(f"❌ 鼠标功能测试失败: {e}")
        return False

def test_search_group(hwnd, group_name="蓝光统计"):
    """测试搜索群聊功能"""
    logger.info(f"🔍 测试搜索群聊: {group_name}")

    try:
        # 激活窗口
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(2)

        # 获取窗口位置
        rect = win32gui.GetWindowRect(hwnd)
        window_width = rect[2] - rect[0]
        window_height = rect[3] - rect[1]

        logger.info(f"  窗口大小: {window_width} x {window_height}")

        # 计算搜索框可能的位置（企业微信一般在顶部）
        search_x = rect[0] + window_width // 2
        search_y = rect[1] + 50  # 假设搜索框在顶部50像素处

        logger.info(f"  尝试点击搜索框位置: ({search_x}, {search_y})")

        # 点击搜索框
        pyautogui.click(search_x, search_y)
        time.sleep(1)

        # 输入群名
        logger.info(f"  输入群名: {group_name}")
        pyautogui.typewrite(group_name)
        time.sleep(2)

        # 按回车
        logger.info("  按回车确认")
        pyautogui.press('enter')
        time.sleep(2)

        logger.info("✅ 搜索群聊操作完成")
        return True

    except Exception as e:
        logger.error(f"❌ 搜索群聊失败: {e}")
        return False

def manual_guidance_mode():
    """手动引导模式 - 帮助用户找到正确的点击位置"""
    logger.info("🎯 进入手动引导模式")

    print("\n" + "="*60)
    print("🎯 手动引导模式")
    print("请按照提示操作，我们来找到正确的点击位置")
    print("="*60)

    input("1. 请确保企业微信窗口打开并可见，然后按回车...")

    # 找到企业微信窗口
    main_window = find_wxwork_main_window()
    if not main_window:
        print("❌ 找不到企业微信窗口")
        return False

    hwnd = main_window['hwnd']

    # 激活窗口
    test_window_activation(hwnd)

    input("2. 企业微信窗口应该已经激活，请按回车继续...")

    # 获取鼠标位置指导
    print("\n📍 现在我们来找搜索框位置：")
    print("请将鼠标移动到企业微信的搜索框上，然后按Ctrl+C")
    print("（搜索框一般在企业微信窗口的顶部）")

    try:
        # 等待用户按键
        import keyboard
        print("等待您按Ctrl+C...")
        keyboard.wait('ctrl+c')

        # 获取当前鼠标位置
        search_pos = pyautogui.position()
        print(f"✅ 记录搜索框位置: {search_pos}")

        # 测试点击
        input("3. 按回车测试点击搜索框...")
        pyautogui.click(search_pos.x, search_pos.y)
        time.sleep(1)

        # 输入群名测试
        group_name = input("4. 请输入要搜索的群名（默认：蓝光统计）: ").strip() or "蓝光统计"
        pyautogui.typewrite(group_name)
        time.sleep(1)

        input("5. 按回车发送搜索...")
        pyautogui.press('enter')

        print("✅ 手动引导完成！")
        print(f"记录的搜索框位置: {search_pos}")

        return True

    except ImportError:
        print("❌ 需要安装keyboard库: pip install keyboard")
        return False
    except Exception as e:
        print(f"❌ 手动引导失败: {e}")
        return False

def main():
    """主测试函数"""
    logger.info("🚀 企业微信点击操作测试")
    logger.info("="*50)

    # 1. 查找企业微信窗口
    logger.info("步骤1: 查找企业微信窗口")
    main_window = find_wxwork_main_window()

    if not main_window:
        logger.error("❌ 找不到企业微信窗口，请确保企业微信已启动")
        return False

    hwnd = main_window['hwnd']
    logger.info(f"✅ 找到企业微信窗口，句柄: {hwnd}")

    # 2. 测试窗口激活
    logger.info("\n步骤2: 测试窗口激活")
    if not test_window_activation(hwnd):
        logger.error("❌ 窗口激活测试失败")
        return False

    # 3. 测试鼠标功能
    logger.info("\n步骤3: 测试鼠标功能")
    if not test_mouse_click():
        logger.error("❌ 鼠标功能测试失败")
        return False

    # 4. 询问用户测试模式
    print("\n" + "="*50)
    print("选择测试模式:")
    print("1. 自动测试（可能不准确）")
    print("2. 手动引导模式（推荐）")
    choice = input("请选择 (1/2): ").strip()

    if choice == "2":
        # 手动引导模式
        return manual_guidance_mode()
    else:
        # 自动测试模式
        logger.info("\n步骤4: 测试搜索群聊")
        return test_search_group(hwnd)

if __name__ == "__main__":
    try:
        success = main()
        if success:
            print("\n✅ 测试完成！现在应该能看到点击操作了")
        else:
            print("\n❌ 测试失败！需要进一步调试")
    except KeyboardInterrupt:
        print("\n⚠️ 用户中断测试")
    except Exception as e:
        print(f"\n❌ 测试出错: {e}")