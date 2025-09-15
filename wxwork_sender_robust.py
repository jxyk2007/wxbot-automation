# -*- coding: utf-8 -*-
"""
企业微信发送器 - 彻底解决重启问题版本
基于完全重新检测的机制，不依赖任何缓存的窗口句柄
"""

import os
import time
import win32gui
import win32con
import win32api
import win32process
import psutil
import pyautogui
import pyperclip
import logging
from datetime import datetime
from typing import Dict, List, Optional, Any
from human_like_operations import HumanLikeOperations

logger = logging.getLogger(__name__)

class WXWorkSenderRobust:
    """企业微信发送器 - 抗重启版本"""

    def __init__(self, config: Dict[str, Any] = None):
        """初始化企业微信发送器"""
        self.config = config or {}

        # 初始化人性化操作模块
        self.human_ops = HumanLikeOperations()

        # 企业微信配置
        self.process_names = ["WXWork.exe", "wxwork.exe"]
        self.default_group = self.config.get('default_group', '蓝光统计')

        # 不缓存任何窗口信息，每次都重新检测
        self.initialized = False

    def find_wxwork_window(self) -> Optional[int]:
        """实时查找企业微信窗口 - 每次都重新检测"""
        try:
            logger.info("🔍 实时查找企业微信窗口...")

            # 1. 查找企业微信进程
            wxwork_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'memory_info']):
                try:
                    proc_name = proc.info['name']
                    if any(name.lower() in proc_name.lower() for name in self.process_names):
                        memory_mb = proc.info['memory_info'].rss / 1024 / 1024
                        wxwork_processes.append({
                            'pid': proc.pid,
                            'name': proc_name,
                            'memory_mb': memory_mb
                        })
                        logger.debug(f"  进程: {proc_name} (PID: {proc.pid}, 内存: {memory_mb:.1f}MB)")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            if not wxwork_processes:
                logger.error("❌ 未找到企业微信进程")
                return None

            # 2. 选择主进程（内存最大的）
            main_process = max(wxwork_processes, key=lambda p: p['memory_mb'])
            logger.info(f"✅ 主进程: PID {main_process['pid']} ({main_process['memory_mb']:.1f}MB)")

            # 3. 枚举该进程的所有窗口
            def enum_windows_callback(hwnd, windows_list):
                try:
                    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if window_pid == main_process['pid']:
                        class_name = win32gui.GetClassName(hwnd)
                        window_title = win32gui.GetWindowText(hwnd)
                        is_visible = win32gui.IsWindowVisible(hwnd)

                        windows_list.append({
                            'hwnd': hwnd,
                            'class': class_name,
                            'title': window_title,
                            'visible': is_visible
                        })
                except Exception:
                    pass

            windows_list = []
            win32gui.EnumWindows(enum_windows_callback, windows_list)

            logger.debug(f"  找到 {len(windows_list)} 个窗口")

            # 4. 查找主窗口 - 多重策略
            candidates = []

            # 策略1: 标准企业微信主窗口
            for w in windows_list:
                if (w['class'].lower() == 'weworkwindow' and
                    w['title'] == '企业微信' and
                    w['visible']):
                    candidates.append((w, 100))  # 最高优先级

            # 策略2: WeWorkWindow类但标题可能不同
            if not candidates:
                for w in windows_list:
                    if w['class'].lower() == 'weworkwindow' and w['visible']:
                        candidates.append((w, 80))

            # 策略3: WeWorkWindow类不管是否可见
            if not candidates:
                for w in windows_list:
                    if w['class'].lower() == 'weworkwindow':
                        candidates.append((w, 60))

            if not candidates:
                logger.error("❌ 未找到 WeWorkWindow 窗口")
                return None

            # 选择优先级最高的窗口
            best_window, priority = max(candidates, key=lambda x: x[1])
            hwnd = best_window['hwnd']

            # 5. 验证窗口有效性
            if not win32gui.IsWindow(hwnd):
                logger.error("❌ 窗口句柄无效")
                return None

            logger.info(f"✅ 找到企业微信窗口: '{best_window['title']}' (句柄: {hwnd})")
            return hwnd

        except Exception as e:
            logger.error(f"❌ 查找企业微信窗口失败: {e}")
            return None

    def activate_window(self, hwnd: int) -> bool:
        """激活企业微信窗口 - 强力版本"""
        try:
            logger.info(f"🎯 激活企业微信窗口: {hwnd}")

            # 1. 检查窗口是否有效
            if not win32gui.IsWindow(hwnd):
                logger.error("❌ 窗口句柄无效")
                return False

            # 2. 强制显示窗口
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            time.sleep(0.3)

            # 3. 恢复窗口（如果被最小化）
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)

            # 4. 多次尝试激活窗口
            for attempt in range(3):
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.5)

                    # 验证是否成功
                    foreground_hwnd = win32gui.GetForegroundWindow()
                    if foreground_hwnd == hwnd:
                        logger.info("✅ 窗口激活成功")
                        return True
                    else:
                        logger.warning(f"⚠️ 第 {attempt + 1} 次激活尝试失败")
                        if attempt < 2:
                            time.sleep(1)

                except Exception as e:
                    logger.warning(f"⚠️ 激活尝试 {attempt + 1} 失败: {e}")
                    if attempt < 2:
                        time.sleep(1)

            # 5. 最后的尝试 - 使用置顶
            try:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP,
                                     0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
                time.sleep(0.5)
                logger.info("🔧 使用置顶方式激活窗口")
                return True
            except Exception as e:
                logger.error(f"❌ 最终激活失败: {e}")
                return False

        except Exception as e:
            logger.error(f"❌ 激活窗口失败: {e}")
            return False

    def send_message(self, message: str, target_group: str = None) -> bool:
        """发送消息 - 完全重新检测版本"""
        try:
            target_group = target_group or self.default_group
            logger.info(f"📤 开始发送消息到: {target_group}")

            # 1. 实时查找企业微信窗口（不使用缓存）
            hwnd = self.find_wxwork_window()
            if not hwnd:
                logger.error("❌ 找不到企业微信窗口")
                return False

            # 2. 激活窗口
            if not self.activate_window(hwnd):
                logger.error("❌ 激活企业微信窗口失败")
                return False

            # 3. 搜索群聊
            logger.info(f"🔍 搜索群聊: {target_group}")

            # 模拟思考停顿
            self.human_ops.simulate_reading_pause()

            # 使用人性化搜索
            self.human_ops.human_search_and_enter(target_group)

            logger.info(f"✅ 已进入群聊: {target_group}")

            # 4. 重新激活窗口（确保焦点）
            if not self.activate_window(hwnd):
                logger.warning("⚠️ 重新激活窗口失败，继续尝试发送")

            # 5. 发送消息
            logger.info("📝 发送消息")

            # 随机小幅移动，模拟查看内容
            self.human_ops.random_small_move()

            # 获取窗口位置，计算输入框位置
            rect = win32gui.GetWindowRect(hwnd)
            input_x = rect[0] + (rect[2] - rect[0]) // 2
            input_y = rect[3] - 50  # 输入框通常在底部

            # 人性化点击输入框
            self.human_ops.human_click(input_x, input_y)

            # 模拟思考要发送什么内容
            self.human_ops.human_delay(0.5, 0.2)

            # 清空输入框
            self.human_ops.human_hotkey('ctrl', 'a')

            # 人性化输入消息
            self.human_ops.human_type_text(message, use_clipboard=True)

            # 检查内容，然后发送
            self.human_ops.human_delay(0.8, 0.3)
            pyautogui.press('enter')
            self.human_ops.human_delay(0.5, 0.2)

            logger.info("✅ 消息发送完成")
            return True

        except Exception as e:
            logger.error(f"❌ 发送消息失败: {e}")
            return False

    def test_connection(self) -> bool:
        """测试连接"""
        hwnd = self.find_wxwork_window()
        if hwnd:
            return self.activate_window(hwnd)
        return False

# 兼容性接口
class WXWorkSender(WXWorkSenderRobust):
    """兼容性包装类"""

    def __init__(self, config: Dict[str, Any] = None):
        super().__init__(config)
        self.is_initialized = True  # 兼容原接口

    def initialize(self) -> bool:
        """兼容原接口"""
        return self.test_connection()

    def SendMsg(self, message: str, target_group: str = None) -> bool:
        """兼容原接口"""
        return self.send_message(message, target_group)

if __name__ == "__main__":
    # 测试代码
    import sys
    import codecs

    # Windows控制台编码修复
    if sys.platform == 'win32':
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    sender = WXWorkSenderRobust()

    print("测试企业微信连接...")
    if sender.test_connection():
        print("连接成功！")

        # 测试发送消息（使用人性化格式）
        test_message = f"测试消息 - {datetime.now().strftime('%m月%d日 %H:%M')}"
        if sender.send_message(test_message, "蓝光统计"):
            print("消息发送成功！")
        else:
            print("消息发送失败！")
    else:
        print("连接失败！")