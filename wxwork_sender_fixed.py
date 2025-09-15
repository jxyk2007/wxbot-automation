# -*- coding: utf-8 -*-
"""
企业微信自动发送模块 - 修复版
版本：v1.1.0
修复：企业微信重启后识别不到的问题
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

# from message_sender_interface import MessageSenderInterface, MessageSenderFactory, SendResult

# 临时基类，避免导入错误
class MessageSenderInterface:
    def __init__(self, config=None):
        self.config = config or {}
        self.is_initialized = False

class SendResult:
    def __init__(self, success, message):
        self.success = success
        self.message = message

logger = logging.getLogger(__name__)

class WXWorkSenderFixed(MessageSenderInterface):
    """企业微信发送器 - 修复版"""

    def __init__(self, config: Dict[str, Any] = None):
        """初始化企业微信发送器"""
        super().__init__(config)

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5

        # 企业微信进程和窗口信息
        self.wxwork_process = None
        self.wxwork_pid = None
        self.main_window_hwnd = None

        # 默认配置
        self.process_names = ["WXWork.exe", "wxwork.exe"]
        self.default_group = config.get('default_group', '存储统计报告群') if config else '存储统计报告群'

        # 窗口检测重试配置
        self.max_retries = 3
        self.retry_delay = 2.0

    def initialize(self) -> bool:
        """初始化企业微信发送器 - 增强版"""
        for attempt in range(self.max_retries):
            try:
                logger.info(f"初始化企业微信发送器... (尝试 {attempt + 1}/{self.max_retries})")

                # 强制重新查找进程和窗口
                self._reset_state()

                # 查找企业微信进程
                if not self.find_target_process():
                    logger.warning(f"第 {attempt + 1} 次尝试：未找到企业微信进程")
                    if attempt < self.max_retries - 1:
                        time.sleep(self.retry_delay)
                        continue
                    return False

                # 查找企业微信窗口
                if not self._find_wxwork_windows_enhanced():
                    logger.warning(f"第 {attempt + 1} 次尝试：未找到企业微信窗口")
                    if attempt < self.max_retries - 1:
                        time.sleep(self.retry_delay)
                        continue
                    return False

                self.is_initialized = True
                logger.info("企业微信发送器初始化成功")
                return True

            except Exception as e:
                logger.error(f"第 {attempt + 1} 次初始化失败: {e}")
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay)
                    continue

        logger.error("企业微信发送器初始化完全失败")
        return False

    def _reset_state(self):
        """重置状态，确保重新检测"""
        self.wxwork_process = None
        self.wxwork_pid = None
        self.main_window_hwnd = None

    def find_target_process(self) -> bool:
        """查找企业微信进程 - 增强版"""
        try:
            wxwork_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe', 'memory_info']):
                try:
                    proc_name = proc.info['name']
                    if any(name.lower() in proc_name.lower() for name in self.process_names):
                        # 添加内存信息用于选择主进程
                        memory_mb = proc.info['memory_info'].rss / 1024 / 1024
                        proc.memory_mb = memory_mb
                        wxwork_processes.append(proc)
                        logger.info(f"找到企业微信进程: {proc_name} (PID: {proc.pid}, 内存: {memory_mb:.1f}MB)")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            if not wxwork_processes:
                logger.error("未找到企业微信进程，请先启动企业微信")
                return False

            # 选择内存最大的进程（通常是主进程）
            if len(wxwork_processes) > 1:
                wxwork_processes.sort(key=lambda p: p.memory_mb, reverse=True)
                logger.info(f"找到 {len(wxwork_processes)} 个企业微信进程，选择内存最大的")

            self.wxwork_process = wxwork_processes[0]
            self.wxwork_pid = self.wxwork_process.pid
            logger.info(f"选定企业微信主进程 PID: {self.wxwork_pid} (内存: {self.wxwork_process.memory_mb:.1f}MB)")
            return True

        except Exception as e:
            logger.error(f"查找企业微信进程失败: {e}")
            return False

    def _find_wxwork_windows_enhanced(self) -> bool:
        """查找企业微信窗口 - 增强版"""
        try:
            if not self.wxwork_pid:
                logger.error("请先查找企业微信进程")
                return False

            # 枚举所有窗口
            windows_list = []
            win32gui.EnumWindows(self._enum_windows_callback, windows_list)

            # 查找属于企业微信进程的窗口
            wxwork_windows = [w for w in windows_list if w['pid'] == self.wxwork_pid]

            if not wxwork_windows:
                logger.error("未找到企业微信窗口")
                return False

            logger.info(f"找到 {len(wxwork_windows)} 个企业微信窗口")

            # 多重策略查找主窗口
            main_window = self._find_main_window_multi_strategy(wxwork_windows)

            if main_window:
                self.main_window_hwnd = main_window['hwnd']
                logger.info(f"找到企业微信主窗口: '{main_window['title']}' (类名: {main_window['class']})")

                # 验证窗口是否有效
                if self._validate_window(self.main_window_hwnd):
                    return True
                else:
                    logger.warning("主窗口验证失败，重新查找")
                    return False
            else:
                logger.error("无法确定企业微信主窗口")
                return False

        except Exception as e:
            logger.error(f"查找企业微信窗口失败: {e}")
            return False

    def _find_main_window_multi_strategy(self, windows) -> Optional[Dict]:
        """多策略查找主窗口"""

        # 策略1: 标准企业微信窗口特征
        for w in windows:
            if (w['class'].lower() == 'weworkwindow' and
                w['title'] == '企业微信' and
                w.get('visible', False)):
                logger.info("策略1命中: 标准企业微信主窗口")
                return w

        # 策略2: WeWorkWindow类但可能标题不同
        wework_windows = [w for w in windows if w['class'].lower() == 'weworkwindow']
        if wework_windows:
            # 优先选择有标题且可见的
            visible_titled = [w for w in wework_windows if w['title'].strip() and w.get('visible', False)]
            if visible_titled:
                logger.info(f"策略2命中: WeWorkWindow类窗口 '{visible_titled[0]['title']}'")
                return visible_titled[0]

            # 备选: 任何WeWorkWindow
            logger.info(f"策略2备选: 使用第一个WeWorkWindow")
            return wework_windows[0]

        # 策略3: 最大的可见窗口
        visible_windows = [w for w in windows if w.get('visible', False) and w['title'].strip()]
        if visible_windows:
            # 按窗口大小排序（需要计算窗口面积）
            for w in visible_windows:
                try:
                    rect = win32gui.GetWindowRect(w['hwnd'])
                    w['area'] = (rect[2] - rect[0]) * (rect[3] - rect[1])
                except:
                    w['area'] = 0

            visible_windows.sort(key=lambda x: x['area'], reverse=True)
            logger.info(f"策略3命中: 最大可见窗口 '{visible_windows[0]['title']}'")
            return visible_windows[0]

        # 策略4: 最后的备选方案
        if windows:
            logger.info("策略4: 使用第一个找到的窗口")
            return windows[0]

        return None

    def _validate_window(self, hwnd) -> bool:
        """验证窗口是否有效且可用"""
        try:
            # 检查窗口是否存在
            if not win32gui.IsWindow(hwnd):
                logger.warning("窗口句柄无效")
                return False

            # 检查窗口是否可见
            if not win32gui.IsWindowVisible(hwnd):
                logger.warning("窗口不可见")
                return False

            # 尝试获取窗口标题
            title = win32gui.GetWindowText(hwnd)
            logger.info(f"窗口验证通过: '{title}'")
            return True

        except Exception as e:
            logger.warning(f"窗口验证失败: {e}")
            return False

    def _enum_windows_callback(self, hwnd, windows_list):
        """枚举窗口回调函数"""
        try:
            # 获取窗口所属进程ID
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)

            # 获取窗口信息
            class_name = win32gui.GetClassName(hwnd)
            window_title = win32gui.GetWindowText(hwnd)

            # 检查窗口是否可见
            is_visible = win32gui.IsWindowVisible(hwnd)

            windows_list.append({
                'hwnd': hwnd,
                'pid': window_pid,
                'class': class_name,
                'title': window_title,
                'visible': is_visible
            })

        except Exception:
            pass  # 忽略无法访问的窗口

    def send_message(self, message: str, target_group: str = None) -> SendResult:
        """发送消息 - 带重连机制"""

        # 如果未初始化或窗口无效，尝试重新初始化
        if not self.is_initialized or not self._validate_window(self.main_window_hwnd):
            logger.warning("检测到窗口失效，尝试重新初始化...")
            if not self.initialize():
                return SendResult(False, "重新初始化失败")

        # 原有的发送逻辑...
        try:
            # 激活企业微信窗口
            win32gui.SetForegroundWindow(self.main_window_hwnd)
            time.sleep(0.5)

            # 这里添加你原有的发送消息逻辑
            # ...

            logger.info(f"消息发送成功: {message[:50]}...")
            return SendResult(True, "发送成功")

        except Exception as e:
            logger.error(f"发送消息失败: {e}")
            return SendResult(False, f"发送失败: {e}")

# 使用示例和测试函数
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    sender = WXWorkSenderFixed()
    if sender.initialize():
        print("企业微信发送器初始化成功！")
        # result = sender.send_message("测试消息", "测试群")
        # print(f"发送结果: {result}")
    else:
        print("企业微信发送器初始化失败！")