# -*- coding: utf-8 -*-
"""
人性化操作模块 - 模拟真人操作行为
避免被企业微信风控系统检测到
"""

import time
import random
import pyautogui
import pyperclip
import math
from typing import Tuple, List

class HumanLikeOperations:
    """人性化操作类"""

    def __init__(self):
        # 禁用pyautogui的failsafe，但保留人工安全检查
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0

    def human_delay(self, base_time: float = 1.0, variance: float = 0.3) -> None:
        """
        人性化延迟 - 模拟真人的不规律停顿

        Args:
            base_time: 基础延迟时间(秒)
            variance: 随机变化幅度
        """
        # 随机延迟：正态分布更符合人的行为
        delay = max(0.1, random.normalvariate(base_time, variance))
        time.sleep(delay)

    def human_move_to(self, x: int, y: int, duration: float = None) -> None:
        """
        人性化鼠标移动 - 模拟真人的曲线轨迹

        Args:
            x, y: 目标坐标
            duration: 移动持续时间，如果不指定则随机生成
        """
        if duration is None:
            # 根据距离计算合理的移动时间
            current_x, current_y = pyautogui.position()
            distance = math.sqrt((x - current_x) ** 2 + (y - current_y) ** 2)
            duration = max(0.3, min(2.0, distance / 800))  # 800像素/秒的移动速度

        # 添加随机抖动，模拟手的不稳定
        actual_x = x + random.randint(-2, 2)
        actual_y = y + random.randint(-2, 2)

        # 使用较长的移动时间，避免直线移动
        pyautogui.moveTo(actual_x, actual_y, duration=duration, tween=pyautogui.easeInOutQuad)

        # 移动后的短暂停顿
        self.human_delay(0.1, 0.05)

    def human_click(self, x: int, y: int, clicks: int = 1) -> None:
        """
        人性化点击 - 包含移动和点击

        Args:
            x, y: 点击坐标
            clicks: 点击次数
        """
        # 先移动到目标位置
        self.human_move_to(x, y)

        # 点击前的短暂停顿
        self.human_delay(0.15, 0.08)

        # 执行点击
        pyautogui.click(clicks=clicks)

        # 点击后停顿
        self.human_delay(0.2, 0.1)

    def human_hotkey(self, *keys) -> None:
        """
        人性化热键操作

        Args:
            *keys: 按键序列
        """
        # 按键前停顿
        self.human_delay(0.1, 0.05)

        # 执行热键
        pyautogui.hotkey(*keys)

        # 按键后停顿
        self.human_delay(0.3, 0.1)

    def human_type_text(self, text: str, use_clipboard: bool = True) -> None:
        """
        人性化文本输入

        Args:
            text: 要输入的文本
            use_clipboard: 是否使用剪贴板（长文本推荐）
        """
        if use_clipboard and len(text) > 20:
            # 长文本使用剪贴板，但添加人性化元素
            self.human_delay(0.2, 0.1)

            # 清空剪贴板（模拟Ctrl+A）
            pyautogui.hotkey('ctrl', 'a')
            self.human_delay(0.15, 0.05)

            # 复制文本到剪贴板
            pyperclip.copy(text)
            self.human_delay(0.1, 0.05)

            # 粘贴
            pyautogui.hotkey('ctrl', 'v')
            self.human_delay(0.5, 0.2)
        else:
            # 短文本模拟打字
            for char in text:
                pyautogui.write(char)
                # 打字速度随机化
                self.human_delay(0.05, 0.03)

    def human_search_and_enter(self, search_text: str) -> None:
        """
        人性化搜索操作：搜索 -> 等待 -> 回车

        Args:
            search_text: 搜索文本
        """
        # 使用Ctrl+F搜索
        self.human_hotkey('ctrl', 'f')

        # 等待搜索框出现
        self.human_delay(0.8, 0.2)

        # 清空搜索框
        self.human_hotkey('ctrl', 'a')
        self.human_delay(0.1, 0.05)

        # 输入搜索内容
        self.human_type_text(search_text, use_clipboard=True)

        # 等待搜索结果
        self.human_delay(1.2, 0.3)

        # 按回车选择第一个结果
        pyautogui.press('enter')
        self.human_delay(1.5, 0.5)

    def simulate_reading_pause(self) -> None:
        """模拟阅读停顿 - 让操作看起来更自然"""
        # 模拟读取内容的时间
        self.human_delay(1.0, 0.8)

    def random_small_move(self) -> None:
        """随机小幅移动鼠标 - 模拟人的无意识动作"""
        if random.random() < 0.3:  # 30%概率
            current_x, current_y = pyautogui.position()
            offset_x = random.randint(-10, 10)
            offset_y = random.randint(-10, 10)
            pyautogui.moveTo(current_x + offset_x, current_y + offset_y,
                           duration=random.uniform(0.2, 0.5))