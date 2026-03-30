# -*- coding: utf-8 -*-
"""
个人微信自动发送模块 v4

设计目标:
1. 原有 v3 文件不改，新建 v4 实现。
2. 全程使用鼠标键盘模拟执行，降低自动化痕迹。
3. 使用相对窗口标定，适配窗口大小变化。
4. 使用 PaddleOCR 做主定位，可选使用本地 LM Studio 视觉模型做复核。
5. 发送前后都做保守校验，宁可失败也不误发。
"""

import base64
import io
import json
import logging
import os
import re
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import psutil
import pyautogui
import pyperclip
import requests
import win32con
import win32gui
import win32process
from PIL import ImageDraw

from human_like_operations import HumanLikeOperations
from message_sender_interface import MessageSenderFactory, MessageSenderInterface

logger = logging.getLogger(__name__)

try:
    import numpy as np
except ImportError:  # pragma: no cover - optional dependency
    np = None

try:
    from paddleocr import PaddleOCR
except ImportError:  # pragma: no cover - optional dependency
    PaddleOCR = None


@dataclass
class OCRMatch:
    text: str
    score: float
    box: List[Tuple[float, float]]


class PaddleOCRRecognizer:
    """对固定区域截图执行 OCR。"""

    def __init__(self, use_angle_cls: bool = False, lang: str = "ch"):
        self.available = PaddleOCR is not None and np is not None
        self._ocr = None
        if self.available:
            try:
                self._ocr = PaddleOCR(use_angle_cls=use_angle_cls, lang=lang, show_log=False)
            except Exception as exc:
                logger.error("初始化 PaddleOCR 失败: %s", exc)
                self.available = False

    def recognize(self, image) -> List[OCRMatch]:
        if not self.available or self._ocr is None:
            raise RuntimeError("PaddleOCR 不可用，请先安装 paddleocr 和 numpy")

        rgb = image.convert("RGB")
        result = self._ocr.ocr(np.array(rgb), cls=False)
        matches: List[OCRMatch] = []

        if not result:
            return matches

        lines = result[0] if isinstance(result[0], list) else result
        for item in lines:
            if not item or len(item) < 2:
                continue
            box = item[0]
            text, score = item[1]
            matches.append(
                OCRMatch(
                    text=str(text).strip(),
                    score=float(score),
                    box=[(float(x), float(y)) for x, y in box],
                )
            )
        return matches


class LocalVLMVerifier:
    """兼容 LM Studio OpenAI 风格接口的视觉复核器。"""

    def __init__(self, config: Dict[str, Any]):
        self.enabled = bool(config.get("enabled", False))
        self.api_url = config.get("api_url", "http://127.0.0.1:1234/v1/chat/completions")
        self.model = config.get("model", "qwen3.5-4b")
        self.api_key = config.get("api_key", "")
        self.timeout = int(config.get("timeout", 20))

    def verify(self, image, prompt: str) -> bool:
        if not self.enabled:
            return True

        try:
            buffer = io.BytesIO()
            image.save(buffer, format="PNG")
            image_b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")

            headers = {"Content-Type": "application/json"}
            if self.api_key:
                headers["Authorization"] = f"Bearer {self.api_key}"

            payload = {
                "model": self.model,
                "temperature": 0.0,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{image_b64}"},
                            },
                        ],
                    }
                ],
            }
            response = requests.post(
                self.api_url,
                headers=headers,
                json=payload,
                timeout=self.timeout,
            )
            response.raise_for_status()
            data = response.json()
            content = data["choices"][0]["message"]["content"]
            if isinstance(content, list):
                content = " ".join(
                    part.get("text", "") for part in content if isinstance(part, dict)
                )
            answer = str(content).strip().lower()
            logger.info("VLM 复核结果: %s", answer)
            return answer.startswith("yes") or answer.startswith("是")
        except Exception as exc:
            logger.error("VLM 复核失败: %s", exc)
            return False


class WeChatSenderV4(MessageSenderInterface):
    """个人微信发送器 v4，高稳 OCR + VLM 复核版。"""

    def __init__(self, config: Dict[str, Any] = None):
        super().__init__(config)
        self.config = config or {}
        self.process_names = self.config.get(
            "process_names", ["WeChat.exe", "Weixin.exe", "wechat.exe"]
        )
        self.default_group = self.config.get("default_group", "存储统计报告群")
        self.config_path = self.config.get(
            "config_path",
            os.path.join(os.path.dirname(__file__), "wechat_sender_v4_config.json"),
        )
        self.force_foreground = bool(self.config.get("force_foreground", True))
        self.temporary_topmost = bool(self.config.get("temporary_topmost", True))
        self.verify_foreground_before_each_step = bool(
            self.config.get("verify_foreground_before_each_step", True)
        )
        self.ocr_threshold = float(self.config.get("ocr_threshold", 0.88))
        self.chat_title_threshold = float(self.config.get("chat_title_threshold", 0.80))
        self.result_refresh_delay = float(self.config.get("result_refresh_delay", 1.6))
        self.post_click_delay = float(self.config.get("post_click_delay", 1.5))
        self.post_send_delay = float(self.config.get("post_send_delay", 1.0))
        self.search_keyword_threshold = float(self.config.get("search_keyword_threshold", 0.70))
        self.calibration: Dict[str, Any] = {}

        self.wechat_process = None
        self.wechat_pid = None
        self.main_window_hwnd = None
        self._topmost_enabled = False

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1

        self.human = HumanLikeOperations()
        self.ocr = PaddleOCRRecognizer(
            use_angle_cls=bool(self.config.get("ocr_use_angle_cls", False)),
            lang=self.config.get("ocr_lang", "ch"),
        )
        self.vlm = LocalVLMVerifier(
            self.config.get(
                "vlm",
                {
                    "enabled": False,
                    "api_url": "http://127.0.0.1:1234/v1/chat/completions",
                    "model": "qwen-vl",
                },
            )
        )

    def initialize(self) -> bool:
        try:
            if not self.find_target_process():
                return False
            if not self._find_wechat_windows():
                return False
            if not self.ocr.available:
                logger.error("PaddleOCR 不可用，无法执行 v4 高稳搜索链路")
                return False
            self._load_calibration()
            self.is_initialized = True
            return True
        except Exception as exc:
            logger.error("初始化微信发送器 v4 失败: %s", exc)
            return False

    def _load_calibration(self) -> None:
        if os.path.exists(self.config_path):
            with open(self.config_path, "r", encoding="utf-8") as file:
                self.calibration = json.load(file)
        else:
            self.calibration = {}

    def _save_calibration(self) -> None:
        with open(self.config_path, "w", encoding="utf-8") as file:
            json.dump(self.calibration, file, ensure_ascii=False, indent=2)

    def _get_search_tuning(self) -> Dict[str, float]:
        tuning = self.calibration.get("search_tuning", {})
        defaults = {
            "left_expand_ratio": 1.8,
            "right_expand_ratio": 3.0,
            "top_expand_ratio": 0.9,
            "bottom_expand_ratio": 0.9,
            "click_x_ratio": 0.28,
            "click_y_ratio": 0.5,
        }
        result = defaults.copy()
        for key in defaults:
            try:
                if key in tuning:
                    result[key] = float(tuning[key])
            except (TypeError, ValueError):
                logger.warning("search_tuning.%s 非法，回退默认值", key)
        return result

    def set_search_tuning(self, **kwargs: float) -> Dict[str, float]:
        tuning = self._get_search_tuning()
        allowed = set(tuning.keys())
        for key, value in kwargs.items():
            if key not in allowed or value is None:
                continue
            tuning[key] = float(value)
        self.calibration["search_tuning"] = tuning
        self._save_calibration()
        return tuning

    def _get_result_tuning(self) -> Dict[str, float]:
        tuning = self.calibration.get("result_tuning", {})
        defaults = {
            "left_expand_ratio": 1.8,
            "left_expand_min": 90.0,
            "top_expand_ratio": 1.1,
            "top_expand_min": 18.0,
            "bottom_expand_ratio": 1.1,
            "bottom_expand_min": 18.0,
            "click_x_ratio": 0.22,
            "click_y_ratio": 0.5,
        }
        result = defaults.copy()
        for key in defaults:
            try:
                if key in tuning:
                    result[key] = float(tuning[key])
            except (TypeError, ValueError):
                logger.warning("result_tuning.%s 非法，回退默认值", key)
        return result

    def set_result_tuning(self, **kwargs: float) -> Dict[str, float]:
        tuning = self._get_result_tuning()
        allowed = set(tuning.keys())
        for key, value in kwargs.items():
            if key not in allowed or value is None:
                continue
            tuning[key] = float(value)
        self.calibration["result_tuning"] = tuning
        self._save_calibration()
        return tuning

    def find_target_process(self) -> bool:
        try:
            candidates = []
            for proc in psutil.process_iter(["pid", "name", "exe"]):
                try:
                    proc_name = str(proc.info.get("name", ""))
                    if any(name.lower() == proc_name.lower() for name in self.process_names):
                        candidates.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            if not candidates:
                logger.error("未找到微信进程，请先启动微信")
                return False

            self.wechat_process = candidates[0]
            self.wechat_pid = self.wechat_process.pid
            logger.info("找到微信进程 PID: %s", self.wechat_pid)
            return True
        except Exception as exc:
            logger.error("查找微信进程失败: %s", exc)
            return False

    def _enum_windows_callback(self, hwnd, windows_list):
        if not win32gui.IsWindowVisible(hwnd):
            return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        windows_list.append(
            {
                "hwnd": hwnd,
                "title": win32gui.GetWindowText(hwnd),
                "class": win32gui.GetClassName(hwnd),
                "pid": pid,
            }
        )

    def _is_wechat_process_name(self, process_name: str) -> bool:
        return any(name.lower() == str(process_name).lower() for name in self.process_names)

    def _get_candidate_window_list(self) -> List[Dict[str, Any]]:
        windows_list: List[Dict[str, Any]] = []
        win32gui.EnumWindows(self._enum_windows_callback, windows_list)

        candidate_pids = set()
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                if self._is_wechat_process_name(proc.info.get("name", "")):
                    candidate_pids.add(proc.info["pid"])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        return [item for item in windows_list if item["pid"] in candidate_pids]

    def _find_wechat_windows(self) -> bool:
        if not self.wechat_pid and not self.find_target_process():
            return False

        candidate_windows = self._get_candidate_window_list()
        wechat_windows = [item for item in candidate_windows if item["pid"] == self.wechat_pid]
        if not wechat_windows:
            wechat_windows = candidate_windows
        if not wechat_windows:
            logger.error("未找到微信主窗口")
            return False

        preferred = [
            item for item in wechat_windows if "WeChatMainWndForPC" in item["class"]
        ]
        titled = [item for item in wechat_windows if str(item["title"]).strip()]
        selected_pool = preferred or titled or wechat_windows
        selected = max(selected_pool, key=lambda item: len(str(item["title"])))
        self.main_window_hwnd = selected["hwnd"]
        if selected["pid"] != self.wechat_pid:
            self.wechat_pid = selected["pid"]
            try:
                self.wechat_process = psutil.Process(self.wechat_pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                self.wechat_process = None
        logger.info("使用微信窗口: %s / %s", selected["class"], selected["title"])
        return True

    def _get_window_rect(self) -> Tuple[int, int, int, int]:
        if not self.main_window_hwnd:
            raise RuntimeError("微信主窗口未初始化")
        return win32gui.GetWindowRect(self.main_window_hwnd)

    def _window_size(self) -> Tuple[int, int]:
        left, top, right, bottom = self._get_window_rect()
        return right - left, bottom - top

    def _get_foreground_hwnd(self) -> int:
        return win32gui.GetForegroundWindow()

    def _is_wechat_foreground(self) -> bool:
        hwnd = self._get_foreground_hwnd()
        if not hwnd:
            return False
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid == self.wechat_pid

    def _ensure_wechat_foreground(self) -> bool:
        if not self.verify_foreground_before_each_step:
            return True
        if self._is_wechat_foreground():
            return True
        logger.error("微信不在前台，终止后续操作")
        return False

    def activate_application(self) -> bool:
        try:
            if not self.main_window_hwnd:
                return False
            if win32gui.IsIconic(self.main_window_hwnd):
                win32gui.ShowWindow(self.main_window_hwnd, win32con.SW_RESTORE)
            if self.force_foreground:
                win32gui.SetForegroundWindow(self.main_window_hwnd)
            time.sleep(0.8)
            return self._is_wechat_foreground() if self.verify_foreground_before_each_step else True
        except Exception as exc:
            logger.error("激活微信窗口失败: %s", exc)
            return False

    def _set_temporary_topmost(self, enabled: bool) -> None:
        if not self.main_window_hwnd or not self.temporary_topmost:
            return

        insert_after = win32con.HWND_TOPMOST if enabled else win32con.HWND_NOTOPMOST
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(self.main_window_hwnd, insert_after, 0, 0, 0, 0, flags)
        self._topmost_enabled = enabled

    def _normalized_to_screen_point(self, anchor: Dict[str, float]) -> Tuple[int, int]:
        left, top, right, bottom = self._get_window_rect()
        width = right - left
        height = bottom - top
        return (
            int(left + width * float(anchor["x"])),
            int(top + height * float(anchor["y"])),
        )

    def _normalized_to_screen_region(self, region: Dict[str, float]) -> Tuple[int, int, int, int]:
        left, top, right, bottom = self._get_window_rect()
        width = right - left
        height = bottom - top
        x = int(left + width * float(region["x"]))
        y = int(top + height * float(region["y"]))
        w = max(1, int(width * float(region["width"])))
        h = max(1, int(height * float(region["height"])))
        return x, y, w, h

    def _to_normalized_point(self, point: Tuple[int, int]) -> Dict[str, float]:
        left, top, right, bottom = self._get_window_rect()
        width = max(1, right - left)
        height = max(1, bottom - top)
        return {
            "x": round((point[0] - left) / width, 6),
            "y": round((point[1] - top) / height, 6),
        }

    def _to_normalized_region(self, start: Tuple[int, int], end: Tuple[int, int]) -> Dict[str, float]:
        left, top, right, bottom = self._get_window_rect()
        width = max(1, right - left)
        height = max(1, bottom - top)
        x1, x2 = sorted([start[0], end[0]])
        y1, y2 = sorted([start[1], end[1]])
        return {
            "x": round((x1 - left) / width, 6),
            "y": round((y1 - top) / height, 6),
            "width": round((x2 - x1) / width, 6),
            "height": round((y2 - y1) / height, 6),
        }

    def _capture_region(self, region: Dict[str, float]):
        actual = self._normalized_to_screen_region(region)
        return pyautogui.screenshot(region=actual), actual

    def _default_region(self, name: str) -> Dict[str, float]:
        defaults = {
            "sidebar_region": {"x": 0.0, "y": 0.0, "width": 0.36, "height": 0.30},
            "search_results_region": {"x": 0.0, "y": 0.10, "width": 0.36, "height": 0.70},
            "chat_title_region": {"x": 0.34, "y": 0.0, "width": 0.46, "height": 0.10},
            "send_button_region": {"x": 0.72, "y": 0.84, "width": 0.22, "height": 0.12},
        }
        if name not in defaults:
            raise KeyError(f"未定义默认区域: {name}")
        return defaults[name]

    def _default_anchor(self, name: str) -> Dict[str, float]:
        defaults = {
            "chat_input_anchor": {"x": 0.60, "y": 0.91},
            "send_button_anchor": {"x": 0.91, "y": 0.91},
        }
        if name not in defaults:
            raise KeyError(f"未定义默认锚点: {name}")
        return defaults[name]

    def _is_valid_region(self, region: Dict[str, float]) -> bool:
        try:
            x = float(region["x"])
            y = float(region["y"])
            width = float(region["width"])
            height = float(region["height"])
        except (KeyError, TypeError, ValueError):
            return False

        return (
            0.0 <= x < 1.0
            and 0.0 <= y < 1.0
            and 0.02 <= width <= 1.0
            and 0.02 <= height <= 1.0
            and x + width <= 1.02
            and y + height <= 1.02
        )

    def _is_valid_anchor(self, anchor: Dict[str, float]) -> bool:
        try:
            x = float(anchor["x"])
            y = float(anchor["y"])
        except (KeyError, TypeError, ValueError):
            return False
        return 0.0 <= x <= 1.0 and 0.0 <= y <= 1.0

    def _get_region_config(self, name: str) -> Dict[str, float]:
        region = self.calibration.get(name)
        if region and self._is_valid_region(region):
            return region
        if region:
            logger.warning("标定区域 %s 非法，回退默认区域: %s", name, region)
        return self._default_region(name)

    def _get_anchor_config(self, name: str) -> Dict[str, float]:
        anchor = self.calibration.get(name)
        if anchor and self._is_valid_anchor(anchor):
            return anchor
        if anchor:
            logger.warning("标定锚点 %s 非法，回退默认锚点: %s", name, anchor)
        return self._default_anchor(name)

    def _prompt_point(self, label: str) -> Dict[str, float]:
        input(f"{label}：请把鼠标移动到目标位置后按回车...")
        point = pyautogui.position()
        logger.info("%s 记录位置: %s", label, point)
        return self._to_normalized_point((point.x, point.y))

    def _prompt_region(self, label: str) -> Dict[str, float]:
        input(f"{label}：请把鼠标移动到区域左上角后按回车...")
        start = pyautogui.position()
        input(f"{label}：请把鼠标移动到区域右下角后按回车...")
        end = pyautogui.position()
        region = self._to_normalized_region((start.x, start.y), (end.x, end.y))
        logger.info("%s 记录区域: %s", label, region)
        return region

    def calibrate(self, force: bool = False) -> bool:
        try:
            if not self.activate_application():
                return False

            if self.calibration and not force:
                logger.info("已存在标定文件，跳过。使用 force=True 可重新标定")
                return True

            print("\n=== 微信发送器 v4 标定模式 ===")
            print("请保持微信窗口可见，后续窗口大小可以变化，但布局应保持一致。")

            self.calibration = {
                "window_meta": {
                    "class_name": win32gui.GetClassName(self.main_window_hwnd),
                    "title": win32gui.GetWindowText(self.main_window_hwnd),
                    "baseline_window_size": {
                        "width": self._window_size()[0],
                        "height": self._window_size()[1],
                    },
                },
                "sidebar_region": self._prompt_region("左侧栏区域（包含搜索框和搜索结果）"),
                "search_results_region": self._prompt_region("搜索结果区域"),
                "chat_title_region": self._prompt_region("聊天标题区域"),
                "chat_input_anchor": self._prompt_point("输入框中心点"),
                "send_button_region": self._prompt_region("发送按钮粗区域"),
            }
            self._save_calibration()
            logger.info("标定完成，配置已保存到 %s", self.config_path)
            return True
        except Exception as exc:
            logger.error("标定失败: %s", exc)
            return False

    def _require_calibration(self) -> bool:
        if self.calibration:
            return True
        self._load_calibration()
        if self.calibration:
            return True
        logger.warning("未找到标定配置，回退到默认区域和默认锚点")
        return True

    def _click_anchor(self, key: str) -> bool:
        if not self._ensure_wechat_foreground():
            return False
        point = self._normalized_to_screen_point(self._get_anchor_config(key))
        self.human.human_click(point[0], point[1])
        return True

    def _click_search_box(self) -> bool:
        if not self._ensure_wechat_foreground():
            return False
        search_point = self._locate_search_box()
        if not search_point:
            logger.error("无法动态定位搜索框")
            return False
        self.human.human_click(search_point[0], search_point[1])
        self.human.human_delay(0.2, 0.05)
        return True

    def _clear_search_box(self) -> bool:
        if not self._click_search_box():
            return False
        self.human.human_delay(0.2, 0.05)
        self.human.human_hotkey("ctrl", "a")
        self.human.human_delay(0.1, 0.05)
        pyautogui.press("backspace")
        self.human.human_delay(0.2, 0.05)
        return True

    def _normalize_text(self, value: str) -> str:
        return "".join(str(value).strip().lower().split())

    def _normalize_chat_title_text(self, value: str) -> str:
        normalized = self._normalize_text(value)
        return re.sub(r"[\(\uff08]\d+[\)\uff09]$", "", normalized)

    def _should_use_clipboard_for_search(self, text: str) -> bool:
        # 中文、特殊字符或长文本直接走剪贴板，更稳。
        return len(text) > 8 or any(ord(char) > 127 for char in text)

    def _get_box_bounds(self, box: List[Tuple[float, float]]) -> Tuple[int, int, int, int]:
        xs = [int(point[0]) for point in box]
        ys = [int(point[1]) for point in box]
        return min(xs), min(ys), max(xs), max(ys)

    def _clamp(self, value: int, lower: int, upper: int) -> int:
        return max(lower, min(value, upper))

    def _calculate_search_box_geometry(
        self, image, actual_region: Tuple[int, int, int, int], matches: List[OCRMatch]
    ) -> Optional[Dict[str, Any]]:
        keyword_matches = []
        for item in matches:
            normalized = self._normalize_text(item.text)
            if "搜索" in normalized and item.score >= self.search_keyword_threshold:
                keyword_matches.append(item)

        if keyword_matches:
            keyword_matches.sort(key=lambda item: (min(point[1] for point in item.box), -item.score))
            match = keyword_matches[0]
            text_left, text_top, text_right, text_bottom = self._get_box_bounds(match.box)

            text_width = max(1, text_right - text_left)
            text_height = max(1, text_bottom - text_top)
            tuning = self._get_search_tuning()

            # 由“搜索”文本反推出整块输入框矩形，而不是直接点在文字上。
            search_box_left = self._clamp(
                text_left - int(text_width * tuning["left_expand_ratio"]), 8, image.width - 40
            )
            search_box_right = self._clamp(
                text_right + int(text_width * tuning["right_expand_ratio"]),
                search_box_left + 80,
                image.width - 8,
            )
            search_box_top = self._clamp(
                text_top - int(text_height * tuning["top_expand_ratio"]), 8, image.height - 24
            )
            search_box_bottom = self._clamp(
                text_bottom + int(text_height * tuning["bottom_expand_ratio"]),
                search_box_top + 28,
                image.height - 8,
            )

            # 点击搜索框内部偏左中部，更接近真实用户行为。
            click_x = int(
                search_box_left + (search_box_right - search_box_left) * tuning["click_x_ratio"]
            )
            click_y = int(
                search_box_top + (search_box_bottom - search_box_top) * tuning["click_y_ratio"]
            )
            return {
                "point": (actual_region[0] + click_x, actual_region[1] + click_y),
                "match": match,
                "text_box": (text_left, text_top, text_right, text_bottom),
                "search_box": (search_box_left, search_box_top, search_box_right, search_box_bottom),
                "tuning": tuning,
            }

        if self.vlm.enabled:
            prompt = (
                "请只回答 Yes 或 No。当前截图中是否能清楚看到微信左上角的搜索输入框？"
                "如果能看到，请仅根据搜索输入框所在位置作出判断；如果看不清或没有，回答 No。"
            )
            if self.vlm.verify(image, prompt):
                # VLM 只作兜底存在性确认，不直接报坐标。
                # 这里退回到左侧栏顶部常见搜索位置。
                fallback_x = actual_region[0] + int(actual_region[2] * 0.48)
                fallback_y = actual_region[1] + int(actual_region[3] * 0.16)
                return {
                    "point": (fallback_x, fallback_y),
                    "match": None,
                    "text_box": None,
                    "search_box": None,
                    "tuning": self._get_search_tuning(),
                }

        return None

    def _locate_search_box(self) -> Optional[Tuple[int, int]]:
        image, actual_region = self._capture_region(self._get_region_config("sidebar_region"))
        matches = self.ocr.recognize(image)
        geometry = self._calculate_search_box_geometry(image, actual_region, matches)
        return geometry["point"] if geometry else None

    def _render_search_debug_overlay(
        self,
        image,
        text_box: Optional[Tuple[int, int, int, int]],
        search_box: Optional[Tuple[int, int, int, int]],
        point: Optional[Tuple[int, int]],
        actual_region: Tuple[int, int, int, int],
    ) -> str:
        canvas = image.copy()
        draw = ImageDraw.Draw(canvas)

        if text_box:
            draw.rectangle(text_box, outline="red", width=2)
        if search_box:
            draw.rectangle(search_box, outline="lime", width=3)
        if point:
            local_x = point[0] - actual_region[0]
            local_y = point[1] - actual_region[1]
            draw.ellipse((local_x - 6, local_y - 6, local_x + 6, local_y + 6), fill="blue")
            draw.line((local_x - 14, local_y, local_x + 14, local_y), fill="blue", width=2)
            draw.line((local_x, local_y - 14, local_x, local_y + 14), fill="blue", width=2)

        debug_path = os.path.join(os.path.dirname(__file__), "debug_search_box_overlay.png")
        canvas.save(debug_path)
        return debug_path

    def debug_search_box(self) -> Dict[str, Any]:
        if not self.initialize():
            return {"ok": False, "reason": "initialize_failed"}
        if not self.activate_application():
            return {"ok": False, "reason": "activate_failed"}
        self._set_temporary_topmost(True)
        try:
            region = self._get_region_config("sidebar_region")
            image, actual_region = self._capture_region(region)
            debug_path = os.path.join(os.path.dirname(__file__), "debug_search_box.png")
            image.save(debug_path)
            matches = self.ocr.recognize(image)
            geometry = self._calculate_search_box_geometry(image, actual_region, matches)
            point = geometry["point"] if geometry else None
            overlay_path = self._render_search_debug_overlay(
                image,
                geometry["text_box"] if geometry else None,
                geometry["search_box"] if geometry else None,
                point,
                actual_region,
            )
            return {
                "ok": point is not None,
                "region_config": region,
                "actual_region": actual_region,
                "point": point,
                "image_path": debug_path,
                "overlay_path": overlay_path,
                "search_tuning": geometry["tuning"] if geometry else self._get_search_tuning(),
                "ocr_texts": [
                    {"text": item.text, "score": round(item.score, 4)}
                    for item in matches
                ],
            }
        finally:
            self._set_temporary_topmost(False)

    def debug_search_results(self, target_name: str) -> Dict[str, Any]:
        if not self.initialize():
            return {"ok": False, "reason": "initialize_failed"}
        if not self.activate_application():
            return {"ok": False, "reason": "activate_failed"}

        self._set_temporary_topmost(True)
        try:
            if not self._clear_search_box():
                return {"ok": False, "reason": "search_box_failed"}

            self.human.human_type_text(
                target_name,
                use_clipboard=self._should_use_clipboard_for_search(target_name),
            )
            self.human.human_delay(self.result_refresh_delay, 0.25)

            region = self._get_region_config("search_results_region")
            image, actual_region = self._capture_region(region)
            debug_path = os.path.join(os.path.dirname(__file__), "debug_search_results.png")
            image.save(debug_path)

            matches = self.ocr.recognize(image)
            target_norm = self._normalize_text(target_name)
            candidate_texts = []
            for item in matches:
                normalized = self._normalize_text(item.text)
                candidate_texts.append(
                    {
                        "text": item.text,
                        "normalized": normalized,
                        "score": round(item.score, 4),
                        "exact": normalized == target_norm,
                        "contains_target": target_norm in normalized,
                    }
                )

            return {
                "ok": True,
                "target_name": target_name,
                "region_config": region,
                "actual_region": actual_region,
                "image_path": debug_path,
                "ocr_texts": candidate_texts,
            }
        finally:
            self._set_temporary_topmost(False)

    def _pick_best_ocr_match(self, target_name: str) -> Optional[Dict[str, Any]]:
        if not self._ensure_wechat_foreground():
            return None

        image, actual_region = self._capture_region(self._get_region_config("search_results_region"))
        matches = self.ocr.recognize(image)
        return self._calculate_result_row_geometry(image, actual_region, matches, target_name)

    def _calculate_result_row_geometry(
        self,
        image,
        actual_region: Tuple[int, int, int, int],
        matches: List[OCRMatch],
        target_name: str,
    ) -> Optional[Dict[str, Any]]:
        target_norm = self._normalize_text(target_name)

        exact_matches = [
            item
            for item in matches
            if self._normalize_text(item.text) == target_norm and item.score >= self.ocr_threshold
        ]
        if not exact_matches:
            logger.error("OCR 未找到群名“%s”的精确匹配", target_name)
            return None

        # 搜索结果列表里，越靠上通常越接近目标。先取最靠上的高置信结果。
        exact_matches.sort(key=lambda item: (min(point[1] for point in item.box), -item.score))
        match = exact_matches[0]

        text_left, text_top, text_right, text_bottom = self._get_box_bounds(match.box)
        text_width = max(1, text_right - text_left)
        text_height = max(1, text_bottom - text_top)
        tuning = self._get_result_tuning()

        # 由名字文字框反推出整行结果区域，点击整行中部偏左，而不是点字。
        row_left = self._clamp(
            text_left - max(int(tuning["left_expand_min"]), int(text_width * tuning["left_expand_ratio"])),
            8,
            image.width - 120,
        )
        row_right = self._clamp(image.width - 8, row_left + 120, image.width)
        row_top = self._clamp(
            text_top - max(int(tuning["top_expand_min"]), int(text_height * tuning["top_expand_ratio"])),
            0,
            image.height - 30,
        )
        row_bottom = self._clamp(
            text_bottom + max(int(tuning["bottom_expand_min"]), int(text_height * tuning["bottom_expand_ratio"])),
            row_top + 30,
            image.height,
        )

        row_click_x = int(row_left + (row_right - row_left) * tuning["click_x_ratio"])
        row_click_x = self._clamp(row_click_x, row_left + 36, row_right - 24)
        row_click_y = int(row_top + (row_bottom - row_top) * tuning["click_y_ratio"])

        row_image = image.crop((row_left, row_top, row_right, row_bottom))

        return {
            "match": match,
            "screen_point": (
                actual_region[0] + row_click_x,
                actual_region[1] + row_click_y,
            ),
            "row_image": row_image,
            "actual_region": actual_region,
            "row_box": (row_left, row_top, row_right, row_bottom),
            "result_tuning": tuning,
        }

    def _verify_candidate_with_vlm(self, target_name: str, row_image) -> bool:
        if not self.vlm.enabled:
            return True
        prompt = (
            f"请只回答 Yes 或 No。图中这一条微信搜索结果，是否明确就是目标群/好友“{target_name}”？"
            "只有在文字与目标完全一致时回答 Yes，模糊、截断、看不清都回答 No。"
        )
        return self.vlm.verify(row_image, prompt)

    def _verify_chat_title(self, target_name: str) -> bool:
        if not self._ensure_wechat_foreground():
            return False

        image, _ = self._capture_region(self._get_region_config("chat_title_region"))
        matches = self.ocr.recognize(image)
        target_norm = self._normalize_chat_title_text(target_name)
        title_ok = any(
            self._normalize_chat_title_text(item.text) == target_norm
            and item.score >= self.chat_title_threshold
            for item in matches
        )
        if not title_ok:
            logger.error("聊天标题 OCR 复核失败，目标=%s，识别结果=%s", target_name, [m.text for m in matches])
            return False

        if self.vlm.enabled:
            prompt = (
                f"请只回答 Yes 或 No。图中微信聊天标题是否明确显示为“{target_name}”？"
                "只有完全一致时回答 Yes。"
            )
            if not self.vlm.verify(image, prompt):
                logger.error("VLM 标题复核失败")
                return False

        return True

    def search_group(self, group_name: str) -> bool:
        try:
            if not self.initialize():
                return False
            if not self._require_calibration():
                return False
            if not self.activate_application():
                return False

            self._set_temporary_topmost(True)

            if not self._clear_search_box():
                return False

            self.human.human_type_text(
                group_name,
                use_clipboard=self._should_use_clipboard_for_search(group_name),
            )
            self.human.human_delay(self.result_refresh_delay, 0.25)

            candidate = self._pick_best_ocr_match(group_name)
            if not candidate:
                return False

            if not self._verify_candidate_with_vlm(group_name, candidate["row_image"]):
                logger.error("VLM 复核未通过，放弃点击")
                return False

            if not self._ensure_wechat_foreground():
                return False
            x, y = candidate["screen_point"]
            self.human.human_click(x, y)
            self.human.human_delay(self.post_click_delay, 0.25)

            return self._verify_chat_title(group_name)
        except Exception as exc:
            logger.error("搜索群聊失败: %s", exc)
            return False
        finally:
            if self._topmost_enabled:
                self._set_temporary_topmost(False)

    def _paste_message_humanly(self, message: str) -> bool:
        if not self._click_anchor("chat_input_anchor"):
            return False
        self.human.human_delay(0.25, 0.08)
        pyperclip.copy(message)
        self.human.human_delay(0.15, 0.05)
        self.human.human_hotkey("ctrl", "v")
        self.human.human_delay(0.35, 0.08)
        return True

    def _click_send_button(self) -> bool:
        if "send_button_anchor" in self.calibration:
            return self._click_anchor("send_button_anchor")

        if not self._ensure_wechat_foreground():
            return False

        region = self._get_region_config("send_button_region")
        x, y, w, h = self._normalized_to_screen_region(region)
        self.human.human_click(x + int(w * 0.72), y + int(h * 0.55))
        return True

    def send_message(self, message: str, target_group: str = None) -> bool:
        target_name = target_group or self.default_group
        try:
            if not self.search_group(target_name):
                return False
            if not self._verify_chat_title(target_name):
                return False

            formatted_message = self.format_report_message(message)
            if not self._paste_message_humanly(formatted_message):
                return False
            if not self._click_send_button():
                return False

            self.human.human_delay(self.post_send_delay, 0.2)
            logger.info("消息发送完成: %s", target_name)
            return True
        except Exception as exc:
            logger.error("发送消息失败: %s", exc)
            return False
        finally:
            self._set_temporary_topmost(False)

    def cleanup(self) -> bool:
        try:
            self._set_temporary_topmost(False)
            self.wechat_process = None
            self.wechat_pid = None
            self.main_window_hwnd = None
            self.is_initialized = False
            return True
        except Exception as exc:
            logger.error("清理资源失败: %s", exc)
            return False

    def auto_send_daily_report(self, group_name: str = None) -> bool:
        try:
            report_path = self.config.get("report_file")
            if not report_path or not os.path.exists(report_path):
                logger.error("未配置有效的 report_file，v4 暂不自动查找日报")
                return False
            with open(report_path, "r", encoding="utf-8") as file:
                content = file.read()
            return self.send_message(content, group_name or self.default_group)
        except Exception as exc:
            logger.error("自动发送日报失败: %s", exc)
            return False

    def get_debug_info(self) -> Dict[str, Any]:
        return {
            "sender_type": self.sender_type,
            "is_initialized": self.is_initialized,
            "wechat_pid": self.wechat_pid,
            "main_window_hwnd": self.main_window_hwnd,
            "window_rect": self._get_window_rect() if self.main_window_hwnd else None,
            "calibration_loaded": bool(self.calibration),
            "ocr_available": self.ocr.available,
            "vlm_enabled": self.vlm.enabled,
            "config_path": self.config_path,
        }


MessageSenderFactory.register_sender("wechat_v4", WeChatSenderV4)


def main():
    import sys

    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

    sender = WeChatSenderV4()
    if len(sys.argv) < 2:
        print("个人微信自动发送工具 v4")
        print("用法:")
        print("  python wechat_sender_v4.py calibrate      # 首次标定")
        print("  python wechat_sender_v4.py recalibrate    # 重新标定")
        print("  python wechat_sender_v4.py test           # 查看调试信息")
        print("  python wechat_sender_v4.py send <群名> <消息>")
        return

    command = sys.argv[1].lower()
    if command == "calibrate":
        print("标定结果:", "成功" if sender.initialize() and sender.calibrate() else "失败")
    elif command == "recalibrate":
        print("重新标定结果:", "成功" if sender.initialize() and sender.calibrate(force=True) else "失败")
    elif command == "test":
        if sender.initialize():
            print(json.dumps(sender.get_debug_info(), ensure_ascii=False, indent=2))
        else:
            print("初始化失败")
    elif command == "debug-search":
        print(json.dumps(sender.debug_search_box(), ensure_ascii=False, indent=2))
    elif command == "set-search-tuning":
        if len(sys.argv) < 3:
            print("用法: python wechat_sender_v4.py set-search-tuning <click_x_ratio>")
            return
        tuning = sender.set_search_tuning(click_x_ratio=float(sys.argv[2]))
        print(json.dumps(tuning, ensure_ascii=False, indent=2))
    elif command == "debug-results":
        if len(sys.argv) < 3:
            print("用法: python wechat_sender_v4.py debug-results <群名>")
            return
        print(json.dumps(sender.debug_search_results(sys.argv[2]), ensure_ascii=False, indent=2))
    elif command == "send":
        if len(sys.argv) < 4:
            print("用法: python wechat_sender_v4.py send <群名> <消息>")
            return
        group_name = sys.argv[2]
        message = sys.argv[3]
        print("发送结果:", "成功" if sender.send_message(message, group_name) else "失败")
    else:
        print(f"未知命令: {command}")


if __name__ == "__main__":
    main()
