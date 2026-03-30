# -*- coding: utf-8 -*-
"""
微信点位可视化调参工具

用途:
1. 实时预览搜索框、搜索结果、输入框、发送按钮的蓝点位置。
2. 通过滑块调整参数，并直接保存到 wechat_sender_v4_config.json。
3. 与 WeChatSenderV4 共用同一套配置，保存后主流程立即生效。
"""

import json
import logging
import tkinter as tk
from tkinter import ttk
from typing import Any, Dict, List, Optional, Tuple

import pyautogui
from PIL import ImageDraw, ImageTk

from wechat_sender_v4 import OCRMatch, WeChatSenderV4

logger = logging.getLogger(__name__)


class WeChatVisualTuner:
    def __init__(self):
        self.sender = WeChatSenderV4()
        self.root = tk.Tk()
        self.root.title("微信点位可视化调参工具")
        self.root.geometry("1480x920")

        self.mode_var = tk.StringVar(value="search")
        self.target_var = tk.StringVar(value="ATTEST")
        self.status_var = tk.StringVar(value="准备就绪")

        self.preview_label: Optional[tk.Label] = None
        self.photo = None
        self.slider_frame: Optional[ttk.Frame] = None
        self.sliders: Dict[str, tk.Variable] = {}
        self.current_actual_region: Optional[Tuple[int, int, int, int]] = None
        self.current_geometry: Dict[str, Any] = {}
        self.current_image = None
        self.current_matches: List[OCRMatch] = []
        self.current_target_name = ""
        self.current_window_image = None
        self.current_window_region: Optional[Tuple[int, int, int, int]] = None
        self.current_sidebar_image = None
        self.current_sidebar_region: Optional[Tuple[int, int, int, int]] = None
        self.current_sidebar_matches: List[OCRMatch] = []
        self.current_result_image = None
        self.current_result_region: Optional[Tuple[int, int, int, int]] = None
        self.current_result_matches: List[OCRMatch] = []
        self._refresh_job = None

        self._build_ui()
        self._ensure_sender_ready()
        self._rebuild_sliders()
        self.refresh_preview()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="模式").pack(side=tk.LEFT)
        mode_box = ttk.Combobox(
            top,
            textvariable=self.mode_var,
            values=["search", "result", "input", "send"],
            width=12,
            state="readonly",
        )
        mode_box.pack(side=tk.LEFT, padx=6)
        mode_box.bind("<<ComboboxSelected>>", self._on_mode_changed)

        ttk.Label(top, text="目标名").pack(side=tk.LEFT, padx=(12, 0))
        target_entry = ttk.Entry(top, textvariable=self.target_var, width=24)
        target_entry.pack(side=tk.LEFT, padx=6)
        target_entry.bind("<Return>", lambda _event: self.refresh_preview())

        ttk.Button(top, text="刷新截图", command=self.refresh_preview).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="保存参数", command=self.save_current_mode).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="显示配置", command=self.show_config).pack(side=tk.LEFT, padx=8)

        content = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        content.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(content)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preview_label = tk.Label(left, bg="#202020")
        self.preview_label.pack(fill=tk.BOTH, expand=True)

        right = ttk.Frame(content, width=520)
        right.pack(side=tk.RIGHT, fill=tk.Y)
        right.pack_propagate(False)

        self.slider_frame = ttk.LabelFrame(right, text="参数", padding=10)
        self.slider_frame.pack(fill=tk.X, pady=(0, 10))

        status_frame = ttk.LabelFrame(right, text="状态", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True)

        self.status_text = tk.Text(status_frame, wrap="word", height=28)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        self.status_text.insert("1.0", "等待截图...")
        self.status_text.config(state=tk.DISABLED)

        bottom = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        bottom.pack(fill=tk.X)
        ttk.Label(bottom, textvariable=self.status_var).pack(side=tk.LEFT)

    def _ensure_sender_ready(self) -> None:
        if not self.sender.initialize():
            raise RuntimeError("无法初始化微信发送器，请确认微信已经打开且 PaddleOCR 可用")

    def _slider_specs(self) -> Dict[str, List[Tuple[str, float, float, float]]]:
        return {
            "search": [
                ("left_expand_ratio", 0.5, 4.0, 0.05),
                ("right_expand_ratio", 1.0, 6.0, 0.05),
                ("top_expand_ratio", 0.2, 2.0, 0.05),
                ("bottom_expand_ratio", 0.2, 2.0, 0.05),
                ("click_x_ratio", 0.05, 0.95, 0.01),
                ("click_y_ratio", 0.10, 0.90, 0.01),
            ],
            "result": [
                ("left_expand_ratio", 0.5, 4.0, 0.05),
                ("left_expand_min", 30.0, 180.0, 2.0),
                ("top_expand_ratio", 0.2, 2.0, 0.05),
                ("top_expand_min", 8.0, 60.0, 1.0),
                ("bottom_expand_ratio", 0.2, 2.0, 0.05),
                ("bottom_expand_min", 8.0, 60.0, 1.0),
                ("click_x_ratio", 0.05, 0.95, 0.01),
                ("click_y_ratio", 0.10, 0.90, 0.01),
            ],
            "input": [
                ("x", 0.00, 1.00, 0.002),
                ("y", 0.00, 1.00, 0.002),
            ],
            "send": [
                ("x", 0.00, 1.00, 0.002),
                ("y", 0.00, 1.00, 0.002),
            ],
        }

    def _current_defaults(self) -> Dict[str, float]:
        mode = self.mode_var.get()
        if mode == "search":
            return self.sender._get_search_tuning()
        if mode == "result":
            return self.sender._get_result_tuning()
        if mode == "input":
            return self.sender._get_anchor_config("chat_input_anchor")
        if mode == "send":
            return self.sender._get_anchor_config("send_button_anchor")
        return {}

    def _rebuild_sliders(self) -> None:
        for child in self.slider_frame.winfo_children():
            child.destroy()

        self.sliders = {}
        defaults = self._current_defaults()
        for row_index, (name, start, end, resolution) in enumerate(
            self._slider_specs()[self.mode_var.get()]
        ):
            ttk.Label(self.slider_frame, text=name).grid(row=row_index, column=0, sticky="w")
            variable = tk.DoubleVar(value=float(defaults.get(name, start)))
            scale = tk.Scale(
                self.slider_frame,
                variable=variable,
                from_=start,
                to=end,
                orient=tk.HORIZONTAL,
                resolution=resolution,
                length=340,
                command=lambda _value, key=name: self._schedule_refresh(key),
            )
            scale.grid(row=row_index, column=1, sticky="ew", padx=(6, 0))
            value_label = ttk.Label(self.slider_frame, textvariable=variable, width=8)
            value_label.grid(row=row_index, column=2, sticky="e", padx=(6, 0))
            self.sliders[name] = variable

        self.slider_frame.columnconfigure(1, weight=1)

    def _schedule_refresh(self, _key: str = "") -> None:
        if self._refresh_job is not None:
            self.root.after_cancel(self._refresh_job)
        self._refresh_job = self.root.after(60, self.render_preview_from_cache)

    def _on_mode_changed(self, _event=None) -> None:
        self.current_image = None
        self.current_actual_region = None
        self.current_matches = []
        self.current_target_name = ""
        self.current_window_image = None
        self.current_window_region = None
        self.current_sidebar_image = None
        self.current_sidebar_region = None
        self.current_sidebar_matches = []
        self.current_result_image = None
        self.current_result_region = None
        self.current_result_matches = []
        self._rebuild_sliders()
        self.refresh_preview()

    def _update_status_panel(self, data: Dict[str, Any]) -> None:
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete("1.0", tk.END)
        self.status_text.insert("1.0", json.dumps(self._sanitize_for_json(data), ensure_ascii=False, indent=2))
        self.status_text.config(state=tk.DISABLED)

    def _sanitize_for_json(self, value):
        if isinstance(value, OCRMatch):
            return {
                "text": value.text,
                "score": round(value.score, 4),
                "box": [[int(x), int(y)] for x, y in value.box],
            }
        if isinstance(value, dict):
            return {str(key): self._sanitize_for_json(item) for key, item in value.items()}
        if isinstance(value, (list, tuple)):
            return [self._sanitize_for_json(item) for item in value]
        return value

    def _show_preview(self, image) -> None:
        max_width = 1050
        max_height = 820
        width, height = image.size
        scale = min(max_width / width, max_height / height, 1.0)
        resized = image.resize((int(width * scale), int(height * scale)))
        self.photo = ImageTk.PhotoImage(resized)
        self.preview_label.configure(image=self.photo)

    def _window_anchor_point(self, anchor_name: str, override: Optional[Dict[str, float]] = None) -> Tuple[int, int]:
        anchor = override or self.sender._get_anchor_config(anchor_name)
        left, top, width, height = self.current_window_region
        return int(left + width * anchor["x"]), int(top + height * anchor["y"])

    def _draw_point_marker(
        self,
        draw: ImageDraw.ImageDraw,
        point: Tuple[int, int],
        label: str,
        color: str,
    ) -> None:
        left, top, _width, _height = self.current_window_region
        local_x = point[0] - left
        local_y = point[1] - top
        draw.ellipse((local_x - 8, local_y - 8, local_x + 8, local_y + 8), fill=color)
        draw.line((local_x - 16, local_y, local_x + 16, local_y), fill=color, width=2)
        draw.line((local_x, local_y - 16, local_x, local_y + 16), fill=color, width=2)
        draw.text((local_x + 12, local_y - 20), label, fill=color)

    def _build_full_window_overlay(
        self,
        active_mode: str,
        search_geometry: Optional[Dict[str, Any]] = None,
        result_geometry: Optional[Dict[str, Any]] = None,
        input_anchor: Optional[Dict[str, float]] = None,
        send_anchor: Optional[Dict[str, float]] = None,
    ):
        canvas = self.current_window_image.copy()
        draw = ImageDraw.Draw(canvas)

        points: Dict[str, Tuple[int, int]] = {}
        if search_geometry and search_geometry.get("point"):
            points["1 搜索"] = search_geometry["point"]
        if result_geometry and result_geometry.get("screen_point"):
            points["2 结果"] = result_geometry["screen_point"]
        points["3 输入"] = self._window_anchor_point("chat_input_anchor", input_anchor)
        points["4 发送"] = self._window_anchor_point("send_button_anchor", send_anchor)

        active_map = {
            "search": "1 搜索",
            "result": "2 结果",
            "input": "3 输入",
            "send": "4 发送",
        }
        active_label = active_map.get(active_mode)

        for label, point in points.items():
            color = "lime" if label == active_label else "deepskyblue"
            self._draw_point_marker(draw, point, label, color)

        # 附带绘制当前模式的框体信息，方便判断蓝点是不是落对了区域。
        if active_mode == "search" and search_geometry:
            if search_geometry.get("text_box") and self.current_sidebar_region:
                sx, sy, _sw, _sh = self.current_sidebar_region
                x1, y1, x2, y2 = search_geometry["text_box"]
                draw.rectangle((sx + x1, sy + y1, sx + x2, sy + y2), outline="red", width=2)
            if search_geometry.get("search_box") and self.current_sidebar_region:
                sx, sy, _sw, _sh = self.current_sidebar_region
                x1, y1, x2, y2 = search_geometry["search_box"]
                draw.rectangle((sx + x1, sy + y1, sx + x2, sy + y2), outline="lime", width=3)

        if active_mode == "result" and result_geometry:
            if result_geometry.get("match") and self.current_result_region:
                rx, ry, _rw, _rh = self.current_result_region
                x1, y1, x2, y2 = self.sender._get_box_bounds(result_geometry["match"].box)
                draw.rectangle((rx + x1, ry + y1, rx + x2, ry + y2), outline="red", width=2)
            if result_geometry.get("row_box") and self.current_result_region:
                rx, ry, _rw, _rh = self.current_result_region
                x1, y1, x2, y2 = result_geometry["row_box"]
                draw.rectangle((rx + x1, ry + y1, rx + x2, ry + y2), outline="lime", width=3)

        return canvas

    def _apply_search_tuning_preview(self) -> Dict[str, Any]:
        image = self.current_sidebar_image
        actual_region = self.current_sidebar_region
        matches = self.current_sidebar_matches
        tuning = {key: float(var.get()) for key, var in self.sliders.items()}

        old_tuning = self.sender.calibration.get("search_tuning")
        self.sender.calibration["search_tuning"] = tuning
        geometry = self.sender._calculate_search_box_geometry(image, actual_region, matches)
        if old_tuning is None:
            self.sender.calibration.pop("search_tuning", None)
        else:
            self.sender.calibration["search_tuning"] = old_tuning

        canvas = self._build_full_window_overlay(
            "search",
            search_geometry=geometry,
            input_anchor=self.sender._get_anchor_config("chat_input_anchor"),
            send_anchor=self.sender._get_anchor_config("send_button_anchor"),
        )
        self._show_preview(canvas)
        return {
            "mode": "search",
            "actual_region": actual_region,
            "ocr_texts": [{"text": item.text, "score": round(item.score, 4)} for item in matches],
            "geometry": geometry,
            "tuning": tuning,
        }

    def _prepare_result_region(self, target_name: str):
        if not target_name.strip():
            raise ValueError("结果位置调参需要输入目标群名/好友名")
        if not self.sender._clear_search_box():
            raise RuntimeError("无法点击搜索框")
        self.sender.human.human_type_text(
            target_name,
            use_clipboard=self.sender._should_use_clipboard_for_search(target_name),
        )
        self.sender.human.human_delay(self.sender.result_refresh_delay, 0.25)
        region = self.sender._get_region_config("search_results_region")
        return self.sender._capture_region(region)

    def _apply_result_tuning_preview(self) -> Dict[str, Any]:
        image = self.current_result_image
        actual_region = self.current_result_region
        matches = self.current_result_matches
        tuning = {key: float(var.get()) for key, var in self.sliders.items()}

        old_tuning = self.sender.calibration.get("result_tuning")
        self.sender.calibration["result_tuning"] = tuning
        geometry = self.sender._calculate_result_row_geometry(
            image, actual_region, matches, self.current_target_name
        )
        if old_tuning is None:
            self.sender.calibration.pop("result_tuning", None)
        else:
            self.sender.calibration["result_tuning"] = old_tuning

        search_geometry = None
        if self.current_sidebar_image is not None:
            search_geometry = self.sender._calculate_search_box_geometry(
                self.current_sidebar_image, self.current_sidebar_region, self.current_sidebar_matches
            )
        canvas = self._build_full_window_overlay(
            "result",
            search_geometry=search_geometry,
            result_geometry=geometry,
            input_anchor=self.sender._get_anchor_config("chat_input_anchor"),
            send_anchor=self.sender._get_anchor_config("send_button_anchor"),
        )
        self._show_preview(canvas)
        return {
            "mode": "result",
            "target_name": self.current_target_name,
            "actual_region": actual_region,
            "ocr_texts": [{"text": item.text, "score": round(item.score, 4)} for item in matches],
            "geometry": geometry,
            "tuning": tuning,
        }

    def _apply_anchor_preview(self, anchor_name: str, title: str) -> Dict[str, Any]:
        tuning = {key: float(var.get()) for key, var in self.sliders.items()}
        search_geometry = None
        result_geometry = None
        if self.current_sidebar_image is not None:
            search_geometry = self.sender._calculate_search_box_geometry(
                self.current_sidebar_image, self.current_sidebar_region, self.current_sidebar_matches
            )
        if self.current_result_image is not None and self.current_target_name:
            result_geometry = self.sender._calculate_result_row_geometry(
                self.current_result_image,
                self.current_result_region,
                self.current_result_matches,
                self.current_target_name,
            )
        input_anchor = tuning if anchor_name == "chat_input_anchor" else self.sender._get_anchor_config("chat_input_anchor")
        send_anchor = tuning if anchor_name == "send_button_anchor" else self.sender._get_anchor_config("send_button_anchor")
        canvas = self._build_full_window_overlay(
            title,
            search_geometry=search_geometry,
            result_geometry=result_geometry,
            input_anchor=input_anchor,
            send_anchor=send_anchor,
        )
        self._show_preview(canvas)

        return {
            "mode": title,
            "window_rect": self.current_window_region,
            "anchor_name": anchor_name,
            "anchor": tuning,
        }

    def capture_source(self) -> None:
        self._ensure_sender_ready()
        if not self.sender.activate_application():
            raise RuntimeError("无法激活微信窗口")
        self.sender._set_temporary_topmost(True)
        try:
            mode = self.mode_var.get()
            if mode == "search":
                left, top, right, bottom = self.sender._get_window_rect()
                self.current_window_region = (left, top, right - left, bottom - top)
                self.current_window_image = pyautogui.screenshot(region=self.current_window_region)
                region = self.sender._get_region_config("sidebar_region")
                image, actual_region = self.sender._capture_region(region)
                matches = self.sender.ocr.recognize(image)
                self.current_sidebar_image = image
                self.current_sidebar_region = actual_region
                self.current_sidebar_matches = matches
                self.current_result_image = None
                self.current_result_region = None
                self.current_result_matches = []
                self.current_target_name = ""
            elif mode == "result":
                target_name = self.target_var.get().strip()
                left, top, right, bottom = self.sender._get_window_rect()
                self.current_window_region = (left, top, right - left, bottom - top)
                image, actual_region = self._prepare_result_region(target_name)
                matches = self.sender.ocr.recognize(image)
                self.current_window_image = pyautogui.screenshot(region=self.current_window_region)
                sidebar_region = self.sender._get_region_config("sidebar_region")
                sidebar_image, sidebar_actual_region = self.sender._capture_region(sidebar_region)
                sidebar_matches = self.sender.ocr.recognize(sidebar_image)
                self.current_sidebar_image = sidebar_image
                self.current_sidebar_region = sidebar_actual_region
                self.current_sidebar_matches = sidebar_matches
                self.current_result_image = image
                self.current_result_region = actual_region
                self.current_result_matches = matches
                self.current_target_name = target_name
            elif mode == "input":
                left, top, right, bottom = self.sender._get_window_rect()
                actual_region = (left, top, right - left, bottom - top)
                self.current_window_region = actual_region
                self.current_window_image = pyautogui.screenshot(region=actual_region)
                self.current_sidebar_image = None
                self.current_sidebar_region = None
                self.current_sidebar_matches = []
                self.current_result_image = None
                self.current_result_region = None
                self.current_result_matches = []
                self.current_target_name = ""
            else:
                left, top, right, bottom = self.sender._get_window_rect()
                actual_region = (left, top, right - left, bottom - top)
                self.current_window_region = actual_region
                self.current_window_image = pyautogui.screenshot(region=actual_region)
                self.current_sidebar_image = None
                self.current_sidebar_region = None
                self.current_sidebar_matches = []
                self.current_result_image = None
                self.current_result_region = None
                self.current_result_matches = []
                self.current_target_name = ""
        finally:
            self.sender._set_temporary_topmost(False)

    def render_preview_from_cache(self) -> None:
        self._refresh_job = None
        try:
            if self.current_window_image is None or self.current_window_region is None:
                self.capture_source()

            mode = self.mode_var.get()
            if mode == "search":
                data = self._apply_search_tuning_preview()
            elif mode == "result":
                data = self._apply_result_tuning_preview()
            elif mode == "input":
                data = self._apply_anchor_preview("chat_input_anchor", "input")
            else:
                data = self._apply_anchor_preview("send_button_anchor", "send")

            self._update_status_panel(data)
            self.status_var.set(f"预览已刷新: {mode}")
        except Exception as exc:
            self.status_var.set(f"刷新失败: {exc}")
            self._update_status_panel({"ok": False, "error": str(exc)})

    def refresh_preview(self) -> None:
        try:
            self.capture_source()
            self.render_preview_from_cache()
        except Exception as exc:
            self.status_var.set(f"刷新失败: {exc}")
            self._update_status_panel({"ok": False, "error": str(exc)})

    def save_current_mode(self) -> None:
        try:
            mode = self.mode_var.get()
            values = {key: float(var.get()) for key, var in self.sliders.items()}
            if mode == "search":
                self.sender.set_search_tuning(**values)
            elif mode == "result":
                self.sender.set_result_tuning(**values)
            elif mode == "input":
                self.sender.calibration["chat_input_anchor"] = values
                self.sender._save_calibration()
            else:
                self.sender.calibration["send_button_anchor"] = values
                self.sender._save_calibration()

            self.status_var.set(f"参数已保存: {mode}")
            self.render_preview_from_cache()
        except Exception as exc:
            self.status_var.set(f"保存失败: {exc}")

    def show_config(self) -> None:
        self.sender._load_calibration()
        self._update_status_panel(self.sender.calibration)
        self.status_var.set("当前配置已显示")

    def run(self) -> None:
        self.root.mainloop()


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    app = WeChatVisualTuner()
    app.run()


if __name__ == "__main__":
    main()
