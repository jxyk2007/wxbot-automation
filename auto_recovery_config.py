# -*- coding: utf-8 -*-
"""
企业微信自动恢复配置管理
解决重启后识别不到的问题
"""

import json
import logging
import time
from datetime import datetime
from typing import Dict, List, Optional, Any

logger = logging.getLogger(__name__)

class AutoRecoveryConfig:
    """自动恢复配置管理器"""

    def __init__(self, config_file: str = "auto_report_config.json"):
        self.config_file = config_file
        self.config = {}
        self.load_config()

    def load_config(self):
        """加载配置文件"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            logger.warning(f"配置文件 {self.config_file} 不存在，创建默认配置")
            self._create_default_config()
        except Exception as e:
            logger.error(f"加载配置文件失败: {e}")
            self._create_default_config()

    def _create_default_config(self):
        """创建默认配置"""
        self.config = {
            "version": "2.1",
            "last_update": datetime.now().isoformat(),
            "auto_recovery": {
                "enabled": True,
                "max_retries": 3,
                "retry_delay": 2.0,
                "validate_windows": True,
                "reset_on_failure": True
            },
            "senders": {
                "wxwork": {
                    "type": "wxwork",
                    "enabled": True,
                    "process_names": ["WXWork.exe", "wxwork.exe"],
                    "default_group": "蓝光统计",
                    "auto_find_windows": True,
                    "window_validation": {
                        "check_visibility": True,
                        "check_title": True,
                        "preferred_class": "WeWorkWindow"
                    },
                    "target_groups": []
                }
            }
        }
        self.save_config()

    def save_config(self):
        """保存配置文件"""
        try:
            self.config["last_update"] = datetime.now().isoformat()
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            logger.info(f"配置已保存到: {self.config_file}")
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")

    def invalidate_all_handles(self):
        """清除所有窗口句柄（强制重新检测）"""
        logger.info("清除所有窗口句柄，强制重新检测...")

        for sender_name, sender_config in self.config.get("senders", {}).items():
            if "target_groups" in sender_config:
                for group in sender_config["target_groups"]:
                    if "hwnd" in group:
                        old_hwnd = group["hwnd"]
                        group["hwnd"] = None
                        logger.info(f"清除 {sender_name}.{group.get('name', 'Unknown')} 句柄: {old_hwnd} → None")

        # 同时清除旧版兼容配置
        if "legacy_compatibility" in self.config:
            for group in self.config["legacy_compatibility"].get("target_groups", []):
                if "hwnd" in group:
                    old_hwnd = group["hwnd"]
                    group["hwnd"] = None
                    logger.info(f"清除 legacy.{group.get('name', 'Unknown')} 句柄: {old_hwnd} → None")

        self.save_config()

    def update_window_handle(self, sender_type: str, group_name: str, new_hwnd: int):
        """更新窗口句柄"""
        logger.info(f"更新 {sender_type}.{group_name} 窗口句柄: {new_hwnd}")

        # 确保发送器配置存在
        if sender_type not in self.config.get("senders", {}):
            logger.warning(f"发送器 {sender_type} 不存在于配置中")
            return False

        sender_config = self.config["senders"][sender_type]

        # 查找或创建目标群组
        target_groups = sender_config.setdefault("target_groups", [])

        group_found = False
        for group in target_groups:
            if group.get("name") == group_name:
                old_hwnd = group.get("hwnd")
                group["hwnd"] = new_hwnd
                group["last_update"] = datetime.now().isoformat()
                logger.info(f"✅ 更新群聊 '{group_name}' 句柄: {old_hwnd} → {new_hwnd}")
                group_found = True
                break

        if not group_found:
            # 创建新的群组配置
            new_group = {
                "name": group_name,
                "hwnd": new_hwnd,
                "enabled": True,
                "last_update": datetime.now().isoformat()
            }
            target_groups.append(new_group)
            logger.info(f"✅ 创建新群聊 '{group_name}' 句柄: {new_hwnd}")

        self.save_config()
        return True

    def get_window_handle(self, sender_type: str, group_name: str) -> Optional[int]:
        """获取窗口句柄"""
        try:
            sender_config = self.config.get("senders", {}).get(sender_type, {})
            for group in sender_config.get("target_groups", []):
                if group.get("name") == group_name and group.get("enabled", True):
                    return group.get("hwnd")
        except Exception as e:
            logger.error(f"获取窗口句柄失败: {e}")
        return None

    def is_recovery_enabled(self) -> bool:
        """检查是否启用自动恢复"""
        return self.config.get("auto_recovery", {}).get("enabled", True)

    def get_recovery_settings(self) -> Dict[str, Any]:
        """获取恢复设置"""
        return self.config.get("auto_recovery", {
            "max_retries": 3,
            "retry_delay": 2.0,
            "validate_windows": True,
            "reset_on_failure": True
        })

    def should_validate_windows(self) -> bool:
        """是否应该验证窗口"""
        return self.config.get("auto_recovery", {}).get("validate_windows", True)

# 使用示例
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    config = AutoRecoveryConfig()

    # 模拟企业微信重启后的恢复
    print("模拟企业微信重启...")
    config.invalidate_all_handles()

    # 模拟重新检测到窗口
    print("模拟重新检测窗口...")
    config.update_window_handle("wxwork", "蓝光统计", 123456)

    # 查询窗口句柄
    hwnd = config.get_window_handle("wxwork", "蓝光统计")
    print(f"获取到窗口句柄: {hwnd}")