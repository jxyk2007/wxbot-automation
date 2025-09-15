# -*- coding: utf-8 -*-
"""
企业微信发送器适配器
让新的 WXWorkSender 兼容原有的 MessageSenderInterface
"""

import logging
from typing import Dict, List, Optional, Any
from message_sender_interface import MessageSenderInterface
from wxwork_sender import WXWorkSenderRobust

logger = logging.getLogger(__name__)

class SendResult:
    """简单的发送结果类"""
    def __init__(self, success: bool, message: str = ""):
        self.success = success
        self.message = message

class WXWorkSenderAdapter(MessageSenderInterface):
    """企业微信发送器适配器"""

    def __init__(self, config: Dict[str, Any] = None):
        """初始化适配器"""
        super().__init__(config)
        self.sender = WXWorkSenderRobust(config)
        self.default_group = config.get('default_group', '蓝光统计') if config else '蓝光统计'

    def initialize(self) -> bool:
        """初始化发送器"""
        try:
            success = self.sender.test_connection()
            self.is_initialized = success
            return success
        except Exception as e:
            logger.error(f"初始化企业微信发送器失败: {e}")
            return False

    def find_target_process(self) -> bool:
        """查找目标进程"""
        try:
            hwnd = self.sender.find_wxwork_window()
            return hwnd is not None
        except Exception as e:
            logger.error(f"查找企业微信进程失败: {e}")
            return False

    def activate_application(self) -> bool:
        """激活应用程序"""
        try:
            hwnd = self.sender.find_wxwork_window()
            if hwnd:
                return self.sender.activate_window(hwnd)
            return False
        except Exception as e:
            logger.error(f"激活企业微信失败: {e}")
            return False

    def search_group(self, group_name: str) -> bool:
        """搜索并进入指定群聊"""
        try:
            hwnd = self.sender.find_wxwork_window()
            if hwnd and self.sender.activate_window(hwnd):
                # 这里可以实现搜索群聊的逻辑，但由于send_message已经包含了搜索
                # 我们直接返回True
                return True
            return False
        except Exception as e:
            logger.error(f"搜索群聊失败: {e}")
            return False

    def send_message(self, message: str, target_group: str = None) -> bool:
        """发送消息"""
        try:
            target_group = target_group or self.default_group
            logger.info(f"准备发送消息到企业微信")
            logger.info(f"搜索企业微信群聊: {target_group}")

            success = self.sender.send_message(message, target_group)

            if success:
                logger.info(f"✅ 成功发送到 wxwork:{target_group}")
                return True
            else:
                logger.error(f"❌ 发送到 wxwork:{target_group} 失败")
                return False

        except Exception as e:
            logger.error(f"发送企业微信消息失败: {e}")
            return False

    def send_msg(self, message: str, target_group: str = None, clear: bool = True,
                 at: List[str] = None, exact: bool = False) -> SendResult:
        """发送消息"""
        try:
            target_group = target_group or self.default_group
            logger.info(f"准备发送消息到企业微信")
            logger.info(f"搜索企业微信群聊: {target_group}")

            success = self.sender.send_message(message, target_group)

            if success:
                logger.info(f"✅ 成功发送到 wxwork:{target_group}")
                return SendResult(True, "发送成功")
            else:
                logger.error(f"❌ 发送到 wxwork:{target_group} 失败")
                return SendResult(False, "发送失败")

        except Exception as e:
            logger.error(f"发送企业微信消息失败: {e}")
            return SendResult(False, f"发送失败: {e}")

    def send_files(self, files: List[str], target_group: str = None, exact: bool = False) -> SendResult:
        """发送文件（暂不支持）"""
        logger.warning("企业微信发送器暂不支持文件发送")
        return SendResult(False, "不支持文件发送")

    def get_info(self) -> Dict[str, Any]:
        """获取发送器信息"""
        return {
            "type": "wxwork",
            "name": "企业微信发送器",
            "version": "2.0.0",
            "is_initialized": self.is_initialized,
            "default_group": self.default_group
        }

    def cleanup(self) -> bool:
        """清理资源"""
        logger.info("清理企业微信发送器资源")
        # 新版本发送器无需特殊清理
        return True

# 注册发送器到工厂
if __name__ == "__main__":
    from message_sender_interface import MessageSenderFactory
    MessageSenderFactory.register_sender("wxwork", WXWorkSenderAdapter)
    print("✅ 企业微信发送器适配器已注册")