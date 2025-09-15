# -*- coding: utf-8 -*-
"""
通用消息发送接口
版本：v1.0.0
创建日期：2025-09-11
功能：定义消息发送器的通用接口，支持多种发送方式
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any
import logging

logger = logging.getLogger(__name__)

class MessageSenderInterface(ABC):
    """消息发送器接口基类"""
    
    def __init__(self, config: Dict[str, Any] = None):
        """
        初始化消息发送器
        
        Args:
            config: 发送器配置字典
        """
        self.config = config or {}
        self.is_initialized = False
        self.sender_type = self.__class__.__name__
        
    @abstractmethod
    def initialize(self) -> bool:
        """
        初始化发送器
        
        Returns:
            bool: 初始化是否成功
        """
        pass
    
    @abstractmethod
    def find_target_process(self) -> bool:
        """
        查找目标进程
        
        Returns:
            bool: 是否找到目标进程
        """
        pass
    
    @abstractmethod
    def activate_application(self) -> bool:
        """
        激活目标应用程序
        
        Returns:
            bool: 激活是否成功
        """
        pass
    
    @abstractmethod
    def search_group(self, group_name: str) -> bool:
        """
        搜索并进入指定群聊
        
        Args:
            group_name: 群聊名称
            
        Returns:
            bool: 是否成功进入群聊
        """
        pass
    
    @abstractmethod
    def send_message(self, message: str, target_group: str = None) -> bool:
        """
        发送消息
        
        Args:
            message: 要发送的消息内容
            target_group: 目标群聊名称（可选）
            
        Returns:
            bool: 发送是否成功
        """
        pass
    
    @abstractmethod
    def cleanup(self) -> bool:
        """
        清理资源
        
        Returns:
            bool: 清理是否成功
        """
        pass
    
    def get_sender_info(self) -> Dict[str, Any]:
        """
        获取发送器信息
        
        Returns:
            Dict: 发送器信息字典
        """
        return {
            "sender_type": self.sender_type,
            "is_initialized": self.is_initialized,
            "config": self.config
        }
    
    def validate_config(self, required_keys: List[str]) -> bool:
        """
        验证配置是否包含必需的键
        
        Args:
            required_keys: 必需的配置键列表
            
        Returns:
            bool: 配置是否有效
        """
        try:
            missing_keys = [key for key in required_keys if key not in self.config]
            if missing_keys:
                logger.error(f"{self.sender_type} 配置缺少必需的键: {missing_keys}")
                return False
            return True
        except Exception as e:
            logger.error(f"验证配置失败: {e}")
            return False
    
    def format_report_message(self, content: str) -> str:
        """
        格式化报告消息（自然格式）

        Args:
            content: 原始报告内容

        Returns:
            str: 格式化后的消息
        """
        from datetime import datetime

        timestamp = datetime.now().strftime('%m月%d日 %H:%M')

        formatted_content = f"""今日存储统计情况({timestamp}):

{content}

各单位请及时关注存储使用情况。"""

        return formatted_content


class MessageSenderFactory:
    """消息发送器工厂类"""
    
    _senders = {}
    
    @classmethod
    def register_sender(cls, sender_type: str, sender_class: type):
        """
        注册消息发送器
        
        Args:
            sender_type: 发送器类型名称
            sender_class: 发送器类
        """
        cls._senders[sender_type] = sender_class
        logger.info(f"已注册消息发送器: {sender_type}")
    
    @classmethod
    def create_sender(cls, sender_type: str, config: Dict[str, Any] = None) -> Optional[MessageSenderInterface]:
        """
        创建消息发送器实例
        
        Args:
            sender_type: 发送器类型
            config: 配置字典
            
        Returns:
            MessageSenderInterface: 发送器实例，如果类型不存在则返回None
        """
        if sender_type not in cls._senders:
            logger.error(f"未知的发送器类型: {sender_type}")
            return None
        
        try:
            sender_class = cls._senders[sender_type]
            return sender_class(config)
        except Exception as e:
            logger.error(f"创建发送器失败: {e}")
            return None
    
    @classmethod
    def get_available_senders(cls) -> List[str]:
        """
        获取可用的发送器类型列表
        
        Returns:
            List[str]: 发送器类型列表
        """
        return list(cls._senders.keys())


# 发送结果枚举
class SendResult:
    """发送结果常量"""
    SUCCESS = "success"
    FAILED = "failed"
    PROCESS_NOT_FOUND = "process_not_found"
    WINDOW_NOT_FOUND = "window_not_found"
    GROUP_NOT_FOUND = "group_not_found"
    MESSAGE_SEND_FAILED = "message_send_failed"
    INITIALIZATION_FAILED = "initialization_failed"