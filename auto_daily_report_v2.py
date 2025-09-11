# -*- coding: utf-8 -*-
"""
每日存储统计自动化系统 - 多发送器版
版本：v2.0.0
创建日期：2025-09-11
功能：支持个人微信和企业微信的完全自动化存储统计+发送系统
"""

import subprocess
import sys
import os
import time
import json
import logging
from datetime import datetime
from typing import Dict, List, Optional, Any

# 导入新的发送器接口
from message_sender_interface import MessageSenderInterface, MessageSenderFactory
from wechat_sender_v3 import WeChatSenderV3
from wxwork_sender import WXWorkSender

# Windows控制台编码修复
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())

# 配置日志
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('auto_report.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class AutoReportSystemV2:
    def __init__(self):
        """初始化自动化系统 v2.0"""
        self.config_file = "auto_report_config.json"
        self.config = self.load_config()
        self.available_senders = {}
        self.active_sender = None
        
    def load_config(self) -> Dict[str, Any]:
        """加载配置文件 - 支持新旧格式"""
        # 新格式默认配置
        default_config_v2 = {
            "version": "2.0",
            "default_sender": "wechat",
            "sender_priority": ["wechat", "wxwork"],
            "fallback_enabled": True,
            "senders": {
                "wechat": {
                    "type": "wechat",
                    "enabled": True,
                    "process_names": ["WeChat.exe", "Weixin.exe", "wechat.exe"],
                    "default_group": "存储统计报告群",
                    "auto_find_windows": True,
                    "target_groups": [
                        {
                            "name": "存储统计报告群",
                            "hwnd": None,
                            "enabled": True
                        }
                    ]
                },
                "wxwork": {
                    "type": "wxwork",
                    "enabled": True,
                    "process_names": ["WXWork.exe", "wxwork.exe"],
                    "default_group": "存储统计报告群",
                    "auto_find_windows": True,
                    "target_groups": [
                        {
                            "name": "存储统计报告群",
                            "hwnd": None,
                            "enabled": True
                        }
                    ]
                }
            },
            "message_settings": {
                "add_timestamp": True,
                "add_sender_info": True,
                "format_style": "emoji"
            }
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 检查是否为旧格式配置
                if not config.get('version') or config.get('version') != "2.0":
                    logger.info("检测到旧格式配置，正在转换为新格式...")
                    config = self._migrate_legacy_config(config, default_config_v2)
                    
                # 确保新格式配置的完整性
                config = self._merge_config(config, default_config_v2)
                return config
            else:
                # 创建默认配置文件
                self.save_config(default_config_v2)
                return default_config_v2
                
        except Exception as e:
            logger.error(f"加载配置文件失败，使用默认配置: {e}")
            return default_config_v2
    
    def _migrate_legacy_config(self, legacy_config: Dict, default_config: Dict) -> Dict:
        """迁移旧格式配置到新格式"""
        try:
            logger.info("正在迁移旧格式配置...")
            
            # 复制默认配置作为基础
            new_config = json.loads(json.dumps(default_config))
            
            # 保存旧配置到兼容性字段
            new_config["legacy_compatibility"] = legacy_config
            
            # 迁移群聊配置到个人微信发送器
            if "target_groups" in legacy_config:
                new_config["senders"]["wechat"]["target_groups"] = legacy_config["target_groups"]
            
            # 迁移微信进程名
            if "wechat_process_name" in legacy_config:
                process_name = legacy_config["wechat_process_name"]
                if process_name not in new_config["senders"]["wechat"]["process_names"]:
                    new_config["senders"]["wechat"]["process_names"].append(process_name)
            
            # 保存迁移后的配置
            self.save_config(new_config)
            logger.info("✅ 配置迁移完成")
            
            return new_config
            
        except Exception as e:
            logger.error(f"配置迁移失败: {e}")
            return default_config
    
    def _merge_config(self, config: Dict, default_config: Dict) -> Dict:
        """合并配置，确保所有必需字段都存在"""
        def deep_merge(target, source):
            for key, value in source.items():
                if key in target:
                    if isinstance(target[key], dict) and isinstance(value, dict):
                        deep_merge(target[key], value)
                else:
                    target[key] = value
        
        deep_merge(config, default_config)
        return config
    
    def save_config(self, config: Dict = None):
        """保存配置文件"""
        try:
            config_to_save = config or self.config
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, ensure_ascii=False, indent=2)
            logger.info(f"配置已保存到: {self.config_file}")
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")
    
    def initialize_senders(self) -> bool:
        """初始化所有可用的发送器"""
        try:
            logger.info("🔧 初始化发送器...")
            
            self.available_senders = {}
            
            for sender_type, sender_config in self.config["senders"].items():
                if not sender_config.get("enabled", True):
                    logger.info(f"跳过禁用的发送器: {sender_type}")
                    continue
                
                try:
                    # 创建发送器实例
                    sender = MessageSenderFactory.create_sender(sender_type, sender_config)
                    if sender:
                        # 尝试初始化
                        if sender.initialize():
                            self.available_senders[sender_type] = sender
                            logger.info(f"✅ 发送器 {sender_type} 初始化成功")
                        else:
                            logger.warning(f"⚠️ 发送器 {sender_type} 初始化失败")
                    else:
                        logger.error(f"❌ 无法创建发送器: {sender_type}")
                        
                except Exception as e:
                    logger.error(f"初始化发送器 {sender_type} 时出错: {e}")
            
            if not self.available_senders:
                logger.error("❌ 没有可用的发送器")
                return False
            
            logger.info(f"📱 可用发送器: {list(self.available_senders.keys())}")
            return True
            
        except Exception as e:
            logger.error(f"初始化发送器失败: {e}")
            return False
    
    def select_best_sender(self) -> Optional[MessageSenderInterface]:
        """根据配置和优先级选择最佳发送器"""
        try:
            # 首先尝试默认发送器
            default_sender = self.config.get("default_sender", "wechat")
            if default_sender in self.available_senders:
                logger.info(f"✅ 使用默认发送器: {default_sender}")
                return self.available_senders[default_sender]
            
            # 按优先级顺序尝试
            priority_list = self.config.get("sender_priority", ["wechat", "wxwork"])
            for sender_type in priority_list:
                if sender_type in self.available_senders:
                    logger.info(f"✅ 使用优先级发送器: {sender_type}")
                    return self.available_senders[sender_type]
            
            # 如果都不可用，使用第一个可用的
            if self.available_senders:
                sender_type = list(self.available_senders.keys())[0]
                logger.info(f"✅ 使用第一个可用发送器: {sender_type}")
                return self.available_senders[sender_type]
            
            logger.error("❌ 没有可用的发送器")
            return None
            
        except Exception as e:
            logger.error(f"选择发送器失败: {e}")
            return None
    
    def run_storage_statistics(self) -> bool:
        """执行存储统计"""
        try:
            logger.info("🔄 开始执行存储统计...")
            
            # 执行存储统计脚本
            result = subprocess.run([
                sys.executable, 'storage_system.py', 'daily'
            ], capture_output=True, text=True, encoding='gbk', errors='replace')
            
            if result.returncode == 0:
                logger.info("✅ 存储统计执行成功")
                if result.stdout.strip():
                    logger.info(f"输出: {result.stdout}")
                return True
            else:
                logger.error(f"❌ 存储统计执行失败")
                logger.error(f"错误输出: {result.stderr}")
                return False
                
        except Exception as e:
            logger.error(f"执行存储统计失败: {e}")
            return False
    
    def send_reports_with_fallback(self) -> bool:
        """使用回退机制发送报告"""
        try:
            logger.info("📤 开始发送报告...")
            
            # 查找最新报告文件
            report_file = self._find_latest_report()
            if not report_file:
                logger.error("❌ 未找到报告文件")
                return False
            
            # 读取报告内容
            with open(report_file, 'r', encoding='utf-8') as f:
                report_content = f.read()
            
            success_count = 0
            total_attempts = 0
            
            # 按优先级尝试发送器
            sender_priority = self.config.get("sender_priority", ["wechat", "wxwork"])
            fallback_enabled = self.config.get("fallback_enabled", True)
            
            for sender_type in sender_priority:
                if sender_type not in self.available_senders:
                    continue
                
                sender = self.available_senders[sender_type]
                sender_config = self.config["senders"][sender_type]
                target_groups = sender_config.get("target_groups", [])
                
                logger.info(f"📨 尝试使用 {sender_type} 发送器...")
                
                sender_success = False
                for group_config in target_groups:
                    if not group_config.get("enabled", True):
                        continue
                    
                    group_name = group_config["name"]
                    total_attempts += 1
                    
                    try:
                        if sender.send_message(report_content, group_name):
                            logger.info(f"✅ 成功发送到 {sender_type}:{group_name}")
                            success_count += 1
                            sender_success = True
                        else:
                            logger.error(f"❌ 发送到 {sender_type}:{group_name} 失败")
                        
                        # 发送间隔
                        time.sleep(2)
                        
                    except Exception as e:
                        logger.error(f"发送到 {sender_type}:{group_name} 时出错: {e}")
                
                # 如果当前发送器成功发送了至少一条消息，且不启用回退，则停止
                if sender_success and not fallback_enabled:
                    break
            
            logger.info(f"📊 发送完成: {success_count}/{total_attempts} 条消息发送成功")
            return success_count > 0
            
        except Exception as e:
            logger.error(f"发送报告失败: {e}")
            return False
    
    def _find_latest_report(self) -> Optional[str]:
        """查找最新的报告文件"""
        try:
            today = datetime.now().strftime('%Y%m%d')
            report_file = f"storage_report_{today}.txt"
            
            if os.path.exists(report_file):
                logger.info(f"找到今天的报告文件: {report_file}")
                return report_file
            else:
                logger.warning(f"未找到今天的报告文件: {report_file}")
                return None
                
        except Exception as e:
            logger.error(f"查找报告文件失败: {e}")
            return None
    
    def run_full_automation(self) -> bool:
        """运行完整的自动化流程"""
        try:
            logger.info("🚀 开始执行完整自动化流程 v2.0")
            logger.info("=" * 70)
            
            # 步骤1: 初始化发送器
            if not self.initialize_senders():
                logger.error("❌ 初始化发送器失败，终止流程")
                return False
            
            # 步骤2: 执行存储统计
            if not self.run_storage_statistics():
                logger.error("❌ 存储统计失败，终止流程")
                return False
            
            # 步骤3: 发送报告
            if not self.send_reports_with_fallback():
                logger.error("❌ 所有发送器都失败")
                return False
            
            logger.info("=" * 70)
            logger.info("🎉 自动化流程执行完成！")
            return True
            
        except Exception as e:
            logger.error(f"自动化流程执行失败: {e}")
            return False
        finally:
            # 清理资源
            self.cleanup()
    
    def cleanup(self):
        """清理所有发送器资源"""
        try:
            logger.info("🧹 清理发送器资源...")
            for sender_type, sender in self.available_senders.items():
                try:
                    sender.cleanup()
                    logger.info(f"✅ 清理 {sender_type} 完成")
                except Exception as e:
                    logger.error(f"清理 {sender_type} 失败: {e}")
            
            self.available_senders.clear()
            self.active_sender = None
            
        except Exception as e:
            logger.error(f"清理资源失败: {e}")
    
    def show_config(self):
        """显示当前配置"""
        print("📋 当前配置 (v2.0):")
        print(json.dumps(self.config, ensure_ascii=False, indent=2))
    
    def show_senders_status(self):
        """显示发送器状态"""
        print("📱 发送器状态:")
        print("-" * 50)
        
        for sender_type, sender_config in self.config["senders"].items():
            enabled = sender_config.get("enabled", True)
            initialized = sender_type in self.available_senders
            
            status = "🟢" if (enabled and initialized) else "🔴" if enabled else "⚪"
            print(f"{status} {sender_type}: {'启用' if enabled else '禁用'} | {'已初始化' if initialized else '未初始化'}")
            
            if enabled:
                groups = sender_config.get("target_groups", [])
                enabled_groups = [g["name"] for g in groups if g.get("enabled", True)]
                print(f"   目标群聊: {', '.join(enabled_groups) if enabled_groups else '无'}")
    
    def test_senders(self):
        """测试所有发送器"""
        try:
            logger.info("🧪 开始测试发送器...")
            
            if not self.initialize_senders():
                print("❌ 无法初始化任何发送器")
                return False
            
            test_message = f"🧪 发送器测试消息\n发送时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n这是一条测试消息，用于验证发送器功能。"
            
            for sender_type, sender in self.available_senders.items():
                print(f"\n🧪 测试 {sender_type} 发送器:")
                
                sender_config = self.config["senders"][sender_type]
                default_group = sender_config.get("default_group")
                
                if default_group:
                    try:
                        if sender.send_message(test_message, default_group):
                            print(f"✅ {sender_type} 测试成功")
                        else:
                            print(f"❌ {sender_type} 测试失败")
                    except Exception as e:
                        print(f"❌ {sender_type} 测试出错: {e}")
                else:
                    print(f"⚠️ {sender_type} 没有配置默认群聊")
            
            return True
            
        except Exception as e:
            logger.error(f"测试发送器失败: {e}")
            return False
        finally:
            self.cleanup()


def main():
    """主程序入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description="自动化存储统计报告系统 v2.0")
    parser.add_argument('command', nargs='?', default='run', 
                       choices=['run', 'config', 'status', 'test', 'migrate'],
                       help='执行的命令')
    parser.add_argument('--sender', type=str, help='指定使用的发送器类型')
    parser.add_argument('--group', type=str, help='指定目标群聊')
    
    args = parser.parse_args()
    
    system = AutoReportSystemV2()
    
    if args.command == 'run':
        # 执行完整自动化流程
        success = system.run_full_automation()
        sys.exit(0 if success else 1)
        
    elif args.command == 'config':
        # 显示配置
        system.show_config()
        
    elif args.command == 'status':
        # 显示发送器状态
        system.show_senders_status()
        
    elif args.command == 'test':
        # 测试发送器
        success = system.test_senders()
        sys.exit(0 if success else 1)
        
    elif args.command == 'migrate':
        # 手动触发配置迁移
        print("🔄 重新加载并迁移配置...")
        system.load_config()
        print("✅ 配置迁移完成")
        
    else:
        parser.print_help()


if __name__ == "__main__":
    main()