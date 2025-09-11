# -*- coding: utf-8 -*-
"""
个人微信自动发送模块 - 接口版
版本：v3.0.0
创建日期：2025-09-11
功能：基于通用接口的个人微信自动发送功能（WeChat.exe/Weixin.exe）
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

from message_sender_interface import MessageSenderInterface, MessageSenderFactory, SendResult

# 配置日志
logger = logging.getLogger(__name__)

class WeChatSenderV3(MessageSenderInterface):
    """个人微信发送器 v3.0"""
    
    def __init__(self, config: Dict[str, Any] = None):
        """初始化个人微信发送器"""
        super().__init__(config)
        
        # 设置pyautogui安全配置
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # 微信进程和窗口信息
        self.wechat_process = None
        self.wechat_pid = None
        self.main_window_hwnd = None
        
        # 默认配置
        self.process_names = ["WeChat.exe", "Weixin.exe", "wechat.exe"]
        self.default_group = config.get('default_group', '存储统计报告群') if config else '存储统计报告群'
        
    def initialize(self) -> bool:
        """初始化个人微信发送器"""
        try:
            logger.info("初始化个人微信发送器...")
            
            # 查找微信进程
            if not self.find_target_process():
                logger.error("未找到个人微信进程")
                return False
            
            # 查找微信窗口
            if not self._find_wechat_windows():
                logger.error("未找到个人微信窗口")
                return False
            
            self.is_initialized = True
            logger.info("个人微信发送器初始化成功")
            return True
            
        except Exception as e:
            logger.error(f"初始化个人微信发送器失败: {e}")
            return False
    
    def find_target_process(self) -> bool:
        """查找个人微信进程"""
        try:
            wechat_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    proc_name = proc.info['name']
                    if any(name.lower() in proc_name.lower() for name in self.process_names):
                        wechat_processes.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not wechat_processes:
                logger.error("未找到个人微信进程，请先启动微信")
                return False
            
            # 选择第一个微信进程
            self.wechat_process = wechat_processes[0]
            self.wechat_pid = self.wechat_process.pid
            logger.info(f"找到个人微信进程 PID: {self.wechat_pid}")
            return True
            
        except Exception as e:
            logger.error(f"查找个人微信进程失败: {e}")
            return False
    
    def _find_wechat_windows(self) -> bool:
        """查找个人微信窗口"""
        try:
            if not self.wechat_pid:
                logger.error("请先查找个人微信进程")
                return False
            
            # 枚举所有窗口
            windows_list = []
            win32gui.EnumWindows(self._enum_windows_callback, windows_list)
            
            # 查找属于微信进程的窗口
            wechat_windows = [w for w in windows_list if w['pid'] == self.wechat_pid]
            
            if not wechat_windows:
                logger.error("未找到个人微信窗口")
                return False
            
            # 查找主窗口（通常类名包含WeChatMainWndForPC）
            main_windows = [w for w in wechat_windows if 'WeChatMainWndForPC' in w['class']]
            if main_windows:
                self.main_window_hwnd = main_windows[0]['hwnd']
                logger.info(f"找到个人微信主窗口: {main_windows[0]['title']}")
            else:
                # 备选方案：选择第一个有标题的窗口
                titled_windows = [w for w in wechat_windows if w['title'].strip()]
                if titled_windows:
                    self.main_window_hwnd = titled_windows[0]['hwnd']
                    logger.info(f"使用备选个人微信窗口: {titled_windows[0]['title']}")
                else:
                    logger.error("无法确定个人微信主窗口")
                    return False
            
            return self.main_window_hwnd is not None
            
        except Exception as e:
            logger.error(f"查找个人微信窗口失败: {e}")
            return False
    
    def _enum_windows_callback(self, hwnd, windows_list):
        """枚举窗口回调函数"""
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            
            # 获取窗口所属进程ID
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            
            windows_list.append({
                'hwnd': hwnd,
                'title': window_text,
                'class': class_name,
                'pid': window_pid
            })
    
    def activate_application(self) -> bool:
        """激活个人微信窗口"""
        try:
            if not self.main_window_hwnd:
                logger.error("个人微信窗口句柄不存在")
                return False
            
            # 检查窗口是否最小化
            if win32gui.IsIconic(self.main_window_hwnd):
                win32gui.ShowWindow(self.main_window_hwnd, win32con.SW_RESTORE)
            
            # 激活窗口
            win32gui.SetForegroundWindow(self.main_window_hwnd)
            time.sleep(1)
            
            logger.info("个人微信窗口已激活")
            return True
            
        except Exception as e:
            logger.error(f"激活个人微信窗口失败: {e}")
            return False
    
    def search_group(self, group_name: str) -> bool:
        """搜索并进入个人微信群聊"""
        try:
            logger.info(f"搜索个人微信群聊: {group_name}")
            
            # 激活微信窗口
            if not self.activate_application():
                return False
            
            # 使用快捷键打开搜索（Ctrl+F）
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(1)
            
            # 输入群名搜索
            pyperclip.copy(group_name)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            
            # 按回车选择第一个结果
            pyautogui.press('enter')
            time.sleep(2)
            
            logger.info(f"已搜索并进入个人微信群聊: {group_name}")
            return True
            
        except Exception as e:
            logger.error(f"搜索个人微信群聊失败: {e}")
            return False
    
    def send_message(self, message: str, target_group: str = None) -> bool:
        """发送消息到个人微信"""
        try:
            logger.info("准备发送消息到个人微信")
            
            # 如果指定了目标群聊，先搜索群聊
            if target_group:
                if not self.search_group(target_group):
                    logger.error(f"搜索群聊失败: {target_group}")
                    return False
            
            # 确保微信窗口处于前台
            if not self.activate_application():
                return False
            
            # 将消息复制到剪贴板
            formatted_message = self.format_report_message(message)
            pyperclip.copy(formatted_message)
            time.sleep(0.2)
            
            # 粘贴消息
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            
            # 发送消息（Alt+S）
            pyautogui.hotkey('alt', 's')
            time.sleep(1)
            
            logger.info("个人微信消息发送完成")
            return True
            
        except Exception as e:
            logger.error(f"发送个人微信消息失败: {e}")
            return False
    
    def cleanup(self) -> bool:
        """清理资源"""
        try:
            logger.info("清理个人微信发送器资源")
            self.wechat_process = None
            self.wechat_pid = None
            self.main_window_hwnd = None
            self.is_initialized = False
            return True
        except Exception as e:
            logger.error(f"清理资源失败: {e}")
            return False
    
    def auto_send_daily_report(self, group_name: str = None) -> bool:
        """自动发送每日报告到个人微信"""
        try:
            logger.info("开始执行个人微信自动发送每日报告")
            
            # 使用指定群名或默认群名
            target_group = group_name or self.default_group
            
            # 初始化发送器
            if not self.initialize():
                logger.error("初始化个人微信发送器失败")
                return False
            
            # 查找最新报告文件
            report_file = self._find_latest_report()
            if not report_file:
                logger.error("未找到报告文件，发送失败")
                return False
            
            # 读取报告内容
            content = self._read_report_content(report_file)
            if not content:
                logger.error("读取报告内容失败，发送失败")
                return False
            
            # 发送消息
            if not self.send_message(content, target_group):
                logger.error("发送消息失败")
                return False
            
            logger.info("个人微信每日报告发送成功！")
            return True
            
        except Exception as e:
            logger.error(f"个人微信自动发送每日报告失败: {e}")
            return False
        finally:
            self.cleanup()
    
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
    
    def _read_report_content(self, report_file: str) -> Optional[str]:
        """读取报告文件内容"""
        try:
            with open(report_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            logger.info(f"成功读取报告文件，内容长度: {len(content)}")
            return content
            
        except Exception as e:
            logger.error(f"读取报告文件失败: {e}")
            return None
    
    def get_debug_info(self) -> Dict[str, Any]:
        """获取调试信息"""
        try:
            info = super().get_sender_info()
            
            # 添加个人微信特有信息
            if self.wechat_process:
                info["个人微信进程信息"] = {
                    "PID": self.wechat_pid,
                    "进程名": self.wechat_process.name(),
                    "可执行文件": getattr(self.wechat_process, 'exe', lambda: "无法获取")()
                }
            
            if self.main_window_hwnd:
                info["窗口信息"] = {
                    "窗口句柄": self.main_window_hwnd,
                    "窗口标题": win32gui.GetWindowText(self.main_window_hwnd),
                    "窗口类名": win32gui.GetClassName(self.main_window_hwnd)
                }
            
            return info
            
        except Exception as e:
            logger.error(f"获取调试信息失败: {e}")
            return {"错误": str(e)}
    
    # ==================== 向后兼容的方法 ====================
    def smart_search_group(self, group_name: str) -> bool:
        """向后兼容：智能搜索群聊"""
        return self.search_group(group_name)
    
    def send_message_to_current_chat(self, message: str) -> bool:
        """向后兼容：发送消息到当前聊天"""
        return self.send_message(message)
    
    def interactive_select_process(self) -> bool:
        """向后兼容：交互式选择进程（简化版）"""
        return self.find_target_process()
    
    def interactive_select_window(self) -> bool:
        """向后兼容：交互式选择窗口（简化版）"""
        return self._find_wechat_windows()


# 注册个人微信发送器到工厂
MessageSenderFactory.register_sender("wechat", WeChatSenderV3)


# ==================== 命令行接口 ====================
def main():
    """主程序入口"""
    import sys
    import json
    
    # 配置日志
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    sender = WeChatSenderV3()
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "send":
            # 发送今日报告
            group_name = sys.argv[2] if len(sys.argv) > 2 else None
            success = sender.auto_send_daily_report(group_name)
            if success:
                print("✅ 个人微信报告发送成功！")
            else:
                print("❌ 个人微信报告发送失败！")
                
        elif command == "debug":
            # 获取调试信息
            sender.initialize()
            debug_info = sender.get_debug_info()
            print("=== 个人微信调试信息 ===")
            print(json.dumps(debug_info, ensure_ascii=False, indent=2))
            
        elif command == "test":
            # 测试功能
            print("测试个人微信进程查找...")
            if sender.find_target_process():
                print("✅ 个人微信进程查找成功")
                
                print("测试个人微信窗口查找...")
                if sender._find_wechat_windows():
                    print("✅ 个人微信窗口查找成功")
                    
                    print("测试窗口激活...")
                    if sender.activate_application():
                        print("✅ 窗口激活成功")
                    else:
                        print("❌ 窗口激活失败")
                else:
                    print("❌ 个人微信窗口查找失败")
            else:
                print("❌ 个人微信进程查找失败")
                
        else:
            print("未知命令。可用命令：")
            print("  send [群名] - 发送今日报告到个人微信")
            print("  debug - 获取调试信息")
            print("  test - 测试功能")
    else:
        print("个人微信自动发送工具 v3.0 (接口版)")
        print("用法:")
        print("  python wechat_sender_v3.py send [群名] - 发送今日报告")
        print("  python wechat_sender_v3.py debug - 获取调试信息") 
        print("  python wechat_sender_v3.py test - 测试功能")


if __name__ == "__main__":
    main()