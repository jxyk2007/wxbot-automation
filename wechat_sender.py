# -*- coding: utf-8 -*-
"""
微信自动发送模块
版本：v1.0.0
创建日期：2025-09-10
功能：读取存储报告文件，通过pyautogui自动发送到微信群
"""

import os
import time
import pyautogui
import pyperclip
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WeChatSender:
    def __init__(self):
        """初始化微信发送器"""
        # 设置pyautogui安全配置
        pyautogui.FAILSAFE = True  # 鼠标移到左上角退出
        pyautogui.PAUSE = 1  # 每次操作间隔1秒
        
        # 微信窗口配置（需要根据实际情况调整）
        self.wechat_config = {
            'search_box': (300, 100),  # 搜索框位置
            'chat_input': (500, 600),  # 聊天输入框位置
            'send_button': (900, 600)  # 发送按钮位置
        }
        
        # 群名称
        self.group_name = "存储统计报告群"  # 需要修改为实际群名
    
    def find_latest_report(self):
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
    
    def read_report_content(self, report_file):
        """读取报告文件内容"""
        try:
            with open(report_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            logger.info(f"成功读取报告文件，内容长度: {len(content)}")
            return content
            
        except Exception as e:
            logger.error(f"读取报告文件失败: {e}")
            return None
    
    def check_wechat_window(self):
        """检查微信窗口是否打开"""
        try:
            # 尝试激活微信窗口
            windows = pyautogui.getWindowsWithTitle('微信')
            if not windows:
                logger.error("未找到微信窗口，请先打开微信")
                return False
            
            # 激活微信窗口
            wechat_window = windows[0]
            wechat_window.activate()
            time.sleep(2)
            
            logger.info("微信窗口已激活")
            return True
            
        except Exception as e:
            logger.error(f"检查微信窗口失败: {e}")
            return False
    
    def find_and_enter_group(self, group_name):
        """查找并进入指定群聊"""
        try:
            logger.info(f"正在查找群聊: {group_name}")
            
            # 点击搜索框
            pyautogui.click(self.wechat_config['search_box'])
            time.sleep(1)
            
            # 清空搜索框并输入群名
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.write(group_name)
            time.sleep(2)
            
            # 按回车进入第一个搜索结果
            pyautogui.press('enter')
            time.sleep(2)
            
            logger.info(f"已进入群聊: {group_name}")
            return True
            
        except Exception as e:
            logger.error(f"进入群聊失败: {e}")
            return False
    
    def send_message(self, message):
        """发送消息到当前聊天窗口"""
        try:
            logger.info("准备发送消息")
            
            # 点击输入框
            pyautogui.click(self.wechat_config['chat_input'])
            time.sleep(1)
            
            # 使用剪贴板发送长文本（避免中文输入问题）
            pyperclip.copy(message)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            
            # 发送消息
            pyautogui.hotkey('ctrl', 'enter')  # 或者点击发送按钮
            time.sleep(1)
            
            logger.info("消息发送成功")
            return True
            
        except Exception as e:
            logger.error(f"发送消息失败: {e}")
            return False
    
    def format_report_for_wechat(self, content):
        """格式化报告内容适合微信发送"""
        try:
            # 添加发送时间戳
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            formatted_content = f"""📊 存储使用量统计报告
🕐 发送时间: {timestamp}

{content}

🤖 此消息由自动化系统发送"""
            
            return formatted_content
            
        except Exception as e:
            logger.error(f"格式化报告内容失败: {e}")
            return content
    
    def configure_positions(self):
        """配置微信界面坐标位置（交互式）"""
        print("\n=== 微信坐标配置助手 ===")
        print("请按照提示操作，将鼠标移动到指定位置后按回车确认")
        
        try:
            input("1. 请将鼠标移动到微信搜索框位置，然后按回车...")
            self.wechat_config['search_box'] = pyautogui.position()
            print(f"搜索框坐标已设置: {self.wechat_config['search_box']}")
            
            input("2. 请将鼠标移动到聊天输入框位置，然后按回车...")
            self.wechat_config['chat_input'] = pyautogui.position()
            print(f"输入框坐标已设置: {self.wechat_config['chat_input']}")
            
            input("3. 请将鼠标移动到发送按钮位置，然后按回车...")
            self.wechat_config['send_button'] = pyautogui.position()
            print(f"发送按钮坐标已设置: {self.wechat_config['send_button']}")
            
            print("配置完成！")
            
        except KeyboardInterrupt:
            print("\n配置已取消")
    
    def auto_send_daily_report(self, group_name=None):
        """自动发送每日报告"""
        try:
            logger.info("开始执行自动发送每日报告")
            
            # 使用指定群名或默认群名
            target_group = group_name or self.group_name
            
            # 1. 查找最新报告文件
            report_file = self.find_latest_report()
            if not report_file:
                logger.error("未找到报告文件，发送失败")
                return False
            
            # 2. 读取报告内容
            content = self.read_report_content(report_file)
            if not content:
                logger.error("读取报告内容失败，发送失败")
                return False
            
            # 3. 检查微信窗口
            if not self.check_wechat_window():
                logger.error("微信窗口检查失败，发送失败")
                return False
            
            # 4. 查找并进入群聊
            if not self.find_and_enter_group(target_group):
                logger.error("进入群聊失败，发送失败")
                return False
            
            # 5. 格式化并发送消息
            formatted_content = self.format_report_for_wechat(content)
            if not self.send_message(formatted_content):
                logger.error("发送消息失败")
                return False
            
            logger.info("每日报告发送成功！")
            return True
            
        except Exception as e:
            logger.error(f"自动发送每日报告失败: {e}")
            return False

# ==================== 命令行接口 ====================
def main():
    """主程序入口"""
    import sys
    
    sender = WeChatSender()
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "send":
            # 发送今日报告
            group_name = sys.argv[2] if len(sys.argv) > 2 else None
            success = sender.auto_send_daily_report(group_name)
            if success:
                print("✅ 报告发送成功！")
            else:
                print("❌ 报告发送失败！")
                
        elif command == "config":
            # 配置坐标
            sender.configure_positions()
            
        elif command == "test":
            # 测试功能
            print("测试微信窗口检查...")
            if sender.check_wechat_window():
                print("✅ 微信窗口检查成功")
            else:
                print("❌ 微信窗口检查失败")
                
        else:
            print("未知命令。可用命令：")
            print("  send [群名] - 发送今日报告")
            print("  config - 配置微信界面坐标")
            print("  test - 测试功能")
    else:
        print("微信自动发送工具")
        print("用法:")
        print("  python wechat_sender.py send [群名] - 发送今日报告")
        print("  python wechat_sender.py config - 配置界面坐标") 
        print("  python wechat_sender.py test - 测试功能")

if __name__ == "__main__":
    main()