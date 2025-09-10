# -*- coding: utf-8 -*-
"""
直接发送器 - 基于窗口句柄
版本：v1.0.0
创建日期：2025-09-10
功能：直接使用窗口句柄发送消息，无需复杂的进程查找
"""

import win32gui
import win32con
import win32api
import pyautogui
import pyperclip
import time
import os
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DirectSender:
    def __init__(self):
        """初始化直接发送器"""
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.3
    
    def activate_window_by_handle(self, hwnd):
        """通过句柄激活窗口"""
        try:
            # 检查窗口是否存在
            if not win32gui.IsWindow(hwnd):
                logger.error(f"窗口句柄 {hwnd} 不存在")
                return False
            
            # 获取窗口信息
            window_title = win32gui.GetWindowText(hwnd)
            window_class = win32gui.GetClassName(hwnd)
            logger.info(f"激活窗口: {window_title} (类: {window_class})")
            
            # 检查窗口是否最小化
            if win32gui.IsIconic(hwnd):
                logger.info("窗口已最小化，正在恢复...")
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.5)
            
            # 激活窗口
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            
            # 验证窗口是否成功激活
            current_hwnd = win32gui.GetForegroundWindow()
            if current_hwnd == hwnd:
                logger.info("✅ 窗口激活成功")
                return True
            else:
                logger.warning(f"窗口激活可能失败，当前前台窗口: {current_hwnd}")
                return True  # 继续尝试，可能仍然可用
                
        except Exception as e:
            logger.error(f"激活窗口失败: {e}")
            return False
    
    def get_input_area_position(self, hwnd):
        """获取输入框区域位置（窗口下方80%处）"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            # 水平居中
            input_x = (rect[0] + rect[2]) // 2
            # 垂直位置在窗口下方80%处（通常是输入框位置）
            input_y = rect[1] + int((rect[3] - rect[1]) * 0.8)
            return input_x, input_y
        except Exception as e:
            logger.error(f"获取输入框位置失败: {e}")
            return None, None
    
    def get_window_center(self, hwnd):
        """获取窗口中心坐标"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            center_x = (rect[0] + rect[2]) // 2
            center_y = (rect[1] + rect[3]) // 2
            return center_x, center_y
        except Exception as e:
            logger.error(f"获取窗口中心失败: {e}")
            return None, None
    
    def send_message_to_window(self, hwnd, message):
        """直接向指定窗口句柄发送消息"""
        try:
            logger.info(f"准备向窗口句柄 {hwnd} 发送消息")
            
            # 激活目标窗口
            if not self.activate_window_by_handle(hwnd):
                return False
            
            # 获取窗口信息用于日志
            window_title = win32gui.GetWindowText(hwnd)
            logger.info(f"向窗口 '{window_title}' 发送消息")
            
            # 点击输入框区域确保焦点
            input_x, input_y = self.get_input_area_position(hwnd)
            if input_x and input_y:
                logger.info(f"点击输入框区域: ({input_x}, {input_y})")
                pyautogui.click(input_x, input_y)
                time.sleep(0.5)
            else:
                # 备选方案：点击窗口中心
                center_x, center_y = self.get_window_center(hwnd)
                if center_x and center_y:
                    logger.info(f"备选方案-点击窗口中心: ({center_x}, {center_y})")
                    pyautogui.click(center_x, center_y)
                    time.sleep(0.5)
            
            # 将消息复制到剪贴板
            pyperclip.copy(message)
            time.sleep(0.2)
            
            # 清空输入框（如果有内容）
            logger.info("清空输入框...")
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            
            # 粘贴消息内容
            logger.info("粘贴消息内容...")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            
            # 发送消息 - 按优先级尝试
            logger.info("发送消息...")
            send_success = False
            
            # 方法1: 直接回车（最常用）
            try:
                logger.info("尝试回车发送...")
                pyautogui.press('enter')
                time.sleep(1)
                send_success = True
            except Exception as e:
                logger.warning(f"回车发送失败: {e}")
            
            # 方法2: Ctrl+Enter（备选）
            if not send_success:
                try:
                    logger.info("尝试Ctrl+Enter发送...")
                    pyautogui.hotkey('ctrl', 'enter')
                    time.sleep(1)
                    send_success = True
                except Exception as e:
                    logger.warning(f"Ctrl+Enter发送失败: {e}")
            
            # 方法3: Alt+S（微信专用）
            if not send_success:
                try:
                    logger.info("尝试Alt+S发送...")
                    pyautogui.hotkey('alt', 's')
                    time.sleep(1)
                    send_success = True
                except Exception as e:
                    logger.warning(f"Alt+S发送失败: {e}")
            
            if send_success:
                logger.info("✅ 消息发送完成")
                return True
            else:
                logger.error("❌ 所有发送方法都失败了")
                return False
            
        except Exception as e:
            logger.error(f"发送消息失败: {e}")
            return False
    
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
    
    def format_report_for_wechat(self, content):
        """格式化报告内容"""
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            formatted_content = f"""📊 存储使用量统计报告
🕐 发送时间: {timestamp}

{content}

🤖 此消息由自动化系统发送 (DirectSender v1.0)"""
            
            return formatted_content
            
        except Exception as e:
            logger.error(f"格式化报告内容失败: {e}")
            return content
    
    def send_daily_report_to_window(self, hwnd):
        """向指定窗口句柄发送每日报告"""
        try:
            logger.info(f"开始向窗口句柄 {hwnd} 发送每日报告")
            
            # 验证窗口句柄
            if not win32gui.IsWindow(hwnd):
                logger.error(f"窗口句柄 {hwnd} 无效")
                return False
            
            window_title = win32gui.GetWindowText(hwnd)
            logger.info(f"目标窗口: {window_title}")
            
            # 查找报告文件
            report_file = self.find_latest_report()
            if not report_file:
                logger.error("未找到报告文件，发送失败")
                return False
            
            # 读取报告内容
            content = self.read_report_content(report_file)
            if not content:
                logger.error("读取报告内容失败，发送失败")
                return False
            
            # 格式化内容
            formatted_content = self.format_report_for_wechat(content)
            
            # 发送消息
            if self.send_message_to_window(hwnd, formatted_content):
                logger.info("✅ 每日报告发送成功！")
                return True
            else:
                logger.error("❌ 发送消息失败")
                return False
                
        except Exception as e:
            logger.error(f"发送每日报告失败: {e}")
            return False
    
    def test_message_to_window(self, hwnd, test_message="🤖 测试消息 - DirectSender"):
        """向指定窗口发送测试消息"""
        try:
            logger.info(f"向窗口句柄 {hwnd} 发送测试消息")
            
            if not win32gui.IsWindow(hwnd):
                logger.error(f"窗口句柄 {hwnd} 无效")
                return False
                
            window_title = win32gui.GetWindowText(hwnd)
            logger.info(f"目标窗口: {window_title}")
            
            timestamp = datetime.now().strftime('%H:%M:%S')
            full_message = f"{test_message}\n发送时间: {timestamp}"
            
            return self.send_message_to_window(hwnd, full_message)
            
        except Exception as e:
            logger.error(f"发送测试消息失败: {e}")
            return False

def main():
    """主程序入口"""
    import sys
    
    sender = DirectSender()
    
    if len(sys.argv) < 2:
        print("🚀 DirectSender - 基于窗口句柄的直接发送器")
        print("用法:")
        print("  python direct_sender.py test <窗口句柄>     - 发送测试消息")
        print("  python direct_sender.py send <窗口句柄>     - 发送今日报告")
        print("  python direct_sender.py info <窗口句柄>     - 显示窗口信息")
        print("  python direct_sender.py click <窗口句柄>    - 调试点击位置")
        print("\n示例:")
        print("  python direct_sender.py info 28847092      - 查看AI TESt群窗口信息")
        print("  python direct_sender.py click 28847092     - 测试点击输入框位置")
        print("  python direct_sender.py test 28847092      - 向AI TESt群发送测试消息")
        print("  python direct_sender.py send 28847092      - 向AI TESt群发送报告")
        return
    
    command = sys.argv[1]
    
    if command == "test":
        if len(sys.argv) < 3:
            print("请指定窗口句柄: python direct_sender.py test <窗口句柄>")
            return
            
        try:
            hwnd = int(sys.argv[2])
            success = sender.test_message_to_window(hwnd)
            if success:
                print("✅ 测试消息发送成功！")
            else:
                print("❌ 测试消息发送失败！")
        except ValueError:
            print("无效的窗口句柄，请输入数字")
            
    elif command == "send":
        if len(sys.argv) < 3:
            print("请指定窗口句柄: python direct_sender.py send <窗口句柄>")
            return
            
        try:
            hwnd = int(sys.argv[2])
            success = sender.send_daily_report_to_window(hwnd)
            if success:
                print("✅ 报告发送成功！")
            else:
                print("❌ 报告发送失败！")
        except ValueError:
            print("无效的窗口句柄，请输入数字")
            
    elif command == "info":
        if len(sys.argv) < 3:
            print("请指定窗口句柄: python direct_sender.py info <窗口句柄>")
            return
            
        try:
            hwnd = int(sys.argv[2])
            if win32gui.IsWindow(hwnd):
                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                
                # 计算点击位置
                input_x, input_y = sender.get_input_area_position(hwnd)
                center_x, center_y = sender.get_window_center(hwnd)
                
                print(f"窗口信息:")
                print(f"  句柄: {hwnd}")
                print(f"  标题: {title}")
                print(f"  类名: {class_name}")
                print(f"  位置: {rect}")
                print(f"  可见: {win32gui.IsWindowVisible(hwnd)}")
                print(f"  窗口中心: ({center_x}, {center_y})")
                print(f"  输入框位置: ({input_x}, {input_y})")
            else:
                print(f"窗口句柄 {hwnd} 无效")
        except ValueError:
            print("无效的窗口句柄，请输入数字")
    
    elif command == "click":
        # 调试模式：只点击不发送
        if len(sys.argv) < 3:
            print("请指定窗口句柄: python direct_sender.py click <窗口句柄>")
            return
            
        try:
            hwnd = int(sys.argv[2])
            if not win32gui.IsWindow(hwnd):
                print(f"窗口句柄 {hwnd} 无效")
                return
            
            print("🎯 调试模式：只激活窗口并点击输入框")
            success = sender.activate_window_by_handle(hwnd)
            if success:
                input_x, input_y = sender.get_input_area_position(hwnd)
                print(f"将点击位置: ({input_x}, {input_y})")
                print("3秒后点击...")
                time.sleep(3)
                pyautogui.click(input_x, input_y)
                print("✅ 点击完成，请检查是否获得输入框焦点")
            else:
                print("❌ 窗口激活失败")
        except ValueError:
            print("无效的窗口句柄，请输入数字")
    else:
        print(f"未知命令: {command}")

if __name__ == "__main__":
    main()