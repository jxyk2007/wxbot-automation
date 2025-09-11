# -*- coding: utf-8 -*-
"""
企业微信自动发送模块
版本：v1.0.0
创建日期：2025-09-11
功能：基于通用接口的企业微信自动发送功能（WXWork.exe）
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

class WXWorkSender(MessageSenderInterface):
    """企业微信发送器"""
    
    def __init__(self, config: Dict[str, Any] = None):
        """初始化企业微信发送器"""
        super().__init__(config)
        
        # 设置pyautogui安全配置
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # 企业微信进程和窗口信息
        self.wxwork_process = None
        self.wxwork_pid = None
        self.main_window_hwnd = None
        
        # 默认配置
        self.process_names = ["WXWork.exe", "wxwork.exe"]
        self.default_group = config.get('default_group', '存储统计报告群') if config else '存储统计报告群'
        
    def initialize(self) -> bool:
        """初始化企业微信发送器"""
        try:
            logger.info("初始化企业微信发送器...")
            
            # 查找企业微信进程
            if not self.find_target_process():
                logger.error("未找到企业微信进程")
                return False
            
            # 查找企业微信窗口
            if not self._find_wxwork_windows():
                logger.error("未找到企业微信窗口")
                return False
            
            self.is_initialized = True
            logger.info("企业微信发送器初始化成功")
            return True
            
        except Exception as e:
            logger.error(f"初始化企业微信发送器失败: {e}")
            return False
    
    def find_target_process(self) -> bool:
        """查找企业微信进程"""
        try:
            wxwork_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    proc_name = proc.info['name']
                    if any(name.lower() in proc_name.lower() for name in self.process_names):
                        wxwork_processes.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not wxwork_processes:
                logger.error("未找到企业微信进程，请先启动企业微信")
                return False
            
            # 选择第一个企业微信进程
            self.wxwork_process = wxwork_processes[0]
            self.wxwork_pid = self.wxwork_process.pid
            logger.info(f"找到企业微信进程 PID: {self.wxwork_pid}")
            return True
            
        except Exception as e:
            logger.error(f"查找企业微信进程失败: {e}")
            return False
    
    def _find_wxwork_windows(self) -> bool:
        """查找企业微信窗口"""
        try:
            if not self.wxwork_pid:
                logger.error("请先查找企业微信进程")
                return False
            
            # 枚举所有窗口
            windows_list = []
            win32gui.EnumWindows(self._enum_windows_callback, windows_list)
            
            # 查找属于企业微信进程的窗口
            wxwork_windows = [w for w in windows_list if w['pid'] == self.wxwork_pid]
            
            if not wxwork_windows:
                logger.error("未找到企业微信窗口")
                return False
            
            # 查找主窗口（基于实际窗口结构：WeWorkWindow类名，企业微信标题）
            main_windows = []
            for w in wxwork_windows:
                class_name = w['class'].lower()
                title = w['title']
                
                # 企业微信主窗口特征匹配（优先选择可见窗口）
                if (class_name == 'weworkwindow' and title == '企业微信') or \
                   (class_name == 'weworkwindow' and w.get('visible', False)):
                    main_windows.append(w)
                    logger.info(f"匹配到主窗口候选: {title} (类名: {w['class']}, 句柄: {w['hwnd']})")
            
            if main_windows:
                self.main_window_hwnd = main_windows[0]['hwnd']
                logger.info(f"找到企业微信主窗口: {main_windows[0]['title']}")
            else:
                # 备选方案：选择第一个有标题的窗口
                titled_windows = [w for w in wxwork_windows if w['title'].strip()]
                if titled_windows:
                    self.main_window_hwnd = titled_windows[0]['hwnd']
                    logger.info(f"使用备选企业微信窗口: {titled_windows[0]['title']}")
                else:
                    logger.error("无法确定企业微信主窗口")
                    return False
            
            return self.main_window_hwnd is not None
            
        except Exception as e:
            logger.error(f"查找企业微信窗口失败: {e}")
            return False
    
    def _enum_windows_callback(self, hwnd, windows_list):
        """枚举窗口回调函数（修复版）"""
        try:
            # 获取窗口所属进程ID
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            
            # 不限制可见性，先获取所有窗口
            window_text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            
            windows_list.append({
                'hwnd': hwnd,
                'title': window_text,
                'class': class_name,
                'pid': window_pid,
                'visible': win32gui.IsWindowVisible(hwnd)
            })
        except Exception as e:
            # 忽略无法访问的窗口
            pass
    
    def activate_application(self) -> bool:
        """激活企业微信窗口（强制显示并置顶）"""
        try:
            if not self.main_window_hwnd:
                logger.error("企业微信窗口句柄不存在")
                return False
            
            logger.info(f"激活企业微信窗口，句柄: {self.main_window_hwnd}")
            
            # 强制显示窗口（无论当前状态如何）
            win32gui.ShowWindow(self.main_window_hwnd, win32con.SW_SHOW)
            time.sleep(0.5)
            
            # 如果窗口被最小化，恢复它
            if win32gui.IsIconic(self.main_window_hwnd):
                win32gui.ShowWindow(self.main_window_hwnd, win32con.SW_RESTORE)
                time.sleep(0.5)
            
            # 置顶窗口
            win32gui.SetForegroundWindow(self.main_window_hwnd)
            time.sleep(1)
            
            # 确保窗口处于最前面
            win32gui.SetWindowPos(self.main_window_hwnd, win32con.HWND_TOP, 
                                 0, 0, 0, 0, 
                                 win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
            time.sleep(0.5)
            
            logger.info("企业微信窗口已激活并置顶")
            return True
            
        except Exception as e:
            logger.error(f"激活企业微信窗口失败: {e}")
            return False
    
    def search_group(self, group_name: str) -> bool:
        """搜索并进入企业微信群聊（简化版）"""
        try:
            logger.info(f"搜索企业微信群聊: {group_name}")
            
            # 激活企业微信窗口
            if not self.activate_application():
                return False
            
            # 企业微信搜索方式：Ctrl+F 或者点击搜索框
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(1)
            
            # 清空搜索框并输入群名
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            pyperclip.copy(group_name)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1.5)
            
            # 按回车进入第一个搜索结果
            pyautogui.press('enter')
            time.sleep(2)
            
            # 确保聊天窗口已加载
            time.sleep(1)
            
            logger.info(f"已进入企业微信群聊: {group_name}")
            return True
            
        except Exception as e:
            logger.error(f"搜索企业微信群聊失败: {e}")
            return False
    
    def send_message(self, message: str, target_group: str = None) -> bool:
        """发送消息到企业微信（简化版：置顶->点击下方->粘贴->发送）"""
        try:
            logger.info("准备发送消息到企业微信")
            
            # 如果指定了目标群聊，先搜索群聊
            if target_group:
                if not self.search_group(target_group):
                    logger.error(f"搜索群聊失败: {target_group}")
                    return False
            
            # 确保企业微信窗口置顶
            if not self.activate_application():
                return False
            
            # 等待界面稳定
            time.sleep(1)
            
            # 点击聊天输入框（通常在窗口下方）
            # 企业微信输入框位置相对固定，点击窗口下半部分
            window_rect = self._get_window_rect()
            if window_rect:
                # 点击窗口下方的输入区域
                input_x = window_rect['left'] + window_rect['width'] // 2
                input_y = window_rect['bottom'] - 100  # 距离底部100像素的位置
                
                logger.info(f"点击输入框位置: ({input_x}, {input_y})")
                pyautogui.click(input_x, input_y)
                time.sleep(0.5)
            else:
                # 备用方案：使用键盘导航到输入框
                pyautogui.press('tab')
                time.sleep(0.3)
            
            # 将格式化的消息复制到剪贴板
            formatted_message = self.format_report_message(message)
            pyperclip.copy(formatted_message)
            time.sleep(0.3)
            
            # 清空输入框并粘贴内容
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            
            # 发送消息（企业微信发送快捷键）
            # 企业微信常用发送方式：Enter 或 Ctrl+Enter
            pyautogui.press('enter')
            time.sleep(1)
            
            logger.info("企业微信消息发送完成")
            return True
            
        except Exception as e:
            logger.error(f"发送企业微信消息失败: {e}")
            return False
    
    def _get_window_rect(self) -> Optional[Dict[str, int]]:
        """获取企业微信窗口位置和大小"""
        try:
            if not self.main_window_hwnd:
                return None
                
            rect = win32gui.GetWindowRect(self.main_window_hwnd)
            return {
                'left': rect[0],
                'top': rect[1],
                'right': rect[2],
                'bottom': rect[3],
                'width': rect[2] - rect[0],
                'height': rect[3] - rect[1]
            }
        except Exception as e:
            logger.error(f"获取窗口位置失败: {e}")
            return None
    
    def cleanup(self) -> bool:
        """清理资源"""
        try:
            logger.info("清理企业微信发送器资源")
            self.wxwork_process = None
            self.wxwork_pid = None
            self.main_window_hwnd = None
            self.is_initialized = False
            return True
        except Exception as e:
            logger.error(f"清理资源失败: {e}")
            return False
    
    def auto_send_daily_report(self, group_name: str = None) -> bool:
        """自动发送每日报告到企业微信"""
        try:
            logger.info("开始执行企业微信自动发送每日报告")
            
            # 使用指定群名或默认群名
            target_group = group_name or self.default_group
            
            # 初始化发送器
            if not self.initialize():
                logger.error("初始化企业微信发送器失败")
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
            
            logger.info("企业微信每日报告发送成功！")
            return True
            
        except Exception as e:
            logger.error(f"企业微信自动发送每日报告失败: {e}")
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
            
            # 添加企业微信特有信息
            if self.wxwork_process:
                info["企业微信进程信息"] = {
                    "PID": self.wxwork_pid,
                    "进程名": self.wxwork_process.name(),
                    "可执行文件": getattr(self.wxwork_process, 'exe', lambda: "无法获取")()
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
    
    def _debug_show_all_windows(self):
        """调试：显示企业微信进程的所有窗口"""
        try:
            if not self.wxwork_pid:
                print("❌ 没有企业微信进程PID")
                return
            
            print(f"\n🔍 查找 PID {self.wxwork_pid} 的所有窗口:")
            print("-" * 80)
            print(f"{'序号':<4} {'句柄':<10} {'可见':<4} {'标题':<30} {'类名'}")
            print("-" * 80)
            
            # 枚举所有窗口
            windows_list = []
            win32gui.EnumWindows(self._enum_windows_callback, windows_list)
            
            # 筛选企业微信进程的窗口
            wxwork_windows = [w for w in windows_list if w['pid'] == self.wxwork_pid]
            
            if not wxwork_windows:
                print("❌ 未找到任何属于该进程的窗口")
                return
            
            for i, window in enumerate(wxwork_windows, 1):
                is_visible = "是" if window.get('visible', False) else "否"
                title = window['title'][:28] + ".." if len(window['title']) > 30 else window['title']
                title = title if title.strip() else "(无标题)"
                
                print(f"{i:<4} {window['hwnd']:<10} {is_visible:<4} {title:<30} {window['class']}")
            
            print(f"\n📊 总共找到 {len(wxwork_windows)} 个窗口")
            
            # 显示主窗口候选
            print("\n🎯 主窗口候选分析:")
            for window in wxwork_windows:
                if window.get('visible', False):  # 只分析可见窗口
                    class_name = window['class'].lower()
                    title = window['title'].lower()
                    
                    score = 0
                    reasons = []
                    
                    # 评分标准
                    if 'wechat' in class_name or 'wxwork' in class_name:
                        score += 3
                        reasons.append("类名匹配")
                    
                    if '企业微信' in title or 'wework' in title or 'work' in title:
                        score += 2
                        reasons.append("标题匹配")
                    
                    if window['title'].strip():  # 有标题
                        score += 1
                        reasons.append("有标题")
                    
                    if score > 0:
                        print(f"  候选: {window['title']} (句柄:{window['hwnd']}, 评分:{score}, 原因:{','.join(reasons)})")
            
            # 如果没有可见窗口，显示所有窗口供参考
            visible_count = sum(1 for w in wxwork_windows if w.get('visible', False))
            if visible_count == 0:
                print("\n⚠️ 没有可见窗口，显示所有窗口供参考:")
            
        except Exception as e:
            print(f"❌ 调试显示窗口失败: {e}")
    
    def _manual_select_window(self) -> bool:
        """手动选择企业微信窗口"""
        try:
            if not self.wxwork_pid:
                print("❌ 没有企业微信进程PID")
                return False
            
            # 显示所有窗口
            self._debug_show_all_windows()
            
            # 让用户输入窗口句柄
            print("\n请输入要使用的窗口句柄（从上面列表中选择）:")
            user_input = input("窗口句柄: ").strip()
            
            if not user_input:
                print("❌ 输入为空")
                return False
            
            try:
                hwnd = int(user_input)
                
                # 验证句柄是否属于企业微信进程
                windows_list = []
                win32gui.EnumWindows(self._enum_windows_callback, windows_list)
                wxwork_windows = [w for w in windows_list if w['pid'] == self.wxwork_pid]
                
                selected_window = None
                for w in wxwork_windows:
                    if w['hwnd'] == hwnd:
                        selected_window = w
                        break
                
                if not selected_window:
                    print(f"❌ 句柄 {hwnd} 不属于企业微信进程")
                    return False
                
                # 设置主窗口句柄
                self.main_window_hwnd = hwnd
                print(f"✅ 已选择窗口: {selected_window['title']} (句柄: {hwnd})")
                return True
                
            except ValueError:
                print("❌ 无效的句柄数字")
                return False
            
        except Exception as e:
            print(f"❌ 手动选择窗口失败: {e}")
            return False


# 注册企业微信发送器到工厂
MessageSenderFactory.register_sender("wxwork", WXWorkSender)


# ==================== 命令行接口 ====================
def main():
    """主程序入口"""
    import sys
    import json
    
    # 配置日志
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    sender = WXWorkSender()
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "send":
            # 发送今日报告
            group_name = sys.argv[2] if len(sys.argv) > 2 else None
            success = sender.auto_send_daily_report(group_name)
            if success:
                print("✅ 企业微信报告发送成功！")
            else:
                print("❌ 企业微信报告发送失败！")
                
        elif command == "debug":
            # 获取调试信息
            sender.initialize()
            debug_info = sender.get_debug_info()
            print("=== 企业微信调试信息 ===")
            print(json.dumps(debug_info, ensure_ascii=False, indent=2))
            
        elif command == "test":
            # 测试功能
            print("测试企业微信进程查找...")
            if sender.find_target_process():
                print("✅ 企业微信进程查找成功")
                
                print("测试企业微信窗口查找...")
                if sender._find_wxwork_windows():
                    print("✅ 企业微信窗口查找成功")
                    
                    print("测试窗口激活...")
                    if sender.activate_application():
                        print("✅ 窗口激活成功")
                    else:
                        print("❌ 窗口激活失败")
                else:
                    print("❌ 企业微信窗口查找失败")
                    print("正在显示所有相关窗口信息...")
                    sender._debug_show_all_windows()
            else:
                print("❌ 企业微信进程查找失败")
                
        elif command == "windows":
            # 显示所有窗口
            print("查找企业微信进程...")
            if sender.find_target_process():
                print(f"找到企业微信进程 PID: {sender.wxwork_pid}")
                sender._debug_show_all_windows()
            else:
                print("❌ 未找到企业微信进程")
                
        elif command == "manual":
            # 手动选择窗口
            print("查找企业微信进程...")
            if sender.find_target_process():
                print(f"找到企业微信进程 PID: {sender.wxwork_pid}")
                if sender._manual_select_window():
                    print("✅ 手动选择窗口成功，测试激活...")
                    if sender.activate_application():
                        print("✅ 窗口激活成功")
                    else:
                        print("❌ 窗口激活失败")
                else:
                    print("❌ 手动选择窗口失败")
            else:
                print("❌ 未找到企业微信进程")
                
        else:
            print("未知命令。可用命令：")
            print("  send [群名] - 发送今日报告到企业微信")
            print("  debug - 获取调试信息")
            print("  test - 测试功能")
            print("  windows - 显示所有企业微信窗口")
            print("  manual - 手动选择窗口并测试")
    else:
        print("企业微信自动发送工具 v1.0")
        print("用法:")
        print("  python wxwork_sender.py send [群名] - 发送今日报告")
        print("  python wxwork_sender.py debug - 获取调试信息") 
        print("  python wxwork_sender.py test - 测试功能")


if __name__ == "__main__":
    main()