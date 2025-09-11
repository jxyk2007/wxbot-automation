# -*- coding: utf-8 -*-
"""
微信自动发送模块 - 增强版
版本：v2.0.0
创建日期：2025-09-10
功能：使用窗口句柄和控件定位，更稳定的微信自动发送
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

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WeChatSenderV2:
    def __init__(self):
        """初始化微信发送器"""
        # 设置pyautogui安全配置
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # 微信进程和窗口信息
        self.wechat_process = None
        self.wechat_pid = None
        self.main_window_hwnd = None
        self.chat_window_hwnd = None
        
        # 群名称
        self.group_name = "存储统计报告群"
    
    def list_all_processes(self, filter_keyword=""):
        """列出所有进程（可选过滤）"""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe', 'memory_info']):
                try:
                    proc_info = {
                        'pid': proc.info['pid'],
                        'name': proc.info['name'],
                        'exe': proc.info.get('exe', '无法获取'),
                        'memory_mb': round(proc.info['memory_info'].rss / 1024 / 1024, 1) if proc.info.get('memory_info') else 0
                    }
                    
                    # 过滤条件
                    if filter_keyword:
                        if (filter_keyword.lower() in proc_info['name'].lower() or 
                            filter_keyword.lower() in str(proc_info['exe']).lower()):
                            processes.append(proc_info)
                    else:
                        processes.append(proc_info)
                        
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
            
            # 按内存使用量排序
            processes.sort(key=lambda x: x['memory_mb'], reverse=True)
            return processes
            
        except Exception as e:
            logger.error(f"列出进程失败: {e}")
            return []
    
    def interactive_select_process(self):
        """交互式选择进程"""
        try:
            print("\n=== 进程选择器 ===")
            print("1. 自动查找微信相关进程")
            print("2. 手动输入关键词搜索")
            print("3. 显示所有进程")
            print("4. 直接输入PID")
            
            choice = input("\n请选择操作 (1-4): ").strip()
            
            # 检查是否直接输入了数字PID（大于4的数字）
            if choice.isdigit() and int(choice) > 4:
                try:
                    pid = int(choice)
                    proc = psutil.Process(pid)
                    self.wechat_process = proc
                    self.wechat_pid = pid
                    print(f"已选择进程: {proc.name()} (PID: {pid})")
                    return True
                except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                    print(f"无效的PID {pid}: {e}")
                    return False
            
            if choice == "1":
                # 自动查找微信（包含weixin和wechat）
                wechat_procs = self.list_all_processes("wechat")
                weixin_procs = self.list_all_processes("weixin")
                # 合并并去重
                all_procs = {p['pid']: p for p in wechat_procs + weixin_procs}
                processes = list(all_procs.values())
                if not processes:
                    print("未找到微信相关进程")
                    return False
                    
            elif choice == "2":
                # 关键词搜索
                keyword = input("请输入搜索关键词: ").strip()
                if not keyword:
                    print("关键词不能为空")
                    return False
                processes = self.list_all_processes(keyword)
                
            elif choice == "3":
                # 显示所有进程（限制前50个）
                print("正在获取进程列表...")
                all_processes = self.list_all_processes()
                processes = all_processes[:50]  # 限制显示数量
                print(f"显示前50个进程（共{len(all_processes)}个）")
                
            elif choice == "4":
                # 直接输入PID
                try:
                    pid = int(input("请输入进程PID: ").strip())
                    proc = psutil.Process(pid)
                    self.wechat_process = proc
                    self.wechat_pid = pid
                    print(f"已选择进程: {proc.name()} (PID: {pid})")
                    return True
                except (ValueError, psutil.NoSuchProcess) as e:
                    print(f"无效的PID: {e}")
                    return False
            else:
                print("无效选择")
                return False
            
            if not processes:
                print("未找到匹配的进程")
                return False
            
            # 显示进程列表
            print(f"\n找到 {len(processes)} 个进程:")
            print("-" * 80)
            print(f"{'序号':<4} {'PID':<8} {'进程名':<20} {'内存(MB)':<10} {'可执行文件'}")
            print("-" * 80)
            
            for i, proc in enumerate(processes, 1):
                exe_short = proc['exe']
                if len(exe_short) > 40:
                    exe_short = "..." + exe_short[-37:]
                    
                print(f"{i:<4} {proc['pid']:<8} {proc['name']:<20} {proc['memory_mb']:<10} {exe_short}")
            
            # 让用户选择
            try:
                selection = input(f"\n请选择进程序号 (1-{len(processes)}): ").strip()
                if not selection:
                    return False
                    
                idx = int(selection) - 1
                if 0 <= idx < len(processes):
                    selected_proc = processes[idx]
                    self.wechat_process = psutil.Process(selected_proc['pid'])
                    self.wechat_pid = selected_proc['pid']
                    print(f"已选择进程: {selected_proc['name']} (PID: {selected_proc['pid']})")
                    return True
                else:
                    print("无效的序号")
                    return False
                    
            except ValueError:
                print("请输入有效数字")
                return False
                
        except Exception as e:
            logger.error(f"交互式选择进程失败: {e}")
            return False
    
    def find_wechat_process(self):
        """查找微信进程"""
        try:
            wechat_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    if 'WeChat.exe' in proc.info['name'] or 'wechat.exe' in proc.info['name'].lower():
                        wechat_processes.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not wechat_processes:
                logger.error("未找到微信进程，请先启动微信")
                return False
            
            # 选择第一个微信进程
            self.wechat_process = wechat_processes[0]
            self.wechat_pid = self.wechat_process.pid
            logger.info(f"找到微信进程 PID: {self.wechat_pid}")
            return True
            
        except Exception as e:
            logger.error(f"查找微信进程失败: {e}")
            return False
    
    def enum_windows_callback(self, hwnd, windows_list):
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
    
    def interactive_select_window(self):
        """交互式选择窗口"""
        try:
            if not self.wechat_pid:
                print("请先选择进程")
                return False
            
            print(f"\n=== 窗口选择器 (PID: {self.wechat_pid}) ===")
            
            # 枚举所有窗口
            windows_list = []
            win32gui.EnumWindows(self.enum_windows_callback, windows_list)
            
            # 查找属于选定进程的窗口
            process_windows = [w for w in windows_list if w['pid'] == self.wechat_pid]
            
            if not process_windows:
                print("该进程没有可见窗口")
                return False
            
            # 显示窗口列表
            print(f"\n找到 {len(process_windows)} 个窗口:")
            print("-" * 100)
            print(f"{'序号':<4} {'窗口句柄':<10} {'窗口标题':<30} {'窗口类名':<30}")
            print("-" * 100)
            
            for i, window in enumerate(process_windows, 1):
                title = window['title'] if window['title'].strip() else "(无标题)"
                title = title[:28] + ".." if len(title) > 30 else title
                class_name = window['class'][:28] + ".." if len(window['class']) > 30 else window['class']
                
                print(f"{i:<4} {window['hwnd']:<10} {title:<30} {class_name:<30}")
            
            # 让用户选择
            try:
                selection = input(f"\n请选择窗口序号 (1-{len(process_windows)}): ").strip()
                if not selection:
                    return False
                    
                idx = int(selection) - 1
                if 0 <= idx < len(process_windows):
                    selected_window = process_windows[idx]
                    self.main_window_hwnd = selected_window['hwnd']
                    print(f"已选择窗口: {selected_window['title']} (句柄: {selected_window['hwnd']})")
                    return True
                else:
                    print("无效的序号")
                    return False
                    
            except ValueError:
                print("请输入有效数字")
                return False
                
        except Exception as e:
            logger.error(f"交互式选择窗口失败: {e}")
            return False
    
    def find_wechat_windows(self):
        """查找微信窗口"""
        try:
            if not self.wechat_pid:
                logger.error("请先查找微信进程")
                return False
            
            # 枚举所有窗口
            windows_list = []
            win32gui.EnumWindows(self.enum_windows_callback, windows_list)
            
            # 查找属于微信进程的窗口
            wechat_windows = [w for w in windows_list if w['pid'] == self.wechat_pid]
            
            if not wechat_windows:
                logger.error("未找到微信窗口")
                return False
            
            # 查找主窗口（通常类名包含WeChatMainWndForPC）
            main_windows = [w for w in wechat_windows if 'WeChatMainWndForPC' in w['class']]
            if main_windows:
                self.main_window_hwnd = main_windows[0]['hwnd']
                logger.info(f"找到微信主窗口: {main_windows[0]['title']}")
            else:
                # 备选方案：选择第一个有标题的窗口
                titled_windows = [w for w in wechat_windows if w['title'].strip()]
                if titled_windows:
                    self.main_window_hwnd = titled_windows[0]['hwnd']
                    logger.info(f"使用备选微信窗口: {titled_windows[0]['title']}")
            
            return self.main_window_hwnd is not None
            
        except Exception as e:
            logger.error(f"查找微信窗口失败: {e}")
            return False
    
    def activate_wechat_window(self):
        """激活微信窗口"""
        try:
            if not self.main_window_hwnd:
                logger.error("微信窗口句柄不存在")
                return False
            
            # 检查窗口是否最小化
            if win32gui.IsIconic(self.main_window_hwnd):
                win32gui.ShowWindow(self.main_window_hwnd, win32con.SW_RESTORE)
            
            # 激活窗口
            win32gui.SetForegroundWindow(self.main_window_hwnd)
            time.sleep(1)
            
            logger.info("微信窗口已激活")
            return True
            
        except Exception as e:
            logger.error(f"激活微信窗口失败: {e}")
            return False
    
    def get_window_rect(self, hwnd):
        """获取窗口位置和大小"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
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
    
    def find_control_in_window(self, class_name, window_text=""):
        """在微信窗口中查找控件"""
        try:
            def enum_child_proc(hwnd, param):
                class_name_target, window_text_target, result_list = param
                child_class = win32gui.GetClassName(hwnd)
                child_text = win32gui.GetWindowText(hwnd)
                
                if class_name_target in child_class:
                    if not window_text_target or window_text_target in child_text:
                        rect = win32gui.GetWindowRect(hwnd)
                        result_list.append({
                            'hwnd': hwnd,
                            'class': child_class,
                            'text': child_text,
                            'rect': rect
                        })
            
            result_list = []
            win32gui.EnumChildWindows(
                self.main_window_hwnd, 
                enum_child_proc, 
                (class_name, window_text, result_list)
            )
            
            return result_list
            
        except Exception as e:
            logger.error(f"查找控件失败: {e}")
            return []
    
    def smart_search_group(self, group_name):
        """智能搜索群聊"""
        try:
            logger.info(f"智能搜索群聊: {group_name}")
            
            # 激活微信窗口
            if not self.activate_wechat_window():
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
            
            logger.info(f"已搜索并进入群聊: {group_name}")
            return True
            
        except Exception as e:
            logger.error(f"智能搜索群聊失败: {e}")
            return False
    
    def send_message_to_current_chat(self, message):
        """发送消息到当前聊天"""
        try:
            logger.info("准备发送消息到当前聊天")
            
            # 确保微信窗口处于前台
            if not self.activate_wechat_window():
                return False
            
            # 将消息复制到剪贴板
            pyperclip.copy(message)
            time.sleep(0.2)
            
            # 粘贴消息
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            
            # 发送消息（Alt+S 或 Ctrl+Enter）
            pyautogui.hotkey('alt', 's')
            time.sleep(1)
            
            logger.info("消息发送完成")
            return True
            
        except Exception as e:
            logger.error(f"发送消息失败: {e}")
            return False
    
    def get_debug_info(self):
        """获取调试信息"""
        try:
            info = {
                "微信进程信息": {},
                "窗口信息": {},
                "系统信息": {}
            }
            
            # 微信进程信息
            if self.wechat_process:
                info["微信进程信息"] = {
                    "PID": self.wechat_pid,
                    "进程名": self.wechat_process.name(),
                    "可执行文件": self.wechat_process.exe() if hasattr(self.wechat_process, 'exe') else "无法获取"
                }
            
            # 窗口信息
            if self.main_window_hwnd:
                rect = self.get_window_rect(self.main_window_hwnd)
                info["窗口信息"] = {
                    "窗口句柄": self.main_window_hwnd,
                    "窗口标题": win32gui.GetWindowText(self.main_window_hwnd),
                    "窗口类名": win32gui.GetClassName(self.main_window_hwnd),
                    "窗口位置": rect
                }
            
            # 系统信息
            info["系统信息"] = {
                "屏幕分辨率": pyautogui.size(),
                "当前时间": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            return info
            
        except Exception as e:
            logger.error(f"获取调试信息失败: {e}")
            return {"错误": str(e)}
    
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

🤖 此消息由自动化系统发送
💻 系统版本: WeChatSender v2.0"""
            
            return formatted_content
            
        except Exception as e:
            logger.error(f"格式化报告内容失败: {e}")
            return content
    
    def auto_send_daily_report(self, group_name=None):
        """自动发送每日报告 - 增强版"""
        try:
            logger.info("开始执行自动发送每日报告 (增强版)")
            
            # 使用指定群名或默认群名
            target_group = group_name or self.group_name
            
            # 1. 查找微信进程
            if not self.find_wechat_process():
                return False
            
            # 2. 查找微信窗口
            if not self.find_wechat_windows():
                return False
            
            # 3. 查找最新报告文件
            report_file = self.find_latest_report()
            if not report_file:
                logger.error("未找到报告文件，发送失败")
                return False
            
            # 4. 读取报告内容
            content = self.read_report_content(report_file)
            if not content:
                logger.error("读取报告内容失败，发送失败")
                return False
            
            # 5. 智能搜索并进入群聊
            if not self.smart_search_group(target_group):
                logger.error("搜索群聊失败，发送失败")
                return False
            
            # 6. 格式化并发送消息
            formatted_content = self.format_report_for_wechat(content)
            if not self.send_message_to_current_chat(formatted_content):
                logger.error("发送消息失败")
                return False
            
            logger.info("每日报告发送成功！(增强版)")
            return True
            
        except Exception as e:
            logger.error(f"自动发送每日报告失败: {e}")
            return False

# ==================== 命令行接口 ====================
def main():
    """主程序入口"""
    import sys
    import json
    
    sender = WeChatSenderV2()
    
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
                
        elif command == "debug":
            # 获取调试信息
            sender.find_wechat_process()
            sender.find_wechat_windows()
            debug_info = sender.get_debug_info()
            print("=== 调试信息 ===")
            print(json.dumps(debug_info, ensure_ascii=False, indent=2))
            
        elif command == "select":
            # 交互式选择进程和窗口
            print("=== 交互式配置模式 ===")
            if sender.interactive_select_process():
                if sender.interactive_select_window():
                    print("\n✅ 配置完成！")
                    
                    # 测试窗口激活
                    print("测试窗口激活...")
                    if sender.activate_wechat_window():
                        print("✅ 窗口激活成功")
                        print("\n现在可以使用 send 命令发送报告了")
                    else:
                        print("❌ 窗口激活失败")
                else:
                    print("❌ 窗口选择失败")
            else:
                print("❌ 进程选择失败")
        
        elif command == "listproc":
            # 列出进程
            keyword = sys.argv[2] if len(sys.argv) > 2 else ""
            processes = sender.list_all_processes(keyword)
            
            if processes:
                print(f"\n找到 {len(processes)} 个进程:")
                print("-" * 80)
                print(f"{'PID':<8} {'进程名':<20} {'内存(MB)':<10} {'可执行文件'}")
                print("-" * 80)
                
                for proc in processes[:20]:  # 限制显示前20个
                    exe_short = proc['exe']
                    if len(exe_short) > 40:
                        exe_short = "..." + exe_short[-37:]
                    print(f"{proc['pid']:<8} {proc['name']:<20} {proc['memory_mb']:<10} {exe_short}")
                
                if len(processes) > 20:
                    print(f"\n... 还有 {len(processes) - 20} 个进程")
            else:
                print("未找到匹配的进程")
                
        elif command == "test":
            # 测试功能
            print("测试微信进程查找...")
            if sender.find_wechat_process():
                print("✅ 微信进程查找成功")
                
                print("测试微信窗口查找...")
                if sender.find_wechat_windows():
                    print("✅ 微信窗口查找成功")
                    
                    print("测试窗口激活...")
                    if sender.activate_wechat_window():
                        print("✅ 窗口激活成功")
                    else:
                        print("❌ 窗口激活失败")
                else:
                    print("❌ 微信窗口查找失败")
            else:
                print("❌ 微信进程查找失败")
                
        else:
            print("未知命令。可用命令：")
            print("  send [群名] - 发送今日报告")
            print("  select - 交互式选择进程和窗口")
            print("  listproc [关键词] - 列出进程")
            print("  debug - 获取调试信息")
            print("  test - 测试功能")
    else:
        print("微信自动发送工具 v2.0 (增强版)")
        print("用法:")
        print("  python wechat_sender_v2.py send [群名] - 发送今日报告")
        print("  python wechat_sender_v2.py select - 交互式选择进程和窗口")
        print("  python wechat_sender_v2.py listproc [关键词] - 列出进程")
        print("  python wechat_sender_v2.py debug - 获取调试信息") 
        print("  python wechat_sender_v2.py test - 测试功能")
        print("\n首次使用建议：")
        print("  1. python wechat_sender_v2.py listproc wechat  # 查看微信进程")
        print("  2. python wechat_sender_v2.py select          # 交互选择")
        print("  3. python wechat_sender_v2.py send           # 发送报告")

if __name__ == "__main__":
    main()