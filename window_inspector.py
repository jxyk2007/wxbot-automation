# -*- coding: utf-8 -*-
"""
窗口检查器 - 点击窗口获取信息
版本：v1.0.0
创建日期：2025-09-10
功能：点击任意窗口获取详细信息，帮助识别微信主窗口
"""

import win32gui
import win32process
import win32api
import win32con
import psutil
import time
import threading
from datetime import datetime

class WindowInspector:
    def __init__(self):
        self.is_listening = False
        self.hook = None
    
    def get_window_at_cursor(self):
        """获取鼠标位置的窗口信息"""
        try:
            # 获取鼠标位置
            cursor_pos = win32gui.GetCursorPos()
            
            # 获取该位置的窗口句柄
            hwnd = win32gui.WindowFromPoint(cursor_pos)
            
            if hwnd:
                return self.get_window_detailed_info(hwnd)
            else:
                return None
                
        except Exception as e:
            print(f"获取窗口信息失败: {e}")
            return None
    
    def get_window_detailed_info(self, hwnd):
        """获取窗口详细信息"""
        try:
            info = {}
            
            # 基本窗口信息
            info['窗口句柄'] = hwnd
            info['窗口标题'] = win32gui.GetWindowText(hwnd)
            info['窗口类名'] = win32gui.GetClassName(hwnd)
            info['是否可见'] = win32gui.IsWindowVisible(hwnd)
            info['是否启用'] = win32gui.IsWindowEnabled(hwnd)
            
            # 窗口位置和大小
            rect = win32gui.GetWindowRect(hwnd)
            info['窗口位置'] = {
                'left': rect[0],
                'top': rect[1],
                'right': rect[2],
                'bottom': rect[3],
                'width': rect[2] - rect[0],
                'height': rect[3] - rect[1]
            }
            
            # 进程信息
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                info['进程ID'] = pid
                
                # 获取进程详细信息
                process = psutil.Process(pid)
                info['进程名'] = process.name()
                info['进程路径'] = process.exe() if hasattr(process, 'exe') else '无法获取'
                info['内存使用'] = f"{round(process.memory_info().rss / 1024 / 1024, 1)}MB"
                
            except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                info['进程信息错误'] = str(e)
            
            # 父窗口信息
            try:
                parent_hwnd = win32gui.GetParent(hwnd)
                if parent_hwnd:
                    info['父窗口'] = {
                        '句柄': parent_hwnd,
                        '标题': win32gui.GetWindowText(parent_hwnd),
                        '类名': win32gui.GetClassName(parent_hwnd)
                    }
                else:
                    info['父窗口'] = None
            except:
                info['父窗口'] = '获取失败'
            
            # 子窗口数量
            try:
                child_windows = []
                def enum_child_proc(child_hwnd, param):
                    child_windows.append(child_hwnd)
                
                win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                info['子窗口数量'] = len(child_windows)
            except:
                info['子窗口数量'] = '获取失败'
                
            return info
            
        except Exception as e:
            return {'错误': str(e)}
    
    def print_window_info(self, info):
        """格式化打印窗口信息"""
        if not info:
            print("❌ 未获取到窗口信息")
            return
            
        print("\n" + "="*60)
        print(f"🔍 窗口信息 - {datetime.now().strftime('%H:%M:%S')}")
        print("="*60)
        
        for key, value in info.items():
            if key == '窗口位置':
                print(f"{key}:")
                for pos_key, pos_value in value.items():
                    print(f"  {pos_key}: {pos_value}")
            elif key == '父窗口' and isinstance(value, dict):
                print(f"{key}:")
                for parent_key, parent_value in value.items():
                    print(f"  {parent_key}: {parent_value}")
            else:
                print(f"{key}: {value}")
        
        print("="*60)
    
    def click_to_inspect(self):
        """点击窗口获取信息模式"""
        print("\n🔍 点击窗口检查模式")
        print("=" * 50)
        print("操作说明：")
        print("1. 将鼠标移到要检查的窗口上")
        print("2. 按 Ctrl+Alt+I 获取窗口信息")
        print("3. 按 Ctrl+C 退出")
        print("=" * 50)
        
        try:
            while True:
                # 检查快捷键
                if (win32api.GetAsyncKeyState(win32con.VK_CONTROL) & 0x8000 and 
                    win32api.GetAsyncKeyState(win32con.VK_MENU) & 0x8000 and 
                    win32api.GetAsyncKeyState(ord('I')) & 0x8000):
                    
                    print("\n🎯 正在获取当前鼠标位置的窗口信息...")
                    info = self.get_window_at_cursor()
                    self.print_window_info(info)
                    
                    # 等待按键释放
                    while (win32api.GetAsyncKeyState(ord('I')) & 0x8000):
                        time.sleep(0.1)
                
                # 检查退出快捷键
                if (win32api.GetAsyncKeyState(win32con.VK_CONTROL) & 0x8000 and 
                    win32api.GetAsyncKeyState(ord('C')) & 0x8000):
                    print("\n👋 退出窗口检查模式")
                    break
                    
                time.sleep(0.1)
                
        except KeyboardInterrupt:
            print("\n👋 退出窗口检查模式")
    
    def find_main_wechat_process(self):
        """智能查找微信主进程"""
        try:
            print("\n🔍 智能查找微信主进程...")
            
            wechat_processes = []
            
            # 查找所有可能的微信进程
            for proc in psutil.process_iter(['pid', 'name', 'exe', 'memory_info']):
                try:
                    name = proc.info['name'].lower()
                    exe = str(proc.info.get('exe', '')).lower()
                    
                    # 识别主微信进程的特征
                    is_main_wechat = False
                    process_type = ""
                    
                    if name in ['wechat.exe', 'weixin.exe'] and 'wechatappex' not in exe:
                        is_main_wechat = True
                        process_type = "主程序"
                    elif 'wechatupdate' in name:
                        process_type = "更新程序"
                    elif 'wechatappex.exe' == name:
                        if 'applet' in exe:
                            process_type = "小程序引擎"
                        else:
                            process_type = "扩展进程"
                    elif 'wechatweb.exe' == name:
                        process_type = "网页进程"
                    elif 'wechat' in name or 'weixin' in name:
                        process_type = "其他微信进程"
                    else:
                        continue
                    
                    wechat_processes.append({
                        'pid': proc.info['pid'],
                        'name': proc.info['name'],
                        'exe': proc.info.get('exe', '无法获取'),
                        'memory_mb': round(proc.info['memory_info'].rss / 1024 / 1024, 1),
                        'type': process_type,
                        'is_main': is_main_wechat
                    })
                    
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not wechat_processes:
                print("❌ 未找到微信相关进程")
                return None
            
            # 按类型排序，主程序优先
            wechat_processes.sort(key=lambda x: (not x['is_main'], x['memory_mb']), reverse=True)
            
            print(f"\n找到 {len(wechat_processes)} 个微信相关进程:")
            print("-" * 90)
            print(f"{'序号':<4} {'PID':<8} {'进程名':<16} {'类型':<12} {'内存(MB)':<10} {'路径'}")
            print("-" * 90)
            
            for i, proc in enumerate(wechat_processes, 1):
                exe_short = proc['exe']
                if len(exe_short) > 30:
                    exe_short = "..." + exe_short[-27:]
                    
                marker = "👑" if proc['is_main'] else "  "
                print(f"{marker}{i:<3} {proc['pid']:<8} {proc['name']:<16} {proc['type']:<12} {proc['memory_mb']:<10} {exe_short}")
            
            # 自动推荐主进程
            main_processes = [p for p in wechat_processes if p['is_main']]
            if main_processes:
                recommended = main_processes[0]
                print(f"\n💡 推荐使用: {recommended['name']} (PID: {recommended['pid']}) - {recommended['type']}")
                return recommended['pid']
            else:
                print(f"\n⚠️ 未找到明确的主进程，建议选择内存占用最大的进程")
                return wechat_processes[0]['pid'] if wechat_processes else None
                
        except Exception as e:
            print(f"查找微信主进程失败: {e}")
            return None
    
    def show_windows_for_pid(self, pid):
        """显示指定PID的所有窗口"""
        try:
            print(f"\n🪟 查找PID {pid} 的所有窗口...")
            
            windows = []
            def enum_windows_callback(hwnd, param):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                        if window_pid == pid:
                            windows.append({
                                'hwnd': hwnd,
                                'title': win32gui.GetWindowText(hwnd),
                                'class': win32gui.GetClassName(hwnd)
                            })
                    except:
                        pass
            
            win32gui.EnumWindows(enum_windows_callback, None)
            
            if not windows:
                print("❌ 该进程没有可见窗口")
                return None
                
            print(f"\n找到 {len(windows)} 个窗口:")
            print("-" * 80)
            print(f"{'序号':<4} {'句柄':<10} {'窗口标题':<30} {'窗口类名'}")
            print("-" * 80)
            
            for i, window in enumerate(windows, 1):
                title = window['title'] if window['title'].strip() else "(无标题)"
                title = title[:28] + ".." if len(title) > 30 else title
                class_name = window['class'][:28] + ".." if len(window['class']) > 30 else window['class']
                
                # 标记可能是主窗口的特征
                marker = "🏠" if 'WeChat' in window['class'] and 'Main' in window['class'] else "  "
                print(f"{marker}{i:<3} {window['hwnd']:<10} {title:<30} {class_name}")
            
            return windows
            
        except Exception as e:
            print(f"显示窗口失败: {e}")
            return None

def main():
    """主程序入口"""
    import sys
    
    inspector = WindowInspector()
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "click":
            # 点击检查模式
            inspector.click_to_inspect()
            
        elif command == "findwechat":
            # 智能查找微信
            main_pid = inspector.find_main_wechat_process()
            if main_pid:
                inspector.show_windows_for_pid(main_pid)
                
        elif command == "pid":
            # 查看指定PID的窗口
            if len(sys.argv) > 2:
                try:
                    pid = int(sys.argv[2])
                    inspector.show_windows_for_pid(pid)
                except ValueError:
                    print("请输入有效的PID")
            else:
                print("请指定PID: python window_inspector.py pid <PID>")
                
        else:
            print("未知命令")
    else:
        print("🔍 窗口检查器")
        print("用法:")
        print("  python window_inspector.py click      - 点击窗口检查模式")
        print("  python window_inspector.py findwechat - 智能查找微信主进程")
        print("  python window_inspector.py pid <PID>  - 查看指定PID的窗口")

if __name__ == "__main__":
    main()