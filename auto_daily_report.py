# -*- coding: utf-8 -*-
"""
每日存储统计自动化系统
版本：v1.0.0
创建日期：2025-09-10
功能：完全自动化的存储统计+微信发送系统
"""

import subprocess
import sys
import os
import time
import json
import logging
from datetime import datetime
import win32gui
import win32process
import psutil

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

class AutoReportSystem:
    def __init__(self):
        """初始化自动化系统"""
        self.config_file = "auto_report_config.json"
        self.config = self.load_config()
        
    def load_config(self):
        """加载配置文件"""
        default_config = {
            "target_groups": [
                {
                    "name": "交付运维日报群",
                    "hwnd": None,
                    "enabled": True
                },
                {
                    "name": "AI TESt",
                    "hwnd": None,
                    "enabled": True
                }
            ],
            "wechat_process_name": "Weixin.exe",
            "auto_find_windows": True,
            "backup_send_enabled": False
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                # 合并默认配置（确保所有必需字段都存在）
                for key in default_config:
                    if key not in config:
                        config[key] = default_config[key]
                return config
            else:
                # 创建默认配置文件
                self.save_config(default_config)
                return default_config
        except Exception as e:
            logger.error(f"加载配置文件失败，使用默认配置: {e}")
            return default_config
    
    def save_config(self, config=None):
        """保存配置文件"""
        try:
            config_to_save = config or self.config
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, ensure_ascii=False, indent=2)
            logger.info(f"配置已保存到: {self.config_file}")
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")
    
    def find_wechat_main_process(self):
        """查找微信主进程（选择内存占用最大的主进程）"""
        try:
            logger.info("查找微信主进程...")
            
            main_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe', 'memory_info']):
                try:
                    if proc.info['name'].lower() in ['weixin.exe', 'wechat.exe']:
                        exe_path = str(proc.info.get('exe', '')).lower()
                        # 排除小程序进程
                        if 'wechatappex' not in exe_path:
                            main_processes.append({
                                'pid': proc.info['pid'],
                                'name': proc.info['name'],
                                'memory_mb': round(proc.info['memory_info'].rss / 1024 / 1024, 1)
                            })
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not main_processes:
                logger.error("未找到微信主进程")
                return None
            
            # 按内存使用量排序，选择最大的（通常是主进程）
            main_processes.sort(key=lambda x: x['memory_mb'], reverse=True)
            
            selected_process = main_processes[0]
            logger.info(f"找到 {len(main_processes)} 个主进程，选择内存最大的: {selected_process['name']} (PID: {selected_process['pid']}, 内存: {selected_process['memory_mb']}MB)")
            
            # 输出所有找到的主进程用于调试
            for i, proc in enumerate(main_processes):
                marker = "👑" if i == 0 else "  "
                logger.info(f"{marker} PID {proc['pid']}: {proc['name']} ({proc['memory_mb']}MB)")
            
            return selected_process['pid']
            
        except Exception as e:
            logger.error(f"查找微信进程失败: {e}")
            return None
    
    def find_wechat_windows(self, pid):
        """查找微信进程的所有窗口"""
        try:
            logger.info(f"查找PID {pid} 的微信窗口...")
            
            windows = []
            def enum_windows_callback(hwnd, param):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                        if window_pid == pid:
                            title = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            windows.append({
                                'hwnd': hwnd,
                                'title': title,
                                'class': class_name
                            })
                    except:
                        pass
            
            win32gui.EnumWindows(enum_windows_callback, None)
            
            # 过滤出可能的群聊窗口（排除主窗口"微信"）
            chat_windows = [w for w in windows if w['title'] and w['title'] != '微信']
            
            logger.info(f"找到 {len(chat_windows)} 个聊天窗口:")
            for window in chat_windows:
                logger.info(f"  - {window['title']} (句柄: {window['hwnd']})")
            
            return chat_windows
            
        except Exception as e:
            logger.error(f"查找微信窗口失败: {e}")
            return []
    
    def update_target_windows(self):
        """更新目标群聊窗口句柄"""
        try:
            logger.info("🔍 自动更新目标群聊窗口句柄...")
            
            # 查找微信主进程
            wechat_pid = self.find_wechat_main_process()
            if not wechat_pid:
                return False
            
            # 查找所有聊天窗口
            chat_windows = self.find_wechat_windows(wechat_pid)
            if not chat_windows:
                logger.error("未找到任何聊天窗口")
                return False
            
            # 更新配置中的群聊句柄
            updated = False
            for group_config in self.config['target_groups']:
                if not group_config.get('enabled', True):
                    continue
                    
                group_name = group_config['name']
                
                # 查找匹配的窗口（精确匹配或包含匹配）
                matching_window = None
                for window in chat_windows:
                    if window['title'] == group_name:
                        matching_window = window
                        break
                    elif group_name in window['title'] or window['title'] in group_name:
                        matching_window = window
                
                if matching_window:
                    old_hwnd = group_config.get('hwnd')
                    new_hwnd = matching_window['hwnd']
                    
                    if old_hwnd != new_hwnd:
                        group_config['hwnd'] = new_hwnd
                        logger.info(f"✅ 更新群聊 '{group_name}' 句柄: {old_hwnd} → {new_hwnd}")
                        updated = True
                    else:
                        logger.info(f"✅ 群聊 '{group_name}' 句柄无需更新: {new_hwnd}")
                else:
                    logger.warning(f"⚠️ 未找到群聊 '{group_name}' 对应的窗口")
            
            if updated:
                self.save_config()
            
            return True
            
        except Exception as e:
            logger.error(f"更新目标窗口失败: {e}")
            return False
    
    def run_storage_statistics(self):
        """执行存储统计"""
        try:
            logger.info("🔄 开始执行存储统计...")
            
            # 执行存储统计脚本
            result = subprocess.run([
                sys.executable, 'storage_system.py', 'daily'
            ], capture_output=True, text=True, encoding='gbk', errors='replace')
            
            if result.returncode == 0:
                logger.info("✅ 存储统计执行成功")
                logger.info(f"输出: {result.stdout}")
                return True
            else:
                logger.error(f"❌ 存储统计执行失败")
                logger.error(f"错误输出: {result.stderr}")
                return False
                
        except Exception as e:
            logger.error(f"执行存储统计失败: {e}")
            return False
    
    def send_reports_to_groups(self):
        """发送报告到所有配置的群聊"""
        try:
            logger.info("📤 开始发送报告到群聊...")
            
            success_count = 0
            total_count = 0
            
            for group_config in self.config['target_groups']:
                if not group_config.get('enabled', True):
                    logger.info(f"跳过禁用的群聊: {group_config['name']}")
                    continue
                
                group_name = group_config['name']
                hwnd = group_config.get('hwnd')
                total_count += 1
                
                if not hwnd:
                    logger.warning(f"⚠️ 群聊 '{group_name}' 没有有效的窗口句柄")
                    continue
                
                logger.info(f"📨 发送报告到群聊: {group_name} (句柄: {hwnd})")
                
                try:
                    # 执行发送命令
                    result = subprocess.run([
                        sys.executable, 'direct_sender.py', 'send', str(hwnd)
                    ], capture_output=True, text=True, encoding='gbk', errors='replace')
                    
                    if result.returncode == 0:
                        logger.info(f"✅ 成功发送到群聊: {group_name}")
                        success_count += 1
                    else:
                        logger.error(f"❌ 发送到群聊 '{group_name}' 失败")
                        logger.error(f"错误输出: {result.stderr}")
                        
                except Exception as e:
                    logger.error(f"发送到群聊 '{group_name}' 时出错: {e}")
                
                # 发送间隔
                time.sleep(2)
            
            logger.info(f"📊 发送完成: {success_count}/{total_count} 个群聊发送成功")
            return success_count > 0
            
        except Exception as e:
            logger.error(f"发送报告失败: {e}")
            return False
    
    def run_full_automation(self):
        """运行完整的自动化流程"""
        try:
            logger.info("🚀 开始执行完整自动化流程")
            logger.info("=" * 60)
            
            # 步骤1: 自动更新窗口句柄
            if self.config.get('auto_find_windows', True):
                if not self.update_target_windows():
                    logger.error("❌ 更新窗口句柄失败，继续使用现有配置")
            
            # 步骤2: 执行存储统计
            if not self.run_storage_statistics():
                logger.error("❌ 存储统计失败，终止流程")
                return False
            
            # 步骤3: 发送报告到群聊
            if not self.send_reports_to_groups():
                logger.error("❌ 所有群聊发送都失败")
                return False
            
            logger.info("=" * 60)
            logger.info("🎉 自动化流程执行完成！")
            return True
            
        except Exception as e:
            logger.error(f"自动化流程执行失败: {e}")
            return False
    
    def show_config(self):
        """显示当前配置"""
        print("📋 当前配置:")
        print(json.dumps(self.config, ensure_ascii=False, indent=2))
    
    def add_group(self, group_name):
        """添加新的目标群聊"""
        try:
            # 检查是否已存在
            for group in self.config['target_groups']:
                if group['name'] == group_name:
                    print(f"群聊 '{group_name}' 已存在")
                    return False
            
            self.config['target_groups'].append({
                "name": group_name,
                "hwnd": None,
                "enabled": True
            })
            
            self.save_config()
            print(f"✅ 已添加群聊: {group_name}")
            return True
            
        except Exception as e:
            print(f"添加群聊失败: {e}")
            return False
    
    def test_send_to_group(self, group_name):
        """测试发送到指定群聊"""
        try:
            # 先更新窗口句柄
            self.update_target_windows()
            
            # 查找目标群聊
            target_group = None
            for group in self.config['target_groups']:
                if group['name'] == group_name:
                    target_group = group
                    break
            
            if not target_group:
                print(f"❌ 未找到群聊: {group_name}")
                return False
            
            hwnd = target_group.get('hwnd')
            if not hwnd:
                print(f"❌ 群聊 '{group_name}' 没有有效的窗口句柄")
                return False
            
            print(f"🧪 测试发送到群聊: {group_name} (句柄: {hwnd})")
            
            # 执行测试发送
            result = subprocess.run([
                sys.executable, 'direct_sender.py', 'test', str(hwnd)
            ], capture_output=True, text=True, encoding='gbk', errors='replace')
            
            if result.returncode == 0:
                print(f"✅ 测试发送成功")
                return True
            else:
                print(f"❌ 测试发送失败: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"测试发送失败: {e}")
            return False

def main():
    """主程序入口"""
    import sys
    
    system = AutoReportSystem()
    
    if len(sys.argv) < 2:
        print("🤖 每日存储统计自动化系统")
        print("用法:")
        print("  python auto_daily_report.py run                    - 执行完整自动化流程")
        print("  python auto_daily_report.py config                 - 显示当前配置")
        print("  python auto_daily_report.py update                 - 更新群聊窗口句柄") 
        print("  python auto_daily_report.py add <群名>             - 添加新的目标群聊")
        print("  python auto_daily_report.py test <群名>            - 测试发送到指定群聊")
        print("\n示例:")
        print("  python auto_daily_report.py run                    - 执行统计并发送")
        print("  python auto_daily_report.py add '数据统计群'        - 添加新群聊")
        print("  python auto_daily_report.py test '交付运维日报群'   - 测试发送")
        return
    
    command = sys.argv[1]
    
    if command == "run":
        # 执行完整自动化流程
        success = system.run_full_automation()
        exit(0 if success else 1)
        
    elif command == "config":
        # 显示配置
        system.show_config()
        
    elif command == "update":
        # 更新窗口句柄
        success = system.update_target_windows()
        if success:
            print("✅ 窗口句柄更新完成")
        else:
            print("❌ 窗口句柄更新失败")
            
    elif command == "add":
        # 添加群聊
        if len(sys.argv) < 3:
            print("请指定群聊名称: python auto_daily_report.py add <群名>")
        else:
            # 支持带空格的群名
            group_name = ' '.join(sys.argv[2:])
            system.add_group(group_name)
            
    elif command == "test":
        # 测试发送
        if len(sys.argv) < 3:
            print("请指定群聊名称: python auto_daily_report.py test <群名>")
        else:
            # 支持带空格的群名：python auto_daily_report.py test AI TESt
            group_name = ' '.join(sys.argv[2:])
            system.test_send_to_group(group_name)
            
    else:
        print(f"未知命令: {command}")

if __name__ == "__main__":
    main()