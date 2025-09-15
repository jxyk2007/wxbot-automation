# -*- coding: utf-8 -*-
"""
企业微信问题修复 - 简化版
基于测试通过的检测逻辑
"""

import logging
import time
import win32gui
import win32process
import psutil
import json
from datetime import datetime

logger = logging.getLogger(__name__)

def find_wxwork_main_window():
    """查找企业微信主窗口 - 基于测试通过的逻辑"""

    logger.info("🔍 查找企业微信主窗口...")

    # 1. 检测企业微信进程
    wxwork_processes = []
    process_names = ["WXWork.exe", "wxwork.exe"]

    for proc in psutil.process_iter(['pid', 'name', 'exe', 'memory_info']):
        try:
            proc_name = proc.info['name']
            if any(name.lower() in proc_name.lower() for name in process_names):
                memory_mb = proc.info['memory_info'].rss / 1024 / 1024
                wxwork_processes.append({
                    'pid': proc.pid,
                    'name': proc_name,
                    'memory_mb': memory_mb
                })
                logger.info(f"  找到进程: {proc_name} (PID: {proc.pid}, 内存: {memory_mb:.1f}MB)")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if not wxwork_processes:
        logger.error("❌ 未找到企业微信进程")
        return None

    # 选择内存最大的进程（主进程）
    main_process = max(wxwork_processes, key=lambda p: p['memory_mb'])
    logger.info(f"✅ 选定主进程: PID {main_process['pid']} ({main_process['memory_mb']:.1f}MB)")

    # 2. 枚举窗口
    def enum_windows_callback(hwnd, windows_list):
        try:
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            window_title = win32gui.GetWindowText(hwnd)
            is_visible = win32gui.IsWindowVisible(hwnd)

            windows_list.append({
                'hwnd': hwnd,
                'pid': window_pid,
                'class': class_name,
                'title': window_title,
                'visible': is_visible
            })
        except Exception:
            pass

    windows_list = []
    win32gui.EnumWindows(enum_windows_callback, windows_list)

    # 查找属于企业微信进程的窗口
    wxwork_windows = [w for w in windows_list if w['pid'] == main_process['pid']]

    logger.info(f"  找到 {len(wxwork_windows)} 个企业微信窗口")

    # 3. 查找主窗口
    main_candidates = []
    for w in wxwork_windows:
        if w['class'].lower() == 'weworkwindow':
            score = 0
            if w['title'] == '企业微信':
                score += 10
            if w['visible']:
                score += 5
            if w['title'].strip():
                score += 2

            w['score'] = score
            main_candidates.append(w)
            logger.info(f"    候选窗口: '{w['title']}' (分数: {score})")

    if not main_candidates:
        logger.error("❌ 未找到 WeWorkWindow 类型的窗口")
        return None

    # 选择得分最高的窗口
    best_window = max(main_candidates, key=lambda x: x['score'])
    logger.info(f"✅ 找到主窗口: '{best_window['title']}' (句柄: {best_window['hwnd']})")

    # 4. 验证窗口
    if win32gui.IsWindow(best_window['hwnd']) and win32gui.IsWindowVisible(best_window['hwnd']):
        logger.info("✅ 窗口验证通过")
        return best_window
    else:
        logger.error("❌ 窗口验证失败")
        return None

def update_config_with_new_hwnd(hwnd):
    """更新配置文件中的窗口句柄"""
    config_file = "auto_report_config.json"

    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)

        # 更新企业微信配置
        if 'senders' in config and 'wxwork' in config['senders']:
            wxwork_config = config['senders']['wxwork']

            # 更新所有目标群组的句柄
            for group in wxwork_config.get('target_groups', []):
                old_hwnd = group.get('hwnd')
                group['hwnd'] = hwnd
                group['last_update'] = datetime.now().isoformat()
                logger.info(f"✅ 更新群聊 '{group.get('name', 'Unknown')}' 句柄: {old_hwnd} → {hwnd}")

        # 更新时间戳
        config['last_update'] = datetime.now().isoformat()

        # 保存配置
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

        logger.info(f"✅ 配置已更新并保存到: {config_file}")
        return True

    except Exception as e:
        logger.error(f"❌ 更新配置文件失败: {e}")
        return False

def test_window_interaction(hwnd):
    """测试窗口交互"""
    try:
        logger.info("🔧 测试窗口交互...")

        # 激活窗口
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)

        # 获取窗口位置
        rect = win32gui.GetWindowRect(hwnd)
        logger.info(f"  窗口位置: {rect}")

        # 检查窗口状态
        title = win32gui.GetWindowText(hwnd)
        logger.info(f"  窗口标题: '{title}'")

        logger.info("✅ 窗口交互测试通过")
        return True

    except Exception as e:
        logger.error(f"❌ 窗口交互测试失败: {e}")
        return False

def main():
    """主修复函数"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    logger.info("🚀 企业微信窗口句柄修复工具")
    logger.info("=" * 50)

    # 1. 查找企业微信主窗口
    main_window = find_wxwork_main_window()

    if not main_window:
        logger.error("❌ 无法找到企业微信主窗口")
        return False

    hwnd = main_window['hwnd']

    # 2. 测试窗口交互
    if not test_window_interaction(hwnd):
        logger.error("❌ 窗口交互测试失败")
        return False

    # 3. 更新配置文件
    if not update_config_with_new_hwnd(hwnd):
        logger.error("❌ 配置文件更新失败")
        return False

    logger.info("=" * 50)
    logger.info("🎉 企业微信窗口句柄修复完成！")
    logger.info(f"主窗口句柄: {hwnd}")
    logger.info("现在可以正常使用企业微信自动化功能了")

    return True

if __name__ == "__main__":
    success = main()
    if success:
        print("\n✅ 修复成功！企业微信重启问题已解决")
    else:
        print("\n❌ 修复失败！请检查企业微信是否正常运行")