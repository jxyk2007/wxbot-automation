# -*- coding: utf-8 -*-
"""
快速测试企业微信发送器
"""

import logging
import sys
import os
import time
import win32gui
import win32process
import psutil

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_wxwork_detection():
    """测试企业微信检测"""

    logger.info("🔍 开始测试企业微信检测...")

    # 1. 检测企业微信进程
    logger.info("步骤1: 检测企业微信进程")
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
        return False

    # 选择内存最大的进程
    main_process = max(wxwork_processes, key=lambda p: p['memory_mb'])
    logger.info(f"✅ 选定主进程: PID {main_process['pid']} ({main_process['memory_mb']:.1f}MB)")

    # 2. 检测企业微信窗口
    logger.info("步骤2: 检测企业微信窗口")

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

    # 枚举所有窗口
    windows_list = []
    win32gui.EnumWindows(enum_windows_callback, windows_list)

    # 查找属于企业微信进程的窗口
    wxwork_windows = [w for w in windows_list if w['pid'] == main_process['pid']]

    logger.info(f"  找到 {len(wxwork_windows)} 个企业微信窗口:")
    for w in wxwork_windows:
        logger.info(f"    - 类名: {w['class']}, 标题: '{w['title']}', 可见: {w['visible']}, 句柄: {w['hwnd']}")

    if not wxwork_windows:
        logger.error("❌ 未找到企业微信窗口")
        return False

    # 3. 查找主窗口
    logger.info("步骤3: 查找主窗口")

    # 策略1: WeWorkWindow + 企业微信标题
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
            logger.info(f"    候选窗口: '{w['title']}' (类名: {w['class']}, 分数: {score})")

    if main_candidates:
        # 选择得分最高的窗口
        best_window = max(main_candidates, key=lambda x: x['score'])
        logger.info(f"✅ 找到主窗口: '{best_window['title']}' (句柄: {best_window['hwnd']})")

        # 4. 验证窗口
        logger.info("步骤4: 验证窗口")
        if win32gui.IsWindow(best_window['hwnd']) and win32gui.IsWindowVisible(best_window['hwnd']):
            logger.info("✅ 窗口验证通过")
            return True
        else:
            logger.error("❌ 窗口验证失败")
            return False
    else:
        logger.error("❌ 未找到合适的主窗口")
        return False

if __name__ == "__main__":
    logger.info("🚀 企业微信检测快速测试")
    logger.info("=" * 50)

    success = test_wxwork_detection()

    logger.info("=" * 50)
    if success:
        logger.info("🎉 检测成功！企业微信可以正常识别")
    else:
        logger.info("❌ 检测失败！请检查企业微信是否正常运行")

    logger.info("测试完成")