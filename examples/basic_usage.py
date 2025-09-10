# -*- coding: utf-8 -*-
"""
wxbot 基本使用示例
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from direct_sender import DirectSender
from window_inspector import WindowInspector

def example_1_find_wechat_windows():
    """示例1: 查找微信窗口"""
    print("=== 示例1: 查找微信窗口 ===")
    
    inspector = WindowInspector()
    
    # 查找微信主进程
    main_pid = inspector.find_main_wechat_process()
    if main_pid:
        print(f"找到微信主进程: PID {main_pid}")
        
        # 显示该进程的所有窗口
        inspector.show_windows_for_pid(main_pid)
    else:
        print("未找到微信进程")

def example_2_send_message():
    """示例2: 发送消息到指定窗口"""
    print("=== 示例2: 发送测试消息 ===")
    
    # 注意: 请替换为你的实际窗口句柄
    target_hwnd = 12345678  # 这里需要替换为实际的窗口句柄
    
    sender = DirectSender()
    
    # 发送测试消息
    test_message = "🤖 这是一条来自 wxbot 的测试消息！"
    success = sender.test_message_to_window(target_hwnd, test_message)
    
    if success:
        print("✅ 消息发送成功！")
    else:
        print("❌ 消息发送失败！")

def example_3_interactive_mode():
    """示例3: 交互式模式"""
    print("=== 示例3: 交互式获取窗口信息 ===")
    print("请按照提示操作...")
    
    inspector = WindowInspector()
    inspector.click_to_inspect()

def example_4_custom_automation():
    """示例4: 自定义自动化脚本"""
    print("=== 示例4: 自定义自动化 ===")
    
    # 步骤1: 查找微信窗口
    inspector = WindowInspector()
    main_pid = inspector.find_main_wechat_process()
    
    if not main_pid:
        print("未找到微信进程")
        return
    
    chat_windows = inspector.find_wechat_windows(main_pid)
    
    if not chat_windows:
        print("未找到聊天窗口")
        return
    
    # 步骤2: 选择目标窗口（这里选择第一个）
    target_window = chat_windows[0]
    target_hwnd = target_window['hwnd']
    
    print(f"目标窗口: {target_window['title']} (句柄: {target_hwnd})")
    
    # 步骤3: 发送自定义消息
    sender = DirectSender()
    
    messages = [
        "📊 自动化报告开始",
        "🔍 系统检查完成",
        "✅ 所有服务运行正常",
        "📝 报告生成完成"
    ]
    
    for i, message in enumerate(messages, 1):
        print(f"发送消息 {i}/{len(messages)}: {message}")
        success = sender.send_message_to_window(target_hwnd, message)
        
        if success:
            print(f"✅ 消息 {i} 发送成功")
        else:
            print(f"❌ 消息 {i} 发送失败")
        
        # 避免发送过快
        import time
        time.sleep(2)

def main():
    """主菜单"""
    while True:
        print("\n" + "="*50)
        print("🤖 wxbot 使用示例")
        print("="*50)
        print("1. 查找微信窗口")
        print("2. 发送测试消息")
        print("3. 交互式获取窗口信息") 
        print("4. 自定义自动化脚本")
        print("0. 退出")
        print("="*50)
        
        choice = input("请选择示例 (0-4): ").strip()
        
        if choice == "1":
            example_1_find_wechat_windows()
        elif choice == "2":
            example_2_send_message()
        elif choice == "3":
            example_3_interactive_mode()
        elif choice == "4":
            example_4_custom_automation()
        elif choice == "0":
            print("👋 再见！")
            break
        else:
            print("❌ 无效选择，请重新输入")
        
        input("\n按回车键继续...")

if __name__ == "__main__":
    main()