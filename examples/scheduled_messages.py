# -*- coding: utf-8 -*-
"""
wxbot 定时消息发送示例
使用 schedule 库实现定时发送功能
"""

import sys
import os
import time
import schedule
from datetime import datetime

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from direct_sender import DirectSender
from window_inspector import WindowInspector

class ScheduledBot:
    def __init__(self):
        self.sender = DirectSender()
        self.inspector = WindowInspector()
        self.target_windows = {}  # 存储群聊名称和窗口句柄的映射
        
    def setup_target_windows(self):
        """设置目标窗口"""
        print("🔍 正在查找微信窗口...")
        
        # 查找微信主进程
        main_pid = self.inspector.find_main_wechat_process()
        if not main_pid:
            print("❌ 未找到微信进程")
            return False
        
        # 获取所有聊天窗口
        chat_windows = self.inspector.find_wechat_windows(main_pid)
        if not chat_windows:
            print("❌ 未找到聊天窗口")
            return False
        
        # 显示所有可用窗口
        print("📋 发现以下聊天窗口:")
        for i, window in enumerate(chat_windows, 1):
            print(f"  {i}. {window['title']} (句柄: {window['hwnd']})")
            self.target_windows[window['title']] = window['hwnd']
        
        return True
    
    def send_morning_greeting(self):
        """发送早晨问候"""
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        message = f"""☀️ 早安！
        
🕐 当前时间: {current_time}
📅 新的一天开始了！

🤖 来自定时机器人的自动问候"""
        
        self.send_to_all_groups(message)
    
    def send_work_reminder(self):
        """发送工作提醒"""
        current_time = datetime.now().strftime('%H:%M')
        message = f"""⏰ 工作提醒 ({current_time})
        
📝 别忘了今天的重要任务：
• 检查邮件
• 更新项目进度
• 准备会议材料

💪 加油工作！"""
        
        self.send_to_all_groups(message)
    
    def send_evening_summary(self):
        """发送晚间总结"""
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        message = f"""🌙 晚间总结
        
🕐 时间: {current_time}
📊 今日工作已完成
🎯 明日计划已制定

😴 早点休息，明天见！

🤖 定时机器人"""
        
        self.send_to_all_groups(message)
    
    def send_custom_reminder(self, reminder_text):
        """发送自定义提醒"""
        current_time = datetime.now().strftime('%H:%M')
        message = f"""🔔 定时提醒 ({current_time})
        
{reminder_text}

🤖 wxbot 定时提醒"""
        
        self.send_to_all_groups(message)
    
    def send_to_all_groups(self, message):
        """发送消息到所有群组"""
        if not self.target_windows:
            print("❌ 没有配置目标窗口")
            return
        
        success_count = 0
        for group_name, hwnd in self.target_windows.items():
            print(f"📨 发送消息到: {group_name}")
            
            success = self.sender.send_message_to_window(hwnd, message)
            if success:
                print(f"✅ 发送到 '{group_name}' 成功")
                success_count += 1
            else:
                print(f"❌ 发送到 '{group_name}' 失败")
            
            # 避免发送过快
            time.sleep(1)
        
        print(f"📊 发送完成: {success_count}/{len(self.target_windows)} 个群聊发送成功")
    
    def send_to_specific_group(self, group_name, message):
        """发送消息到指定群组"""
        if group_name not in self.target_windows:
            print(f"❌ 未找到群聊: {group_name}")
            return False
        
        hwnd = self.target_windows[group_name]
        return self.sender.send_message_to_window(hwnd, message)

def setup_schedule(bot):
    """设置定时任务"""
    print("⏰ 设置定时任务...")
    
    # 每天早上 9:00 发送问候
    schedule.every().day.at("09:00").do(bot.send_morning_greeting)
    
    # 每天下午 14:00 发送工作提醒
    schedule.every().day.at("14:00").do(bot.send_work_reminder)
    
    # 每天晚上 18:00 发送总结
    schedule.every().day.at("18:00").do(bot.send_evening_summary)
    
    # 每周一早上 10:00 发送周会提醒
    schedule.every().monday.at("10:00").do(
        bot.send_custom_reminder, 
        "📅 提醒：今天有周会，请准备好汇报材料"
    )
    
    # 每天每隔2小时发送一次系统状态（工作时间）
    for hour in range(9, 18, 2):  # 9, 11, 13, 15, 17
        schedule.every().day.at(f"{hour:02d}:00").do(
            bot.send_custom_reminder,
            "💻 系统状态检查：所有服务运行正常"
        )
    
    print("✅ 定时任务设置完成")
    print("📋 任务列表:")
    for job in schedule.jobs:
        print(f"  • {job}")

def run_scheduler():
    """运行调度器"""
    print("🚀 定时任务调度器启动...")
    print("按 Ctrl+C 退出")
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # 每分钟检查一次
    except KeyboardInterrupt:
        print("\n👋 定时任务调度器已停止")

def interactive_test(bot):
    """交互式测试"""
    print("\n🧪 交互式测试模式")
    print("可用群聊:")
    
    group_list = list(bot.target_windows.keys())
    for i, group_name in enumerate(group_list, 1):
        print(f"  {i}. {group_name}")
    
    try:
        choice = int(input("选择群聊编号: ")) - 1
        if 0 <= choice < len(group_list):
            group_name = group_list[choice]
            message = input("输入测试消息: ")
            
            success = bot.send_to_specific_group(group_name, message)
            if success:
                print("✅ 测试消息发送成功")
            else:
                print("❌ 测试消息发送失败")
        else:
            print("❌ 无效选择")
    except ValueError:
        print("❌ 请输入有效数字")

def main():
    """主程序"""
    print("🤖 wxbot 定时消息发送器")
    print("=" * 50)
    
    # 创建机器人实例
    bot = ScheduledBot()
    
    # 设置目标窗口
    if not bot.setup_target_windows():
        print("❌ 初始化失败")
        return
    
    while True:
        print("\n📋 操作菜单:")
        print("1. 立即发送早晨问候")
        print("2. 立即发送工作提醒")
        print("3. 立即发送晚间总结")
        print("4. 发送自定义消息")
        print("5. 交互式测试")
        print("6. 启动定时调度器")
        print("0. 退出")
        
        choice = input("请选择操作 (0-6): ").strip()
        
        if choice == "1":
            bot.send_morning_greeting()
        elif choice == "2":
            bot.send_work_reminder()
        elif choice == "3":
            bot.send_evening_summary()
        elif choice == "4":
            custom_message = input("输入自定义消息: ")
            bot.send_custom_reminder(custom_message)
        elif choice == "5":
            interactive_test(bot)
        elif choice == "6":
            setup_schedule(bot)
            run_scheduler()
        elif choice == "0":
            print("👋 再见！")
            break
        else:
            print("❌ 无效选择")

if __name__ == "__main__":
    # 检查是否安装了 schedule
    try:
        import schedule
    except ImportError:
        print("❌ 缺少依赖: pip install schedule")
        exit(1)
    
    main()