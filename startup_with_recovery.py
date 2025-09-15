# -*- coding: utf-8 -*-
"""
企业微信自动化启动脚本 - 带恢复机制
解决重启后识别不到问题的完整方案
"""

import sys
import time
import logging
import subprocess
from datetime import datetime
from auto_recovery_config import AutoRecoveryConfig
from wxwork_sender_fixed import WXWorkSenderFixed

def setup_logging():
    """设置日志"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('startup_recovery.log', encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

def check_wxwork_running():
    """检查企业微信是否运行"""
    try:
        import psutil
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() in ['wxwork.exe', 'wework.exe']:
                return True
        return False
    except Exception:
        return False

def start_wxwork():
    """启动企业微信"""
    logger = logging.getLogger(__name__)
    try:
        # 常见的企业微信安装路径
        possible_paths = [
            r"C:\Program Files (x86)\Tencent\WXWork\WXWork.exe",
            r"C:\Program Files\Tencent\WXWork\WXWork.exe",
            r"D:\Program Files\Tencent\WXWork\WXWork.exe",
            r"E:\Program Files\Tencent\WXWork\WXWork.exe"
        ]

        import os
        for path in possible_paths:
            if os.path.exists(path):
                logger.info(f"启动企业微信: {path}")
                subprocess.Popen([path])
                return True

        logger.warning("未找到企业微信安装路径")
        return False

    except Exception as e:
        logger.error(f"启动企业微信失败: {e}")
        return False

def wait_for_wxwork_ready(timeout=60):
    """等待企业微信完全启动"""
    logger = logging.getLogger(__name__)
    logger.info("等待企业微信完全启动...")

    start_time = time.time()
    while time.time() - start_time < timeout:
        if check_wxwork_running():
            # 额外等待UI完全加载
            time.sleep(5)
            logger.info("企业微信已启动")
            return True
        time.sleep(1)

    logger.error("等待企业微信启动超时")
    return False

def main():
    """主函数"""
    setup_logging()
    logger = logging.getLogger(__name__)

    logger.info("=" * 60)
    logger.info("企业微信自动化启动脚本 - 带恢复机制")
    logger.info(f"启动时间: {datetime.now()}")
    logger.info("=" * 60)

    # 1. 初始化恢复配置
    config = AutoRecoveryConfig()

    # 2. 检查企业微信是否运行
    if not check_wxwork_running():
        logger.info("企业微信未运行，尝试启动...")
        if start_wxwork():
            if not wait_for_wxwork_ready():
                logger.error("企业微信启动失败")
                return False
        else:
            logger.error("无法启动企业微信")
            return False
    else:
        logger.info("企业微信已在运行")

    # 3. 清除旧的窗口句柄（确保重新检测）
    if config.is_recovery_enabled():
        logger.info("启用自动恢复机制，清除旧窗口句柄...")
        config.invalidate_all_handles()

    # 4. 初始化企业微信发送器
    logger.info("初始化企业微信发送器...")
    sender = WXWorkSenderFixed()

    # 5. 尝试初始化（带重试机制）
    if sender.initialize():
        logger.info("✅ 企业微信发送器初始化成功！")

        # 6. 更新配置中的窗口句柄
        if sender.main_window_hwnd:
            config.update_window_handle("wxwork", config.config["senders"]["wxwork"]["default_group"], sender.main_window_hwnd)

        # 7. 测试发送（可选）
        test_message = f"🚀 企业微信自动化系统启动成功\n时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        # result = sender.send_message(test_message)
        # logger.info(f"测试消息发送结果: {result}")

        logger.info("🎉 系统启动完成，可以正常使用了！")
        return True

    else:
        logger.error("❌ 企业微信发送器初始化失败！")

        # 诊断信息
        logger.info("诊断信息:")
        logger.info(f"- 企业微信进程: {'运行' if check_wxwork_running() else '未运行'}")
        logger.info(f"- 自动恢复: {'启用' if config.is_recovery_enabled() else '禁用'}")

        return False

def run_with_schedule():
    """定时运行版本"""
    logger = logging.getLogger(__name__)

    while True:
        try:
            logger.info("\n" + "=" * 50)
            logger.info("定时检查企业微信状态...")

            if main():
                logger.info("系统运行正常")
            else:
                logger.warning("系统检查发现问题")

            # 每30分钟检查一次
            logger.info("下次检查时间: 30分钟后")
            time.sleep(30 * 60)

        except KeyboardInterrupt:
            logger.info("用户中断，程序退出")
            break
        except Exception as e:
            logger.error(f"定时检查出错: {e}")
            time.sleep(60)  # 出错后1分钟后重试

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--schedule":
        run_with_schedule()
    else:
        success = main()
        sys.exit(0 if success else 1)