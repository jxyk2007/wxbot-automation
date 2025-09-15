# -*- coding: utf-8 -*-
"""
升级企业微信发送器到抗重启版本
"""

import shutil
import os
import logging

def upgrade_wxwork_sender():
    """升级企业微信发送器"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
    logger = logging.getLogger(__name__)

    try:
        logger.info("🔧 开始升级企业微信发送器...")

        # 1. 备份原文件
        if os.path.exists('wxwork_sender.py'):
            backup_name = 'wxwork_sender_backup.py'
            shutil.copy2('wxwork_sender.py', backup_name)
            logger.info(f"✅ 已备份原文件到: {backup_name}")

        # 2. 替换为新版本
        if os.path.exists('wxwork_sender_robust.py'):
            shutil.copy2('wxwork_sender_robust.py', 'wxwork_sender.py')
            logger.info("✅ 已替换为抗重启版本")

            # 3. 添加兼容性导入
            with open('wxwork_sender.py', 'r', encoding='utf-8') as f:
                content = f.read()

            # 确保有兼容性类
            if 'class WXWorkSender(' not in content:
                content += '''

# 兼容性接口 - 确保与现有代码兼容
class WXWorkSender(WXWorkSenderRobust):
    """兼容性包装类"""

    def __init__(self, config=None):
        super().__init__(config)
        self.is_initialized = True

    def initialize(self):
        return self.test_connection()

    def send_msg(self, message, target_group=None, **kwargs):
        return self.send_message(message, target_group)

    def SendMsg(self, message, target_group=None, **kwargs):
        return self.send_message(message, target_group)
'''

            with open('wxwork_sender.py', 'w', encoding='utf-8') as f:
                f.write(content)

            logger.info("✅ 已添加兼容性接口")

        logger.info("🎉 升级完成！")
        logger.info("现在企业微信发送器支持：")
        logger.info("  ✅ 重启后自动重新检测")
        logger.info("  ✅ 多重窗口查找策略")
        logger.info("  ✅ 强力窗口激活")
        logger.info("  ✅ 完全兼容现有代码")

        return True

    except Exception as e:
        logger.error(f"❌ 升级失败: {e}")
        return False

if __name__ == "__main__":
    upgrade_wxwork_sender()