@echo off
chcp 65001 >nul
echo 🤖 启动每日存储统计自动化系统...
echo =====================================

cd /d "%~dp0"

REM 激活虚拟环境（如果存在）
if exist "wxbot-py3.9\Scripts\activate.bat" (
    echo 激活虚拟环境...
    call wxbot-py3.9\Scripts\activate.bat
)

REM 执行自动化流程
echo 开始执行自动化流程...
python auto_daily_report.py run

REM 检查执行结果
if %errorlevel% == 0 (
    echo.
    echo ✅ 自动化流程执行成功！
    echo 📊 存储统计报告已生成并发送到微信群
) else (
    echo.
    echo ❌ 自动化流程执行失败！
    echo 请检查日志文件 auto_report.log
)

echo.
echo 按任意键退出...
pause >nul