@echo off
chcp 65001 >nul
echo ========================================
echo CorelDRAW 弹窗处理工具 - 打包脚本
echo ========================================
echo.

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

REM 安装 pyinstaller
echo [1/3] 安装打包工具...
pip install pyinstaller -q

REM 打包（带控制台，方便查看日志）
echo [2/3] 正在打包...
pyinstaller --onefile --console --name "CDR弹窗处理工具" --icon=NONE cdr_popup_handler.py

REM 清理
echo [3/3] 清理临时文件...
rmdir /s /q build 2>nul
del /q *.spec 2>nul

echo.
echo ========================================
echo 打包完成！
echo EXE 文件位置: dist\CDR弹窗处理工具.exe
echo ========================================
pause
