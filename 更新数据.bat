@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
py -3 -c "import sys; assert sys.version_info >= (3,9)" >nul 2>nul
if not errorlevel 1 (
    set "TRACKER_PY=py -3"
    goto ready
)
python -c "import sys; assert sys.version_info >= (3,9)" >nul 2>nul
if not errorlevel 1 (
    set "TRACKER_PY=python"
    goto ready
)
echo 未找到 Python 3.9 或更新版本。请从 python.org 安装并勾选 Add Python to PATH。
echo 只使用 Excel 表格不需要 Python。
pause
exit /b 1
:ready
%TRACKER_PY% -c "import openpyxl, xlsxwriter" >nul 2>nul
if errorlevel 1 (
    echo 缺少依赖。请在此目录打开终端运行：
    echo %TRACKER_PY% -m pip install -r requirements.txt
    pause
    exit /b 1
)
%TRACKER_PY% "_source\update_data.py" %*
set "TRACKER_EXIT=%ERRORLEVEL%"
echo.
if not "%TRACKER_EXIT%"=="0" echo 操作未完成，请查看上方提示。
pause
exit /b %TRACKER_EXIT%
