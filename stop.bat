@echo off
setlocal enabledelayedexpansion

echo [INFO] 正在查找 feishu_docget 服务进程...

:: 使用 wmic 查找运行 src/app.py 的 Python 进程
set FOUND=0

for /f "tokens=2 delims=," %%i in ('wmic process where "commandline like '%%src\\app.py%%' and name like '%%python%%'" get processid /format:csv 2^>nul ^| findstr /r "[0-9]"') do (
    echo [INFO] 找到进程 PID: %%i，正在终止...
    taskkill /PID %%i /F >nul 2>&1
    if !errorlevel! equ 0 (
        echo [INFO] 已成功终止进程 %%i
    ) else (
        echo [WARN] 无法终止进程 %%i
    )
    set FOUND=1
)

if "%FOUND%"=="0" (
    echo [INFO] 未找到正在运行的 feishu_docget 服务。
) else (
    echo [INFO] feishu_docget 服务已停止。
)

pause
