@echo off
setlocal

:: Check if Word is running
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I "WINWORD.EXE" >NUL
if %ERRORLEVEL% equ 0 (
    echo Error: Microsoft Word is running. Please close Word and try again.
    exit /b 2
)

:: 定义源文件和目标路径
set "source=%~dp0Normal.dotm"
set "target=%appdata%\Microsoft\Templates\Normal.dotm"

:: 检查源文件是否存在
if not exist "%source%" (
    echo Error: Source file not found: %source%
    exit /b 1
)

:: 检查并删除旧的备份文件
if exist "%appdata%\Microsoft\Templates\Normal.dotm.bak" (
    del "%appdata%\Microsoft\Templates\Normal.dotm.bak"
    echo Deleted existing Normal.dotm.bak
)

:: 备份原文件
if exist "%target%" (
    ren "%target%" "Normal.dotm.bak"
    echo Renamed existing Normal.dotm to Normal.dotm.bak
)

:: 复制新文件
copy "%source%" "%target%"
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to copy file
    exit /b 1
)

endlocal
