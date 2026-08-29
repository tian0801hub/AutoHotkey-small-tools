@echo off
setlocal enabledelayedexpansion
title (@天tian)图片批量序号重命名工具（保留原后缀）

:: ============== 可配置参数 ==============
set "start_num=1"
set "file_filter=*.jpg *.jpeg *.png *.bmp *.gif *.webp"
set "sort_order=/on"
:: ========================================

echo ==============================================
echo      (@天tian)图片批量序号重命名工具（保留原后缀）
echo ==============================================
echo 功能：按顺序重命名为 1.原后缀, 2.原后缀 ...
echo 起始编号：%start_num%
echo 处理格式：%file_filter%
echo ==============================================
echo.
set /p "confirm=确认开始重命名？(输入 Y 继续，其他键退出): "
if /i not "%confirm%"=="Y" (
    echo 已取消操作。
    pause
    exit /b
)

echo.
echo 正在处理中...

set "num=%start_num%"
set "success=0"
set "failed=0"

for /f "delims=" %%f in ('dir /b /a-d %sort_order% %file_filter% 2^>nul') do (
    ren "%%f" "!num!%%~xf"
    if !errorlevel! equ 0 (
        set /a success+=1
    ) else (
        set /a failed+=1
        echo 重命名失败：%%f
    )
    set /a num+=1
)

echo.
echo ==============================================
echo 处理完成！
echo 成功：%success% 个
echo 失败：%failed% 个
echo ==============================================
pause
endlocal
