@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

:: ============================================
::  构建 portable 版（先 pull 再打包）
:: ============================================

echo.
echo  [1/2] 正在拉取最新代码...
git pull
if errorlevel 1 (
    echo.
    echo  git pull 失败，请检查网络或仓库状态。
    pause
    exit /b 1
)
echo.

echo  [2/2] 正在打包 portable 版本...
echo.
call npm run electron:build:portable
if errorlevel 1 (
    echo.
    echo  打包失败，请检查上方错误信息。
    pause
    exit /b 1
)

echo.
echo  打包完成！
pause
