@echo off
chcp 65001 >nul
REM ============================================================
REM Excel 通用拆分工具 - 一键打包脚本 (Windows)
REM Copyright (c) 2026 Abelin
REM ============================================================
REM 关键：--collect-all customtkinter 必须有！
REM 否则打包出的 exe 会因找不到 customtkinter 的主题/字体资源而启动崩溃。
REM ============================================================

echo.
echo [1/3] 安装依赖...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/3] 清理旧的打包产物...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Excel通用拆分工具v2.0.spec" del "Excel通用拆分工具v2.0.spec"

echo.
echo [3/3] 开始打包...
pyinstaller --onefile --noconsole ^
  --name "Excel通用拆分工具v2.0" ^
  --icon app.ico ^
  --version-file version.txt ^
  --collect-all customtkinter ^
  --add-data "config;config" ^
  main.py

echo.
if exist "dist\Excel通用拆分工具v2.0.exe" (
    echo ============================================================
    echo  打包成功！exe 位置：dist\Excel通用拆分工具v2.0.exe
    echo ============================================================
) else (
    echo  打包失败，请检查上方错误信息。
)
echo.
pause
