@echo off
chcp 65001 >nul
REM ============================================================
REM ExcelRouter - 一键打包脚本 (Windows)
REM Copyright (c) 2026 Abelin
REM ============================================================
REM 关键：--collect-all customtkinter 必须有！
REM 否则打包出的 exe 会因找不到 customtkinter 的主题/字体资源而启动崩溃。
REM
REM 本脚本产出两种形态，推荐分发 onedir（文件夹）版：
REM   - onedir：无运行时自解压行为，启动更快，更不容易被杀毒软件误报
REM   - onefile：单文件，下载即用，但个别杀毒软件可能误报（见 docs/FAQ.md）
REM ============================================================

echo.
echo [1/4] 安装依赖...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/4] 清理旧的打包产物...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "ExcelRouter.spec" del "ExcelRouter.spec"
if exist "ExcelRouter-onefile.spec" del "ExcelRouter-onefile.spec"

echo.
echo [3/4] 打包 onedir 版（推荐分发）...
pyinstaller --onedir --noconsole ^
  --name "ExcelRouter" ^
  --icon app.ico ^
  --version-file version.txt ^
  --collect-all customtkinter ^
  --add-data "config;config" ^
  main.py

echo.
echo [4/4] 打包 onefile 版（备选，单文件）...
pyinstaller --onefile --noconsole ^
  --name "ExcelRouter-onefile" ^
  --icon app.ico ^
  --version-file version.txt ^
  --collect-all customtkinter ^
  --add-data "config;config" ^
  main.py

echo.
if exist "dist\ExcelRouter\ExcelRouter.exe" if exist "dist\ExcelRouter-onefile.exe" (
    echo ============================================================
    echo  打包成功！
    echo   onedir 版（推荐）：dist\ExcelRouter\ExcelRouter.exe
    echo   onefile 版：       dist\ExcelRouter-onefile.exe
    echo ============================================================
) else (
    echo  打包失败，请检查上方错误信息。
)
echo.
pause
