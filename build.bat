@echo off
chcp 65001 >nul
echo ========================================
echo   ChroLens_Clear 1.4 打包腳本
echo ========================================
echo.

REM 檢查是否安裝 pyinstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [錯誤] 未安裝 PyInstaller，正在安裝...
    pip install pyinstaller
    if errorlevel 1 (
        echo [錯誤] PyInstaller 安裝失敗
        pause
        exit /b 1
    )
)

echo [1/4] 清理舊的打包檔案...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"
echo      ✓ 清理完成

echo.
echo [2/4] 檢查必要檔案...
if not exist "ChroLens_Clear.py" (
    echo [錯誤] 找不到 ChroLens_Clear.py
    pause
    exit /b 1
)
if not exist "Nekoneko.ico" (
    echo [警告] 找不到 Nekoneko.ico，將不使用圖示
    set ICO_OPTION=
) else (
    set ICO_OPTION=--icon=Nekoneko.ico --add-data "Nekoneko.ico;."
    echo      ✓ 圖示檔案存在
)

echo.
echo [3/4] 開始打包...
echo      這可能需要幾分鐘時間，請稍候...
echo.

pyinstaller --onedir --noconsole ^
    %ICO_OPTION% ^
    --hidden-import=win32timezone ^
    --hidden-import=win32gui ^
    --hidden-import=win32con ^
    --hidden-import=ttkbootstrap ^
    --name="ChroLens_Clear_1.4" ^
    ChroLens_Clear.py

if errorlevel 1 (
    echo.
    echo [錯誤] 打包失敗
    pause
    exit /b 1
)

echo.
echo [4/4] 複製必要檔案到 dist 目錄...
if exist "Nekoneko.ico" copy /y "Nekoneko.ico" "dist\ChroLens_Clear_1.4\" >nul
if exist "config.json" copy /y "config.json" "dist\ChroLens_Clear_1.4\" >nul
echo      ✓ 檔案複製完成

echo.
echo ========================================
echo   ✓ 打包完成！
echo ========================================
echo.
echo 執行檔位置: dist\ChroLens_Clear_1.4\ChroLens_Clear_1.4.exe
echo.
echo 按任意鍵開啟 dist 資料夾...
pause >nul
explorer "dist\ChroLens_Clear_1.4"
