@echo off
REM ===========================================================
REM  GRC 團名修正工具 - 啟動腳本
REM  台中勤美洲際酒店
REM ===========================================================
chcp 65001 > nul
title GRC 團名修正工具

cd /d "%~dp0"

echo.
echo ===========================================================
echo   GRC 團名修正工具
echo   台中勤美洲際酒店
echo ===========================================================
echo.

REM --- 檢查 Python ---
python --version > nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.10 以上
    echo        https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM --- 檢查套件 ---
python -c "import flask, openpyxl" > nul 2>&1
if errorlevel 1 (
    echo [安裝] 首次執行，正在安裝相依套件...
    python -m pip install --quiet -r requirements.txt
    if errorlevel 1 (
        echo [錯誤] 套件安裝失敗，請檢查網路連線
        pause
        exit /b 1
    )
    echo [OK] 安裝完成
    echo.
)

REM --- 自動開啟瀏覽器（延遲 2 秒等伺服器啟動）---
start "" /b cmd /c "timeout /t 2 /nobreak > nul & start http://127.0.0.1:5000"

REM --- 啟動 Flask ---
echo [啟動中] 伺服器將在 http://127.0.0.1:5000 執行
echo [啟動中] 瀏覽器會自動開啟，如未開啟請手動貼上網址
echo.
echo ★ 關閉此視窗即可停止服務 ★
echo.
python app.py

pause
