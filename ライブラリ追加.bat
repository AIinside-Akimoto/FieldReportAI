@echo off
chcp 65001 > nul
cd /d %~dp0

echo =================================================
echo  業務特化AIエージェントサービス 環境セットアップ
echo =================================================

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Pythonが見つかりません。
    echo PDFファイルを参考にPythonをインストールしてください。
    pause
    exit /b
)

echo.
echo [INFO] 必要なライブラリをインストールしています...
pip install -q requests streamlit streamlit-back-camera-input python-docx

echo.
echo =================================================
echo  セットアップ完了！
echo =================================================
pause