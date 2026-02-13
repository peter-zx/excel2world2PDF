@echo off
echo ========================================
echo   合同自动填写工具 - 安装与启动
echo ========================================
echo.

cd /d %~dp0

echo [1/3] 激活虚拟环境...
call venv\Scripts\activate.bat

echo [2/3] 检查依赖...
pip show pandas >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖...
    pip install pandas python-docx docxtpl openpyxl streamlit python-dateutil
) else (
    echo 依赖已安装
)

echo [3/3] 启动应用...
echo.
echo 应用将在浏览器中打开 http://localhost:8501
echo 按 Ctrl+C 可停止应用
echo.
streamlit run app.py

pause
