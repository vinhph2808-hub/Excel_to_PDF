@echo off
chcp 65001 >nul
title Excel to PDF Converter Pro - Trình Chuyển Đổi Trên Web
color 0A

echo =====================================================================
echo    EXCEL TO PDF CONVERTER PRO - WEB APPLICATION (STREAMLIT)
echo =====================================================================
echo.
echo [1/2] Dang kiem tra moi truong Python...
where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON_CMD=py
) else (
    set PYTHON_CMD=python
)

echo [2/2] Dang khoi dong may chu Web Streamlit tai http://localhost:8501 ...
echo.
echo Ban co the mo trinh duyet bat ky (Chrome, Edge, Firefox) tai:
echo      http://localhost:8501
echo.
echo Nhan Ctrl+C de dung ung dung khi can.
echo =====================================================================
echo.

%PYTHON_CMD% -m streamlit run excel_to_pdf_streamlit.py --server.port 8501 --server.headless false

pause
