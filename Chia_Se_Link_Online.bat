@echo off
chcp 65001 >nul
title Chia Se Excel to PDF Streamlit Online
color 0B

echo =====================================================================
echo       DANG KHOI DONG DICH VU CHIA SE LINK CHO MAY TINH KHAC
echo =====================================================================
echo.
echo [1/2] Dang kiem tra may chu Streamlit tren cong 8501...

netstat -ano | findstr :8501 >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo May chu Streamlit chua bat, dang tu dong khoi chay ngam...
    start /b py -m streamlit run excel_to_pdf_streamlit.py --server.port 8501 --server.headless true
    timeout /t 3 >nul
)

echo [2/2] Dang tao duong link HTTPS cong khai bang Cloudflare Tunnel...
echo.
echo =====================================================================
echo   VUI LONG DOI VAI GIAY DE DUONG LINK HIEN THI BEN DUOI
echo   (Copy link "https://....trycloudflare.com" va gui cho moi nguoi)
echo =====================================================================
echo.

.\tools\cloudflared.exe tunnel --url http://localhost:8501

pause
