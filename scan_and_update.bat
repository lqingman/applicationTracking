@echo off
echo ============================================
echo   Job Application Tracker
echo ============================================
echo.

REM Install dependencies if not already installed
echo Checking dependencies...
pip install -q pywin32 flask pandas

echo.
echo Starting email scan from Outlook...
python email_scraper.py

echo.
echo Starting dashboard server...
echo Opening http://localhost:5000 in your browser...
start http://localhost:5000

python app.py

pause
