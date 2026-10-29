@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -3.12 -m venv .venv
  if errorlevel 1 goto error
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto error
if not exist ".env" copy ".env.example" ".env" >nul
echo.
echo IMPORTANT: Edit .env and add GROQ_API_KEY to enable Groq answers.
echo You can verify the key using the "Test Groq connection" button in the browser.
echo Open http://127.0.0.1:8000 after server starts.
".venv\Scripts\python.exe" -m uvicorn app:app --host 127.0.0.1 --port 8000
goto end
:error
echo Setup failed. Check Python 3.12 and internet access, then retry.
pause
:end
