@echo off
setlocal
title JARVIS

:: Check .env
if not exist ".env" (
    echo ERROR: .env not found. Run setup.bat first.
    pause & exit /b 1
)

:: Check API keys are set
findstr /c:"your-anthropic-api-key-here" .env >nul 2>&1
if not errorlevel 1 (
    echo WARNING: ANTHROPIC_API_KEY not set in .env
    echo Edit .env and add your key, then re-run start.bat
    pause & exit /b 1
)

echo.
echo  Starting J.A.R.V.I.S...
echo.

:: Start backend in new terminal window
start "JARVIS Backend" cmd /k "python server.py"

:: Wait for backend to start
timeout /t 3 /nobreak >nul

:: Start frontend dev server
start "JARVIS Frontend" cmd /k "cd frontend && npm run dev"

:: Wait for frontend to start
timeout /t 3 /nobreak >nul

:: Open browser (Chrome preferred for Web Speech API)
set CHROME=
for %%p in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LocalAppData%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%p set CHROME=%%p
)

if defined CHROME (
    start "" %CHROME% --allow-running-insecure-content https://localhost:5173
) else (
    start https://localhost:5173
)

echo  JARVIS is starting...
echo  Backend:  https://localhost:8340
echo  Frontend: https://localhost:5173
echo.
echo  If you see a certificate warning in Chrome:
echo  Click Advanced then Proceed to localhost
echo.
