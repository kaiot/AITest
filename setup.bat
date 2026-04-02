@echo off
setlocal enabledelayedexpansion
title JARVIS Windows Setup

echo.
echo  ============================================
echo   J.A.R.V.I.S. Windows 11 Setup
echo  ============================================
echo.

:: ── Check Python ────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.11+ from https://python.org
    pause & exit /b 1
)
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] Python %PYVER%

:: ── Check Node ──────────────────────────────────────────────────────────────
node --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Node.js not found. Install Node.js 18+ from https://nodejs.org
    pause & exit /b 1
)
for /f %%v in ('node --version') do set NODEVER=%%v
echo [OK] Node.js %NODEVER%

:: ── Check Git ───────────────────────────────────────────────────────────────
git --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Git not found. Install Git from https://git-scm.com
    pause & exit /b 1
)
echo [OK] Git

:: ── Check Claude Code CLI ───────────────────────────────────────────────────
where claude >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Installing Claude Code CLI...
    npm install -g @anthropic-ai/claude-code
    if errorlevel 1 (
        echo WARNING: Could not install Claude Code. Install manually:
        echo   npm install -g @anthropic-ai/claude-code
    )
) else (
    echo [OK] Claude Code CLI
)

:: ── Download cross-platform files from original repo ────────────────────────
echo.
echo  Downloading source files from GitHub...
echo.

set BASE=https://raw.githubusercontent.com/ethanplusai/jarvis/main

:: Python modules (cross-platform, keep verbatim)
for %%f in (
    server.py
    memory.py
    browser.py
    planner.py
    conversation.py
    learning.py
    suggestions.py
    tracking.py
    evolution.py
    ab_testing.py
    monitor.py
    qa.py
    templates.py
    dispatch_registry.py
) do (
    if not exist "%%f" (
        curl -fsSL "%BASE%/%%f" -o "%%f"
        if errorlevel 1 (
            echo WARNING: Could not download %%f
        ) else (
            echo   [downloaded] %%f
        )
    ) else (
        echo   [exists]     %%f
    )
)

:: YAML templates directory
if not exist "templates\prompts" mkdir "templates\prompts"
for %%f in (
    api_integration.yaml
    bug_fix.yaml
    feature_add.yaml
    landing_page.yaml
    refactor.yaml
    research_report.yaml
) do (
    if not exist "templates\prompts\%%f" (
        curl -fsSL "%BASE%/templates/prompts/%%f" -o "templates\prompts\%%f" 2>nul
    )
)

:: Frontend files
if not exist "frontend\src" mkdir "frontend\src"

for %%f in (index.html package.json tsconfig.json) do (
    if not exist "frontend\%%f" (
        curl -fsSL "%BASE%/frontend/%%f" -o "frontend\%%f" 2>nul
        echo   [downloaded] frontend/%%f
    )
)

for %%f in (main.ts orb.ts voice.ts ws.ts settings.ts style.css) do (
    if not exist "frontend\src\%%f" (
        curl -fsSL "%BASE%/frontend/src/%%f" -o "frontend\src\%%f" 2>nul
        echo   [downloaded] frontend/src/%%f
    )
)

:: ── Apply Windows patches to server.py ─────────────────────────────────────
echo.
echo  Applying Windows patches to server.py...
python apply_windows_patches.py
if errorlevel 1 (
    echo WARNING: Patching failed. Server may still have macOS-specific code.
)

:: ── Write Windows vite.config.ts ────────────────────────────────────────────
echo.
echo  Writing frontend config...

(
echo import { defineConfig } from 'vite'
echo.
echo export default defineConfig^(^{
echo   server: ^{
echo     https: true,
echo     port: 5173,
echo     proxy: ^{
echo       '/ws': ^{
echo         target: 'wss://localhost:8340',
echo         ws: true,
echo         secure: false,
echo       ^},
echo       '/api': ^{
echo         target: 'https://localhost:8340',
echo         secure: false,
echo       ^},
echo     ^},
echo   ^},
echo ^}^)
) > "frontend\vite.config.ts"
echo   [written] frontend/vite.config.ts

:: ── Install Python dependencies ─────────────────────────────────────────────
echo.
echo  Installing Python dependencies...
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: pip install failed.
    pause & exit /b 1
)
echo [OK] Python packages installed

:: ── Install Playwright browsers ─────────────────────────────────────────────
echo.
echo  Installing Playwright Chromium...
python -m playwright install chromium
echo [OK] Playwright

:: ── Install Node dependencies ────────────────────────────────────────────────
echo.
echo  Installing frontend packages...
cd frontend
call npm install
if errorlevel 1 (
    echo ERROR: npm install failed.
    cd ..
    pause & exit /b 1
)

:: Add basic-ssl plugin for Vite HTTPS
call npm install --save-dev @vitejs/plugin-basic-ssl 2>nul
cd ..
echo [OK] Frontend packages installed

:: ── Generate SSL certificates ────────────────────────────────────────────────
echo.
echo  Generating SSL certificates...
if not exist "cert.pem" (
    python generate_certs.py
) else (
    echo   [exists] cert.pem / key.pem
)

:: ── Create data directory ────────────────────────────────────────────────────
if not exist "data" mkdir "data"

:: ── Install espeak-ng for Kokoro TTS ─────────────────────────────────────────
echo.
echo  NOTE: Kokoro TTS requires espeak-ng on Windows.
echo  If you hear no voice, download and install espeak-ng from:
echo    https://github.com/espeak-ng/espeak-ng/releases
echo  (get the .msi installer)
echo.

:: ── Copy .env template ───────────────────────────────────────────────────────
if not exist ".env" (
    copy ".env.example" ".env" >nul
    echo.
    echo  IMPORTANT: Edit .env and add your API key:
    echo    - ANTHROPIC_API_KEY  (from console.anthropic.com^)
    echo.
    notepad .env
)

echo.
echo  ============================================
echo   Setup complete!
echo  ============================================
echo.
echo  Next steps:
echo   1. Make sure .env has your API keys
echo   2. Run start.bat to launch JARVIS
echo   3. Open https://localhost:5173 in Chrome
echo   4. Click Advanced -> Proceed (cert warning, one time only^)
echo.
pause
