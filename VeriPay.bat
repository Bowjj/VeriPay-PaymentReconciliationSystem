@echo off
setlocal
title VeriPay - Local
cd /d "%~dp0"

echo ==============================
echo       Starting VeriPay
echo ==============================
echo.

echo [1/3] Checking Docker Desktop...
docker info >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Docker Desktop is not running.
    echo Start Docker Desktop, wait for the engine to be ready, then run this file again.
    echo.
    pause
    exit /b 1
)

echo [2/3] Starting local Supabase...
call supabase start >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Supabase could not start.
    echo.
    pause
    exit /b 1
)

for /f "tokens=1,* delims==" %%a in ('supabase status -o env 2^>nul ^| findstr /b "API_URL="') do set "NEXT_PUBLIC_SUPABASE_URL=%%b"
for /f "tokens=1,* delims==" %%a in ('supabase status -o env 2^>nul ^| findstr /b "PUBLISHABLE_KEY="') do set "NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY=%%b"
set "NEXT_PUBLIC_SUPABASE_URL=%NEXT_PUBLIC_SUPABASE_URL:"=%"
set "NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY=%NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY:"=%"

if not defined NEXT_PUBLIC_SUPABASE_URL (
    echo.
    echo ERROR: Local Supabase URL could not be read.
    echo.
    pause
    exit /b 1
)

if not defined NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY (
    echo.
    echo ERROR: Local Supabase publishable key could not be read.
    echo.
    pause
    exit /b 1
)

for /f "tokens=5" %%p in ('netstat -ano ^| findstr /r /c:":3000 .*LISTENING"') do set "VERIPAY_PORT_PID=%%p"
if defined VERIPAY_PORT_PID (
    echo.
    echo ERROR: Port 3000 is already in use by PID %VERIPAY_PORT_PID%.
    echo Stop that process, then run this file again.
    echo.
    pause
    exit /b 1
)

echo.
echo [3/3] Starting VeriPay...
echo.
echo VeriPay will be available at:
echo http://127.0.0.1:3000
echo.
echo Press CTRL+C when you want to stop VeriPay.
echo.

start "" /b powershell -NoProfile -Command "$deadline = (Get-Date).AddSeconds(60); while ((Get-Date) -lt $deadline) { try { if ((Invoke-WebRequest -UseBasicParsing 'http://127.0.0.1:3000/login' -TimeoutSec 2).StatusCode -eq 200) { Start-Process 'http://127.0.0.1:3000'; exit } } catch {}; Start-Sleep -Milliseconds 500 }"

node node_modules\next\dist\bin\next dev --hostname 127.0.0.1 --port 3000

echo.
echo VeriPay has stopped.
set "NEXT_PUBLIC_SUPABASE_URL="
set "NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY="
pause
