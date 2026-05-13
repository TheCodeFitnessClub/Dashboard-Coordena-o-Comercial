@echo off
cd /d "%~dp0"

:: ── Verifica se o proxy já está em execução ──────────────────
netstat -ano | find ":3001" | find "LISTEN" >nul 2>&1
if %errorlevel% == 0 (
    echo   Proxy já está activo.
    goto :open
)

:: ── Inicia o proxy em janela minimizada ──────────────────────
start "Proxy Closum" /min powershell -ExecutionPolicy Bypass -WindowStyle Minimized -File "%~dp0proxy.ps1"

:: ── Aguarda o proxy ficar disponível (máx 8s) ────────────────
setlocal enabledelayedexpansion
for /l %%i in (1,1,8) do (
    timeout /t 1 /nobreak >nul
    netstat -ano | find ":3001" | find "LISTEN" >nul 2>&1
    if !errorlevel! == 0 goto :open
)

:open
:: ── Abre o dashboard no Chrome (ou browser default se Chrome não existir) ──
set "FILE=%~dp0Dashboard_v2.html"
set "CHROME_64=C:\Program Files\Google\Chrome\Application\chrome.exe"
set "CHROME_32=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if exist "%CHROME_64%" (
    start "" "%CHROME_64%" --new-window "%FILE%"
) else if exist "%CHROME_32%" (
    start "" "%CHROME_32%" --new-window "%FILE%"
) else (
    start "" "%FILE%"
)
