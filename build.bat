@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo  VN Pathfinder — Build Script
echo ============================================================
echo.

:: ── Check Python ─────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ and add it to PATH.
    pause
    exit /b 1
)

for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PY_VER=%%v
echo [OK] Python %PY_VER% found.

:: ── Install / upgrade dependencies ──────────────────────────────────────────
echo.
echo [INFO] Installing dependencies...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)
echo [OK] Dependencies installed.

:: ── Clean previous build ─────────────────────────────────────────────────────
echo.
echo [INFO] Cleaning previous build...
if exist dist\VNPathfinder.exe del /f /q dist\VNPathfinder.exe
if exist build rmdir /s /q build
echo [OK] Clean done.

:: ── Build portable EXE ───────────────────────────────────────────────────────
echo.
echo [INFO] Building portable EXE with PyInstaller...
python -m PyInstaller vn_pathfinder.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

if not exist dist\VNPathfinder.exe (
    echo [ERROR] EXE not found after build.
    pause
    exit /b 1
)

for %%F in (dist\VNPathfinder.exe) do set EXE_SIZE=%%~zF
set /a EXE_MB=!EXE_SIZE! / 1048576
echo [OK] dist\VNPathfinder.exe built successfully (!EXE_MB! MB)

:: ── Build installer (Inno Setup) ─────────────────────────────────────────────
echo.
echo [INFO] Looking for Inno Setup...

set ISCC=""
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if %ISCC%=="" (
    echo [SKIP] Inno Setup not found. Skipping installer build.
    echo        Download from: https://jrsoftware.org/isinfo.php
) else (
    echo [INFO] Building installer...
    %ISCC% installer.iss
    if errorlevel 1 (
        echo [ERROR] Inno Setup build failed.
    ) else (
        for %%F in (dist\VNPathfinder_Setup.exe) do set ISS_SIZE=%%~zF
        set /a ISS_MB=!ISS_SIZE! / 1048576
        echo [OK] dist\VNPathfinder_Setup.exe built successfully (!ISS_MB! MB)
    )
)

:: ── Summary ───────────────────────────────────────────────────────────────────
echo.
echo ============================================================
echo  Build complete!
echo ============================================================
echo.
echo  Outputs in dist\:
if exist dist\VNPathfinder.exe       echo    VNPathfinder.exe          (portable)
if exist dist\VNPathfinder_Setup.exe echo    VNPathfinder_Setup.exe    (installer)
echo.
echo  Upload both to GitHub Releases, or run the EXE directly.
echo.
pause
