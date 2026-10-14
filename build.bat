@echo off
REM Build script for TTS TF2 64-bit executable
REM Requires Python 3.11+ (64-bit) with PyInstaller

echo ========================================
echo TTS TF2 Build Script
echo ========================================
echo.

REM Check Python installation: C:\Python313 if it's there, otherwise the "python" on PATH
set "PY=C:\Python313\python.exe"
if not exist "%PY%" set "PY=python"
echo Checking Python installation...
"%PY%" --version
if errorlevel 1 (
    echo ERROR: Python 3.11+ 64-bit not found - install it from python.org, or set PY at the top of this script
    pause
    exit /b 1
)

REM Install requirements
echo.
echo Installing requirements...
"%PY%" -m pip install -r requirements.txt

REM The running TTS locks its exe, which would leave dist half-deleted
tasklist /FI "IMAGENAME eq TTS_TF2_64BIT.exe" | find /i "TTS_TF2_64BIT.exe" >nul
if not errorlevel 1 (
    echo.
    echo ERROR: TTS_TF2_64BIT.exe is running. Close it, then build again.
    pause
    exit /b 1
)

REM Keep your settings: everything in dist (config.json, backups, tts.cfg, logs, old exe copies) is set
REM aside in dist_keep before the clean and put back after the build. The old exe only comes back if the
REM build failed, so you still have a working TTS.
echo.
echo Saving your settings from dist...
REM A dist_keep left by an interrupted build may hold the only copy of your settings, so put back whatever
REM dist is missing (never overwriting newer files) before starting over - never just delete it
if exist "dist_keep" (
    echo Found dist_keep from an earlier build - putting it back into dist first
    robocopy "dist_keep" "dist" /E /XO /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >nul
    if errorlevel 8 (
        echo ERROR: Could not put dist_keep back into dist - stopping so nothing is lost
        pause
        exit /b 1
    )
    rmdir /s /q "dist_keep"
)
if exist "dist" robocopy "dist" "dist_keep" /E /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >nul
if exist "dist\config.json" if not exist "dist_keep\config.json" (
    echo.
    echo ERROR: Could not save dist\config.json - stopping so your settings are not lost
    pause
    exit /b 1
)

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build executable
echo.
echo Building executable...
"%PY%" -m PyInstaller --clean --distpath dist --workpath build build_specs\build_64bit_final.spec
set BUILD_RESULT=%errorlevel%

REM Put your settings back (done even if the build failed)
if exist "dist_keep" (
    echo.
    echo Restoring your settings to dist...
    if "%BUILD_RESULT%"=="0" (
        robocopy "dist_keep" "dist" /E /XF TTS_TF2_64BIT.exe /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >nul
    ) else (
        robocopy "dist_keep" "dist" /E /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >nul
    )
    if errorlevel 8 (
        echo WARNING: Could not restore everything - your settings are still in dist_keep
    ) else (
        rmdir /s /q "dist_keep"
    )
)

if not "%BUILD_RESULT%"=="0" (
    echo.
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo Executable: dist\TTS_TF2_64BIT.exe
echo ========================================
pause