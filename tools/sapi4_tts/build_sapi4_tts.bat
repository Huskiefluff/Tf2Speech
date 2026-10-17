@echo off
REM Builds sapi4_tts.exe (32-bit - SAPI 4 voices are 32-bit only) with the Visual Studio C++ build tools.
REM The result is copied to src\ so the TTS build picks it up. Only needed if you change sapi4_tts.cpp.
setlocal
set "VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars32.bat"
if not exist "%VCVARS%" set "VCVARS=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars32.bat"
if not exist "%VCVARS%" (
    echo Visual Studio C++ build tools not found - install "Desktop development with C++" or edit VCVARS here
    exit /b 1
)
call "%VCVARS%" >nul
cd /d "%~dp0"
cl /nologo /O2 /MT /EHsc /W3 sapi4_tts.cpp ole32.lib user32.lib /Fe:sapi4_tts.exe
if errorlevel 1 exit /b 1
copy /y sapi4_tts.exe "..\..\src\sapi4_tts.exe" >nul
del sapi4_tts.obj 2>nul
echo Built sapi4_tts.exe
