@echo off
rem ============================================================
rem  MoActivation (C version) build script - MSVC (cl + rc)
rem  Output: moactivation.exe
rem ============================================================
setlocal

where cl.exe >nul 2>&1
if not errorlevel 1 goto :build

rem ---- Locate vcvars64.bat (Community/Professional/Enterprise/BuildTools) ----
set "VSROOT="
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2022\Community"
if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2022\Professional"
if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2022\Enterprise"
if exist "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2022\BuildTools"
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files (x86)\Microsoft Visual Studio\2022\Community"
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools"
if exist "C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\18\Community"
if exist "C:\Program Files\Microsoft Visual Studio\18\Professional\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\18\Professional"
if exist "C:\Program Files\Microsoft Visual Studio\18\Enterprise\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\18\Enterprise"
if exist "C:\Program Files\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\18\BuildTools"
if exist "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools"
if exist "C:\Program Files (x86)\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files (x86)\Microsoft Visual Studio\18\Community"
if exist "C:\Program Files\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2019\Community"
if exist "C:\Program Files\Microsoft Visual Studio\2019\Professional\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2019\Professional"
if exist "C:\Program Files\Microsoft Visual Studio\2019\Enterprise\VC\Auxiliary\Build\vcvars64.bat" set "VSROOT=C:\Program Files\Microsoft Visual Studio\2019\Enterprise"
if not defined VSROOT goto :nocl
call "%VSROOT%\VC\Auxiliary\Build\vcvars64.bat" >nul

:build
where cl.exe >nul 2>&1
if errorlevel 1 goto :nocl

rem ---- Embed manifest (admin + PerMonitorV2 DPI awareness) ----
rc.exe /nologo app.rc || goto :err

rem ---- Compile and link (/utf-8 source, /O1 size, /MD shared CRT) ----
cl.exe /nologo /O1 /GL /MD /GS- /utf-8 /W3 /Fe:moactivation.exe ^
    moactivation.c app.res ^
    /link /SUBSYSTEM:WINDOWS /LTCG /ENTRY:wWinMainCRTStartup /OPT:REF /OPT:ICF || goto :err

rem ---- Cleanup intermediate files ----
if exist moactivation.obj del moactivation.obj >nul 2>&1
if exist moactivation.exe.manifest del moactivation.exe.manifest >nul 2>&1

for %%A in (moactivation.exe) do @echo OK: moactivation.exe size=%%~zA bytes
goto :eof

:nocl
echo [ERROR] cl.exe not found. Run this script inside x64 Native Tools Command Prompt.
goto :err

:err
echo BUILD FAILED!
exit /b 1