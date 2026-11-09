@echo off
rem ==========================================================================
rem  PBIP_Context.cmd
rem  Drop this file into a Power BI Project (.pbip) folder and run it.
rem  It writes PBIP_Context.txt: a folder tree plus the full text of every
rem  readable project file (PBIR report, TMDL/BIM semantic model, etc.) so an
rem  AI assistant can read the whole project as one document.
rem ==========================================================================
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

set "ROOT=%~dp0"
set "OUTNAME=PBIP_Context.txt"
set "OUT=%ROOT%%OUTNAME%"
set "SELF=%~nx0"

rem File types treated as text and included in the context file
set "EXTS=.pbip .pbir .pbism .pbidataset .pbivis .json .tmdl .bim .platform .dax .m .pq .md .txt .xml .yaml .yml"

if exist "%OUT%" del "%OUT%"

set /a COUNT=0
set /a SKIPPED=0

echo Scanning "%ROOT%" ...

> "%OUT%" echo ################################################################
>> "%OUT%" echo # POWER BI PROJECT CONTEXT FILE
>> "%OUT%" echo # Generated: %DATE% %TIME%
>> "%OUT%" echo # Source folder: %ROOT%
>> "%OUT%" echo #
>> "%OUT%" echo # Section 1 is the folder tree. Section 2 contains every text
>> "%OUT%" echo # file in the project, each starting with a "FILE:" header.
>> "%OUT%" echo # Skipped: .pbi cache, .git, binary files, and this script.
>> "%OUT%" echo ################################################################
>> "%OUT%" echo.
>> "%OUT%" echo ======== SECTION 1: FOLDER TREE ========
>> "%OUT%" echo.
pushd "%ROOT%"
>> "%OUT%" tree "%ROOT%" /f /a
popd
>> "%OUT%" echo.
>> "%OUT%" echo ======== SECTION 2: FILE CONTENTS ========

for /r "%ROOT%" %%F in (*) do call :process "%%F"

>> "%OUT%" echo.
>> "%OUT%" echo ======== END OF PROJECT CONTEXT (!COUNT! files included^) ========

echo.
echo Done. Included !COUNT! files, skipped !SKIPPED!.
echo Output: %OUT%
if /i not "%~1"=="/nopause" pause
endlocal
exit /b 0

:process
set "FULL=%~1"
set "NAME=%~nx1"
set "EXT=%~x1"

rem Never include our own output or this script
if /i "!NAME!"=="%OUTNAME%" exit /b 0
if /i "!NAME!"=="%SELF%" exit /b 0

rem Skip cache / VCS / tooling folders
for %%D in (.pbi .git node_modules) do (
    if not "!FULL:\%%D\=!"=="!FULL!" ( set /a SKIPPED+=1 & exit /b 0 )
)

rem Only include whitelisted text extensions
set "OK="
for %%E in (%EXTS%) do if /i "!EXT!"=="%%E" set "OK=1"
if not defined OK ( set /a SKIPPED+=1 & exit /b 0 )

set "REL=!FULL:%ROOT%=!"
>> "%OUT%" echo.
>> "%OUT%" echo ================================================================
>> "%OUT%" echo FILE: !REL!
>> "%OUT%" echo ================================================================
>> "%OUT%" type "!FULL!"
>> "%OUT%" echo.
set /a COUNT+=1
exit /b 0
