@echo off
::***************************************************************************
:: PhotoDemon Build Script
:: Copyright 2019 by wqweto@gmail.com
:: Created: 25/October/19
::
:: All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
::  projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
::
::***************************************************************************
setlocal
set "VbCodeLines=%~dp0\VbCodeLines.exe"
set "Vb6=%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist "%Vb6%" set "Vb6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"

for %%i in ("%~dp0Temp") do set "temp_dir=%%~dpnxi"
for %%i in ("%~dp0..\..\.") do set "src_dir=%%~dpnxi"
set "log_file=%temp_dir%\compile.out"
set "out_dir=%temp_dir%"
if not "%~1"=="" set "out_dir=%~1"

echo Cleanup %temp_dir%...
rd /s /q "%temp_dir%" 2>&1

echo Copy sources from %src_dir%...
xcopy /q /y "%src_dir%\PhotoDemon.vbp" "%temp_dir%\" 2>&1 > nul
xcopy /q /y /s "%src_dir%\Classes" "%temp_dir%\Classes\" 2>&1 > nul
xcopy /q /y /s "%src_dir%\Controls" "%temp_dir%\Controls\" 2>&1 > nul
xcopy /q /y /s "%src_dir%\Forms" "%temp_dir%\Forms\" 2>&1 > nul
xcopy /q /y /s "%src_dir%\Interfaces" "%temp_dir%\Interfaces\" 2>&1 > nul
xcopy /q /y /s "%src_dir%\Modules" "%temp_dir%\Modules\" 2>&1 > nul
xcopy /q /y /s "%src_dir%\Resources" "%temp_dir%\Resources\" 2>&1 > nul

echo Compiling to %out_dir%...
for %%i in ("%temp_dir%\*.vbp") do (
    del "%log_file%" > nul 2>&1
    start "" /w "%Vb6%" /make "%%i" /out "%log_file%" /outdir "%out_dir%"
    findstr /r /c:"Build of '.*' succeeded" "%log_file%" || (type "%log_file%" 1>&2 & exit /b 1)
)
echo Done.
