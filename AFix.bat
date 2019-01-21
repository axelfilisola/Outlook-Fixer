@echo off
:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
:-------------------------------------- 

@ECHO OFF
SET "REGY=15.0"
SET "REGX=O15.reg"

IF exist "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE" GOTO O1664
IF exist "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE" GOTO O1632 
IF exist "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE" GOTO  O1364
IF exist "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" GOTO O1332
EXIT

:O1664
    ECHO "Outlook 2016 (64-Bit) Version detected."
    SET "REGY=16.0"
    SET "REGX=O16.reg"
    SET "OX=C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
    GOTO aFixP
EXIT

:O1632{
    ECHO "Outlook 2016 (32-Bit) Version detected."
    SET "REGY =16.0"
    SET "REGX =O16.reg"
    SET "OX=C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
    GOTO aFixP
EXIT

:O1364 
    ECHO "Outlook 2013 (64-Bit) Version detected."
    SET "OX=C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE"
    GOTO aFixP
EXIT

:O1332 
    ECHO "Outlook 2013 (32-Bit) Version detected."
    SET "OX=C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE"
    GOTO aFixP
EXIT

:aFixP
TASKKILL /F /IM OUTLOOK.EXE /T
REG DELETE HKCU\Software\Microsoft\Office\%REGY%\Outlook\Resiliency\DisabledItems /va /f
REG DELETE HKCU\Software\Microsoft\Office\%REGY%\Outlook\Resiliency\CrashingAddinList /va /f
REGEDIT /S %REGX%
echo "%OX%"
TASKKILL /F /IM OUTLOOK.EXE /T
start "" "%OX%" 
EXIT


