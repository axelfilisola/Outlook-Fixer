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

TASKKILL /F /IM OUTLOOK.EXE /T
TASKKILL /F /IM EXCEL.EXE /T
TASKKILL /F /IM WINWORD.EXE /T
TASKKILL /F /IM LYNC.EXE /T
TASKKILL /F /IM VISIO.EXE /T
TASKKILL /F /IM POWERPNT.EXE /T
TASKKILL /F /IM SKYPE.EXE /T
TASKKILL /F /IM ONENOTE.EXE /T
TASKKILL /F /IM ONENOTEM.EXE /T

NET stop WSearch

RMDIR /S /Q %LOCALAPPDATA%\Microsoft\Office.old2
MOVE %LOCALAPPDATA%\Microsoft\Office.old %LOCALAPPDATA%\Microsoft\Office.old2
MOVE %LOCALAPPDATA%\Microsoft\Office %LOCALAPPDATA%\Microsoft\Office.old
REG EXPORT HKCU\SOFTWARE\Microsoft\Office\15.0  %LOCALAPPDATA%\Microsoft\Office.old.reg /Y
REG DELETE HKCU\SOFTWARE\Microsoft\Office\16.0 /va /f
REG EXPORT HKCU\SOFTWARE\Microsoft\Office\16.0  %LOCALAPPDATA%\Microsoft\Office2.old.reg /Y
REG DELETE HKCU\SOFTWARE\Microsoft\Office\16.0 /va /f

NET start WSearch

::Run Outlook
IF exist "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE" SET "OX=C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
IF exist "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE" SET "OX=C:\Program Files (X86)\Microsoft Office\Office16\OUTLOOK.EXE"
IF exist "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE" SET "OX=C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE"
IF exist "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" SET "OX=C:\Program Files (X86)\Microsoft Office\Office15\OUTLOOK.EXE"

AFix.bat 


