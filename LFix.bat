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

:: Kill IE and Outlook
TASKKILL /F /IM OUTLOOK.EXE /T &
TASKKILL /F /IM iexplore.exe /T &

:: Register IE DLLs
regsvr32 /s mshtml.dll 
regsvr32 /s shdocvw.dll 
regsvr32 /s browseui.dll 
regsvr32 /s msjava.dll

REGEDIT /S Links.reg

:: Restore HTML Extensions to default
reg add "HKCU\Software\Classes\.htm" /t REG_SZ /d "htmlfile" /ve /f
reg add "HKCU\Software\Classes\.html" /t REG_SZ /d "htmlfile" /ve /f

:: Set IE as Default for all Web Classes
reg add "HKCU\Software\Classes\http\shell\open\command" /t REG_SZ /d "\"C:\Program Files\Internet Explorer\IEXPLORE.EXE\" -nohome" /ve /f
reg add "HKCU\Software\Classes\ftp\shell\open\command" /t REG_SZ /d "\"C:\Program Files\Internet Explorer\IEXPLORE.EXE\" %%1" /ve /f
reg add "HKCU\Software\Classes\https\shell\open\command" /t REG_SZ /d "\"C:\Program Files\Internet Explorer\IEXPLORE.EXE\" -nohome" /ve /f
reg delete "HKCU\Software\Classes\http\DefaultIcon" /ve /f
reg add "HKCU\Software\Classes\http\DefaultIcon" /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\url.dll,0" /ve /f
reg delete "HKCU\Software\Classes\ftp\DefaultIcon" /ve /f
reg add "HKCU\Software\Classes\ftp\DefaultIcon" /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\url.dll,0" /ve /f
reg delete "HKCU\Software\Classes\https\DefaultIcon" /ve /f
reg add "HKCU\Software\Classes\https\DefaultIcon" /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\url.dll,0" /ve /f
reg add "HKCU\Software\Classes\http\shell\open\ddeexec" /t REG_SZ /d "\"%%1\",,-1,0,,,," /ve /f
reg add "HKCU\Software\Classes\ftp\shell\open\ddeexec" /t REG_SZ /d "\"%%1\",,-1,0,,,," /ve /f
reg add "HKCU\Software\Classes\https\shell\open\ddeexec" /t REG_SZ /d "\"%%1\",,-1,0,,,," /ve /f
reg delete "HKCU\Software\Clients\StartMenuInternet" /ve /f
reg add "HKLM\Software\Clients\StartMenuInternet" /t REG_SZ /d "IEXPLORE.EXE" /ve /f
reg add "HKCR\HTTP\shell\open\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKCR\HTTPS\shell\open\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKCR\FTP\shell\open\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKCR\htmlfile\shell\open\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKCR\htmlfile\shell\opennew\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKCR\mhtmlfile\shell\open\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKCR\mhtmlfile\shell\opennew\ddeexec\Application" /t REG_SZ /d "IExplore" /ve /f
reg add "HKLM\SOFTWARE\Classes\ftp\shell\open\ddeexec\ifExec" /t REG_SZ /d "*" /ve /f

:: Clear Internet Explorer Temp/Cache
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351  
RunDll32.exe InetCpl.cpl,ResetIEtoDefaults 



::Run Outlook
IF exist "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE" SET "OX=C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
IF exist "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE" SET "OX=C:\Program Files (X86)\Microsoft Office\Office16\OUTLOOK.EXE"
IF exist "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE" SET "OX=C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE"
IF exist "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" SET "OX=C:\Program Files (X86)\Microsoft Office\Office15\OUTLOOK.EXE"

TASKKILL /F /IM OUTLOOK.EXE /T
start "" "%OX%" 

EXIT