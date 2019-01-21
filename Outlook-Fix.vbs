
aCheck = MsgBox ("Welcome to Outlook Fixer" & vbcrlf & vbcrlf & "Keep in mind that all the fixes require to restart Outlook." & vbcrlf & "And if Outlook fails to reload, don't worry!" & vbcrlf & "Just start it manualy" & vbcrlf & vbcrlf & "Let's see..." & vbcrlf &  "Are you having issues with Addins not loading or disappearing?", vbYesNo, "Outlook Fixer")
Select Case aCheck
Case vbYes
        Act1="AFIX.BAT"
Case vbNo
        Act1="EXIT.BAT"
End Select


lCheck = MsgBox ("Or..." & vbcrlf & vbcrlf & "Are you having issues opening these elements from your emails?" & vbcrlf & vbcrlf &  "- Hyperlinks" & vbcrlf & "- Calendar Items" & vbcrlf & "- Skype Meeting Invitation", vbYesNo, "Outlook Fixer")
Select Case lCheck
Case vbYes
       Act2="LFIX.BAT"
    Case vbNo
       Act2="EXIT.BAT"
End Select

bCheck = MsgBox ("Is Outlook crashing or unable to open?      :(" & vbcrlf & vbcrlf & "Psst... This might also resolve crashing issues on other Microsoft Office application.    ;)", vbYesNo, "Outlook Fixer")
Select Case bCheck
Case vbYes
        zCheck = MsgBox ("Warning!" & vbcrlf & vbcrlf & " The next operation will close all the MS Office applications (including One Note and Skype)" & vbcrlf & "All of the settings for all MS Office applications will be Restored to Default." & vbcrlf & vbcrlf & "Please save all your files from Word, Excel, Outlook, Powerpoint or any MS Office application." & vbcrlf & vbcrlf & "RESTORE PROCES::" & vbcrlf & "Rename the folder AppData\Local\Microsoft\Office to Office.b" & vbcrlf & "Rename the folder AppData\Local\Microsot\Office.old to Office" & vbcrlf & "Run as admin AppData\Local\Microsoft\Office.old.reg and Office2.old.reg", vbYesNo, "Outlook Fixer")
            Select Case zCheck
            Case vbYes
                Act1="EXIT.BAT"
                Act3="BFix.bat"
            Case vbNo
                MsgBox ("This fix has been cancelled.")
                Act3="EXIT.BAT"
            End Select
Case vbNo
    Act3="EXIT.BAT"
End Select

runAction()
    
Function runAction()
dim shell
set shell=createobject("wscript.shell")
Set oShell = CreateObject("Shell.Application")
    oShell.ShellExecute Act1
    WScript.Sleep 5569
    oShell.ShellExecute Act2
    WScript.Sleep 9169
    oShell.ShellExecute Act3
set shell=nothing
End Function
