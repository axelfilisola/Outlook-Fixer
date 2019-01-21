# Outlook-FIX

Fixes:
- Missing Let's Meet AddIn
- Missing Skype for Business AddIn
- Links not opening due to restrictions by Administrator

What it does:
- Removes all the Items from HKCU\Software\Microsoft\Office\Outlook\Resilency\DisabledItems
- Removes all the Items from HKCU\Software\Microsoft\Office\Outlook\Resilency\CrashingAddinList
- Adjust registry keys of Let's Meet and Skype Business to LoadBehavior=3
- Register DLLs : mshtml.dll, shdocvw.dll, browseui.dll and msjava.dll
- Clear Temporary Files from Internet Explorer
- Restores Internet Options to Default
- Set Internet Explorer and Outlook as Default Application

To do:
- Support for 64-Bit version
- Support for Office 2016
