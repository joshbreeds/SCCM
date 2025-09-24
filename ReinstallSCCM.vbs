Option Explicit

Dim objShell, objFSO, strCCMSetup, strArgs, intReturn, strLogTime
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' --- Config: update these values ---
Const MP       = "SCCMSERVER.domain.local"
Const SITECODE = "ABC"

' --- Local ccmsetup path ---
strCCMSetup = objShell.ExpandEnvironmentStrings("%windir%") & "\ccmsetup\ccmsetup.exe"

' --- Helper function for timestamped logging ---
Function Log(msg)
    strLogTime = Now
    WScript.Echo "[" & strLogTime & "] " & msg
End Function

Log "Starting SCCM client reinstall..."

' --- Uninstall if ccmsetup exists ---
If objFSO.FileExists(strCCMSetup) Then
    Log "Uninstalling SCCM client..."
    objShell.Run """" & strCCMSetup & """ /uninstall", 0, True
    Log "Waiting 60 seconds for uninstall to complete..."
    WScript.Sleep 60000
Else
    Log "No existing ccmsetup.exe found, skipping uninstall."
End If

' --- Reinstall from local ccmsetup.exe ---
If objFSO.FileExists(strCCMSetup) Then
    strArgs = "/mp:" & MP & " SMSSITECODE=" & SITECODE & " /logon"
    Log "Installing SCCM client from local source..."
    intReturn = objShell.Run("""" & strCCMSetup & """ " & strArgs, 0, True)

    If intReturn = 0 Then
        Log "SCCM client reinstall initiated successfully."
    Else
        Log "SCCM client reinstall failed with exit code " & intReturn
        WScript.Quit intReturn
    End If
Else
    Log "ERROR: Local ccmsetup.exe not found."
    WScript.Quit 1
End If

WScript.Quit 0
