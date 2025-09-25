Option Explicit

Dim objShell, objFSO, strRemove, strInstall, intReturn, strLogTime
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' --- Config ---
Const ISO_DRIVE = "Z:\"
Const REMOVE_BAT  = ISO_DRIVE & "remove.bat"
Const INSTALL_BAT = ISO_DRIVE & "install.bat"

' --- Helper for logging ---
Function Log(msg)
    strLogTime = Now
    WScript.Echo "[" & strLogTime & "] " & msg
End Function

Log "Starting SCCM client uninstall + reinstall from mounted ISO..."

' --- Run remove.bat ---
If objFSO.FileExists(REMOVE_BAT) Then
    Log "Running remove.bat..."
    intReturn = objShell.Run("""" & REMOVE_BAT & """", 0, True)
    If intReturn <> 0 Then
        Log "Warning: remove.bat returned exit code " & intReturn
    End If
    Log "Waiting 60 seconds for uninstall to complete..."
    WScript.Sleep 60000
Else
    Log "ERROR: remove.bat not found on " & REMOVE_BAT
End If

' --- Run install.bat ---
If objFSO.FileExists(INSTALL_BAT) Then
    Log "Running install.bat..."
    intReturn = objShell.Run("""" & INSTALL_BAT & """", 0, True)
    If intReturn = 0 Then
        Log "SCCM client reinstall initiated successfully."
    Else
        Log "ERROR: install.bat returned exit code " & intReturn
        WScript.Quit intReturn
    End If
Else
    Log "ERROR: install.bat not found on " & INSTALL_BAT
    WScript.Quit 1
End If

WScript.Quit 0
