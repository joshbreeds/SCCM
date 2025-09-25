Option Explicit

Dim objShell, objFSO, strInstall, strArgs, intReturn, strLogTime
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' --- Config: update these values ---
Const MP       = "SCCMSERVER.domain.local"
Const SITECODE = "ABC"
Const ISO_DRIVE = "Z:\"
Const LOCAL_INSTALLER = ISO_DRIVE & "install.bat"   ' your fixed batch file on the ISO

' --- Helper for logging ---
Function Log(msg)
    strLogTime = Now
    WScript.Echo "[" & strLogTime & "] " & msg
End Function

Log "Starting SCCM client reinstall from mounted ISO..."

' --- Check that install.bat exists on the ISO ---
If objFSO.FileExists(LOCAL_INSTALLER) Then
    Log "Running install.bat from " & LOCAL_INSTALLER
    intReturn = objShell.Run("""" & LOCAL_INSTALLER & """", 0, True)

    If intReturn = 0 Then
        Log "SCCM client install initiated successfully."
    Else
        Log "SCCM client install failed with exit code " & intReturn
        WScript.Quit intReturn
    End If
Else
    Log "ERROR: install.bat not found on " & LOCAL_INSTALLER
    WScript.Quit 1
End If

WScript.Quit 0
