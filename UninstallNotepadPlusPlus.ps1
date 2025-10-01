# Notepad++ Cleanup Script
# Run as SYSTEM via SCCM

$ErrorActionPreference = "SilentlyContinue"

# 1. Stop any running Notepad++ processes
Write-Output "Stopping Notepad++ processes..."
Get-Process "notepad++" -ErrorAction SilentlyContinue | Stop-Process -Force

# 2. Uninstall any installed Notepad++ versions
Write-Output "Looking for Notepad++ uninstall entries..."
$UninstKeys = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
              "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

foreach ($key in $UninstKeys) {
    Get-ChildItem $key | ForEach-Object {
        $dispName = (Get-ItemProperty $_.PSPath).DisplayName
        if ($dispName -like "Notepad++*") {
            $uninst = (Get-ItemProperty $_.PSPath).UninstallString
            if ($uninst) {
                Write-Output "Uninstalling $dispName ..."
                Start-Process "cmd.exe" -ArgumentList "/c $uninst /S" -Wait -NoNewWindow
            }
        }
    }
}

# 3. Remove leftover folders
$Paths = @(
    "C:\Program Files\Notepad++",
    "C:\Program Files (x86)\Notepad++"
)

foreach ($p in $Paths) {
    if (Test-Path $p) {
        Write-Output "Removing leftover folder $p"
        Remove-Item -Path $p -Recurse -Force
    }
}

Write-Output "Cleanup complete."
exit 0
