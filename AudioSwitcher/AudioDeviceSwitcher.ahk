; Set flags for this script
#Requires AutoHotkey v2.0
#SingleInstance

; Define the hotkey (Ctrl+Alt+A)
^!a::RunPowerShellScript()

;Define the script actions
RunPowerShellScript() {
    Run "powershell.exe -noexit -ExecutionPolicy Bypass -Command `"& { $scriptContent = Get-Content -Path 'Switch-AudioDevices.txt' -Raw; Invoke-Expression -Command $scriptContent }`""
}