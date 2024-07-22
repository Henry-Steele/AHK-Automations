# Audio Output Switcher

## Abstract
The onus of this project came about as a desire to make better use of a pair of speakers I purchased and set up months ago, but rarely used. Given how abhorrent the steps are to simply switch between audio device outputs with Windows 10/11, I took it upon myself to create a new solution that would work for myself, while also being replicable for others. The inspiration for this dynamic switching functionality came about via popular softwares like Razer Synapse, which allows THX Spatial Audio Devices to dynamically swap with a macro press—as well as from tools like Loupdecks/Streamdecks which allow swapping via a simple button.

It should be noted that while I did not directly make use of any code or AHK Scripts from the following, I did take insperation from this project as to how the PoSh App is called by AHK:
https://github.com/samfisherirl/AHKv2-with-integration-for-500-popular-Powershell-scripts-Generator?tab=readme-ov-file

Lastly, a huge thanks to the folks who have contributed to the 'AudioDeviceCmdlets' PowerShell Module. This project would not have been possible without their work on the following project. Please show them some love here: https://github.com/samfisherirl/AHKv2-with-integration-for-500-popular-Powershell-scripts-Generator?tab=readme-ov-file

## License

This project is licensed under the terms of the GNU General Public License v3.0.

You should have received a copy of the GNU General Public License along with this project. If not, see [https://www.gnu.org/licenses/](https://www.gnu.org/licenses/).

## Attribution

This project was created by Henry Steele. In accordance with GPL v3 Licensing, all use of the materials contained within this project is subject to attributation referencing the author. Additionally, any works derived from materials contained herein are to be licensed under the GPL v3 or under future GPL iterations.

## How It Works
- The PowerShell script provides a GUI to switch audio devices. It hides the PowerShell console window and displays a graphical interface to the user.
- The AutoHotkey script sets a hotkey (Ctrl+Alt+A) to run the PowerShell script. When the hotkey is pressed, the PowerShell script is executed, and the GUI is displayed to the user.
- The GUI comes prefilled with all availible Audio Output devices in the dropdown menu. Additionally the current default device will be shown as the preselected option in the dropdown and if it is the default communication device then that checkbox will be preselected as well.
- The user can select an audio output device from the dropdown list and choose whether to set it as the default device, the default communication device, or both.
- In cases where an audio device is connected after the GUI is launched for whatever reason, a UI element in the form of a button with the label `Refresh` has been included as to reload the device list without forcing the end user to close and reload the app.

## Prerequisites
To set up and run this project, you will need the following prerequisites:
- Windows operating system
- PowerShell
- AutoHotkey v2.0

## Installation and Setup

### 1. Install PowerShell Module
The project uses the `AudioDeviceCmdlets` PowerShell module to manage audio devices. You can install it using the following command:

```powershell
Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser
```

### 2. Save the PowerShell Script
Save the following PowerShell script to a file named `Switch-AudioDevices.txt`. This file should reside in the same directory as the .AHK script we are creating in the step below. For my lazy friends a copy of this file has been included within the project contents.

 Note: The codeblock below and the .txt file located within this project are not the main source for this code. Please reference https://github.com/Henry-Steele/PowerShell-Tools for up to date versions

```powershell
# Load necessary assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Hide the PowerShell console window
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
}
"@
$consolePtr = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($consolePtr, 0) # 0 = SW_HIDE

# Check if the module is installed
if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
    $response = [System.Windows.MessageBox]::Show("The AudioDeviceCmdlets module is not installed. Would you like to install it now?", "Module Not Found", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($response -eq [System.Windows.MessageBoxResult]::Yes) {
        Install-Module -Name AudioDeviceCmdlets -Scope CurrentUser -Force | Out-Null
        if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
            [System.Windows.MessageBox]::Show("Failed to install AudioDeviceCmdlets module. Exiting.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            exit
        }
    } else {
        [System.Windows.MessageBox]::Show("AudioDeviceCmdlets module is required. Exiting.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        exit
    }
}

Import-Module AudioDeviceCmdlets | Out-Null

function Get-AudioDevices {
    Get-AudioDevice -List | Where-Object { $_.Type -eq 'Playback' }
}

function Get-DefaultAudioDevice {
    (Get-AudioDevice -Playback).Name
}

function Get-DefaultCommunicationDevice {
    (Get-AudioDevice -PlaybackCommunication).Name
}

function Set-DefaultAudioDevice {
    param (
        [int]$deviceIndex
    )
    Set-AudioDevice -Index $deviceIndex -DefaultOnly | Out-Null
}

function Set-DefaultCommunicationDevice {
    param (
        [int]$deviceIndex
    )
    Set-AudioDevice -Index $deviceIndex -CommunicationOnly | Out-Null
}

# Initialize GUI components
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Audio Device Switcher" Height="300" Width="400">
    <Grid>
        <StackPanel Margin="10">
            <TextBlock Text="Select Audio Output Device:" Margin="0,0,0,10"/>
            <ComboBox Name="DeviceComboBox" Width="300"/>
            <CheckBox Name="DefaultDeviceCheckBox" Content="Set as Default Device" Margin="0,10,0,0"/>
            <CheckBox Name="CommDeviceCheckBox" Content="Set as Default Communication Device" Margin="0,10,0,0"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
                <Button Name="RefreshButton" Content="Refresh" Width="75" Margin="0,0,10,0"/>
                <Button Name="ConfirmButton" Content="Switch" Width="75"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$DeviceComboBox = $Window.FindName("DeviceComboBox")
$DefaultDeviceCheckBox = $Window.FindName("DefaultDeviceCheckBox")
$CommDeviceCheckBox = $Window.FindName("CommDeviceCheckBox")
$RefreshButton = $Window.FindName("RefreshButton")
$ConfirmButton = $Window.FindName("ConfirmButton")

# Store device index mapping
$global:deviceIndexMap = @{}

function Refresh-DeviceList {
    try {
        $DeviceComboBox.Items.Clear()
        $global:deviceIndexMap.Clear()
        $devices = Get-AudioDevices
        if ($devices.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No audio devices found.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        foreach ($device in $devices) {
            $DeviceComboBox.Items.Add($device.Name)
            $global:deviceIndexMap[$device.Name] = $device.Index
        }
        $defaultDevice = Get-DefaultAudioDevice
        $defaultCommDevice = Get-DefaultCommunicationDevice
        $DeviceComboBox.SelectedItem = $defaultDevice
        $DefaultDeviceCheckBox.IsChecked = $true

        if ($defaultDevice -eq $defaultCommDevice) {
            $CommDeviceCheckBox.IsChecked = $true
        } else {
            $CommDeviceCheckBox.IsChecked = $false
        }

    } catch {
        [System.Windows.MessageBox]::Show("Error retrieving audio devices: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

$RefreshButton.Add_Click({ Refresh-DeviceList | Out-Null })
$ConfirmButton.Add_Click({
    try {
        if (-not $DefaultDeviceCheckBox.IsChecked -and -not $CommDeviceCheckBox.IsChecked) {
            [System.Windows.MessageBox]::Show("Please select at least one option: Default Device or Default Communication Device.", "Selection Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        $selectedDeviceName = $DeviceComboBox.SelectedItem
        $deviceIndex = $global:deviceIndexMap[$selectedDeviceName]

        if ($DefaultDeviceCheckBox.IsChecked -eq $true) {
            Set-DefaultAudioDevice -deviceIndex $deviceIndex
        }
        
        if ($CommDeviceCheckBox.IsChecked -eq $true) {
            Set-DefaultCommunicationDevice -deviceIndex $deviceIndex
        }

        $Window.Close()
    } catch {
        [System.Windows.MessageBox]::Show("Error setting audio device: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Initial population of devices
$null = Refresh-DeviceList

# Show the window
$null = $Window.ShowDialog()

# Close PowerShell window after script execution
Stop-Process -Id $PID
```

### 3. Create the AutoHotkey Script
Save the following AutoHotkey script to a file named `Switch-AudioDevices.ahk`. This file should live in the same directory as the PoSh saved to a text file in the last step. For my lazy friends, I have included a copy of this file in the project for your convenience. All you need to do is download it and proceed to step 3. 

Additionally, to customize the macro used to run this AHK script swap `^!a` for any other combination of characters. For reference please see: https://www.autohotkey.com/docs/v1/KeyList.htm#toc

```ahk
; Set flags for this script
#Requires AutoHotkey v2.0
#SingleInstance

; Define the hotkey (Ctrl+Alt+A)
^!a::RunPowerShellScript()

;Define the script actions
RunPowerShellScript() {
    Run "powershell.exe -noexit -ExecutionPolicy Bypass -Command `"& { $scriptContent = Get-Content -Path 'Switch-AudioDevices.txt' -Raw; Invoke-Expression -Command $scriptContent }`""
}
```

### 4. Set Up the Project
1. Make sure you have both the PowerShell script (`Switch-AudioDevices.txt`) and the AutoHotkey script (`Switch-AudioDevices.ahk`) saved in the same directory.
2. Install the `AudioDeviceCmdlets` PowerShell module if you haven't already.
3. Install AutoHotkey v2.0 from [here](https://www.autohotkey.com/).

### 5. Autorun the Script
1. Create a shortcut corresponding the the location of `Switch-AudioDevices.ahk`
2. Copy the shortcut to `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp`
3. To test usage doubleclick the shortcut to your AHK script. This will launch the script in the background. From here you can press `Ctrl + Alt + A` to launch the PoSh on an ad hoc basis whenever you want to swap audio devices.
