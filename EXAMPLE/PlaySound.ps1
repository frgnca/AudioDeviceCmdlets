<#
Copyright (c) 2016-2022 Francois Gendron <fg@frgn.ca>
MIT License

This file is a script that toggles between two playback audio devices using 
AudioDeviceCmdlets then plays a sound for confirmation
AudioDeviceCmdlets is a suite of PowerShell Cmdlets to control audio devices 
on Windows
https://github.com/frgnca/AudioDeviceCmdlets
#>

# This script toggles between two playback audio devices then plays a sound for confirmation
# The devices are defined by their ID (Get-AudioDevice -List)

# Bonus: Run this PowerShell script from a VBScript to avoid visible window
# Toggle-AudioDevice.vbs
<#
command = "powershell.exe -nologo -command C:\Path\To\PlaySound.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0
#>

# Play sound
$Sound = new-Object System.Media.SoundPlayer
$Sound.SoundLocation = "c:\WINDOWS\Media\Windows Background.wav"
$Sound.Play()
Start-Sleep -s 3
