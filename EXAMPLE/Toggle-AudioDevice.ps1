<#
Copyright (c) 2016-2017 Francois Gendron <fg@frgn.ca>
The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

# AudioDeviceCmdlets is a suite of PowerShell Cmdlets to control audio devices on Windows
# https://github.com/frgnca/AudioDeviceCmdlets

# This script toggles between two playback audio devices then plays a sound for confirmation
# The devices are defined by their ID (Get-AudioDevice -List)

# Bonus: Run this PowerShell script from a VBScript to avoid visible window
# Toggle-AudioDevice.vbs
<#
command = "powershell.exe -nologo -command C:\Path\To\Toggle-AudioDevice.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0
#>

# Define AudioDevice by ID (ex: "{0.0.0.00000000}.{c4aadd95-74c7-4b3b-9508-b0ef36ff71ba}")
$AudioDevice_A = "{0.0.0.00000000}.{48300fc4-2125-492d-ab28-c6b01b9eee6b}"
$AudioDevice_B = "{0.0.0.00000000}.{c4aadd95-74c7-4b3b-9508-b0ef36ff71ba}"

# Toggle default playback device
$DefaultPlayback = Get-AudioDevice -Playback
If ($DefaultPlayback.ID -eq $AudioDevice_A) {Set-AudioDevice -ID $AudioDevice_B | Out-Null}
Else {Set-AudioDevice -ID $AudioDevice_A | Out-Null}

# Play sound
$Sound = new-Object System.Media.SoundPlayer
$Sound.SoundLocation = "c:\WINDOWS\Media\Windows Background.wav"
$Sound.Play()
Start-Sleep -s 3
