## Description
AudioDeviceCmdlets is a PowerShell Cmdlet to control audio devices on Windows


## Features  
Get list of audio devices (playback/recording)  
Set default audio device (playback/recording)  
Get volume and mute state of default audio device (playback/recording)  
Set volume and mute state of default audio device (playback/recording)


## Import Cmdlet to PowerShell
Download <a href="https://github.com/frgnca/AudioDeviceCmdlets/raw/master/AudioDeviceCmdlets.dll">AudioDeviceCmdlets.dll</a>
```powershell
New-Item "$($profile | split-path)\Modules\AudioDeviceCmdlets" -Type directory -Force
Copy-Item "C:\Path\to\AudioDeviceCmdlets.dll" "$($profile | split-path)\Modules\AudioDeviceCmdlets\AudioDeviceCmdlets.dll"
Set-Location "$($profile | Split-Path)\Modules\AudioDeviceCmdlets"
Get-ChildItem | Unblock-File
Import-Module AudioDeviceCmdlets
```


## Usage
```PowerShell
Get-AudioDeviceList
Get-AudioDeviceDefaultList
Get-AudioDevicePlaybackList
Get-AudioDeviceRecordingList
```
```PowerShell
Get-AudioDevicePlaybackDefault
Get-AudioDevicePlaybackDefaultMute
Get-AudioDevicePlaybackDefaultVolume
Get-AudioDeviceRecordingDefault
Get-AudioDeviceRecordingDefaultMute
Get-AudioDeviceRecordingDefaultVolume
```
```PowerShell
Set-AudioDevicePlaybackDefault -Index <Int>
Set-AudioDevicePlaybackDefault -Name <String>
Set-AudioDevicePlaybackDefault -InputObject <AudioDevice>
Set-AudioDevicePlaybackDefaultMute <Bool>
Set-AudioDevicePlaybackDefaultMute #Toggle
Set-AudioDevicePlaybackDefaultVolume -Volume <Float>
Set-AudioDeviceRecordingDefault -Index <Int>
Set-AudioDeviceRecordingDefault -Name <String>
Set-AudioDeviceRecordingDefault  -InputObject <AudioDevice>
Set-AudioDeviceRecordingDefaultMute <Bool>
Set-AudioDeviceRecordingDefaultMute #Toggle
Set-AudioDeviceRecordingDefaultVolume -Volume <Float>
```
```PowerShell
Write-DefaultAudioDeviceValue -StreamValue
```


## Build Cmdlet from source

1. Using Visual Studio Community, create new project from SOURCE folder  
    File -> New -> Project From Existing Code...
    
		Type of project: Visual C#
		Folder: SOURCE
		Name: AudioDeviceCmdlets
		Output type: Class Library

2. Install System.Management.Automation NuGet package  
    Project -> Manage NuGet Packages...

		Browse: System.Management.Automation
		Install: v6.3+

3. Set project properties  
	Project -> AudioDeviceCmdlets Properties...

		Assembly name: AudioDeviceCmdlets
		Target framework: .NET Framework 4.5+

4. Set solution configuration  
    Build -> Configuration Manager...

		Active solution configuration: Release

5. Build Cmdlet  
    Build -> Build AudioDeviceCmdlets

        AudioDeviceCmdlets\bin\Release\AudioDeviceCmdlets.dll


## Attribution

Based on code posted to Code Project by Ray Molenkamp with comments and suggestions by MadMidi  
http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control  
Based on code posted to GitHub by Chris Hunt  
https://github.com/cdhunt/WindowsAudioDevice-Powershell-Cmdlet  
