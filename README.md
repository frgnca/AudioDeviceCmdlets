## Description
AudioDeviceCmdlets is a suite of PowerShell Cmdlets to control audio devices on Windows


## Features
Get list of all audio devices  
Get default audio device (playback/recording)  
Get volume and mute state of default audio device (playback/recording)  
Set default audio device (playback/recording)  
Set volume and mute state of default audio device (playback/recording)  


## Installation
Run as administrator
```PowerShell
Install-Module -Name AudioDeviceCmdlets
```


## Usage
```PowerShell
Get-AudioDevice -List             # Outputs a list of all devices as <AudioDevice>
                -ID <string>      # Outputs the device with the ID corresponding to the given <string>
                -Index <int>      # Outputs the device with the Index corresponding to the given <int>
                -Playback         # Outputs the default playback device as <AudioDevice>
                -PlaybackMute     # Outputs the default playback device's mute state as <bool>
                -PlaybackVolume   # Outputs the default playback device's volume level on 100 as <float>
                -Recording        # Outputs the default recording device as <AudioDevice>
                -RecordingMute    # Outputs the default recording device's mute state as <bool>
                -RecordingVolume  # Outputs the default recording device's volume level on 100 as <float>
```
```PowerShell
Set-AudioDevice <AudioDevice>             # Sets the default playback/recording device to the given <AudioDevice>, can be piped
                    -PlaybackOnly         # Only set default playback device, not communication device
                -ID <string>              # Sets the default playback/recording device to the device with the ID corresponding to the given <string>
                    -PlaybackOnly         # Only set default playback device, not communication device
                -Index <int>              # Sets the default playback/recording device to the device with the Index corresponding to the given <int>
                    -PlaybackOnly         # Only set default playback device, not communication device
                -PlaybackMute <bool>      # Sets the default playback device's mute state to the given <bool>
                -PlaybackMuteToggle       # Toggles the default playback device's mute state
                -PlaybackVolume <float>   # Sets the default playback device's volume level on 100 to the given <float>
                -RecordingMute <bool>     # Sets the default recording device's mute state to the given <bool>
                -RecordingMuteToggle      # Toggles the default recording device's mute state
                -RecordingVolume <float>  # Sets the default recording device's volume level on 100 to the given <float>
```
```PowerShell
Write-AudioDevice -PlaybackMeter  # Writes the default playback device's power output on 100 as a meter
                  -PlaybackSteam  # Writes the default playback device's power output on 100 as a stream of <int>
                  -RecordingMeter # Writes the default recording device's power output on 100 as a meter
                  -RecordingSteam # Writes the default recording device's power output on 100 as a stream of <int>
```


## Build Cmdlet from source

1. Using Visual Studio Community, create new project from SOURCE folder  
    File -> New -> Project From Existing Code...
    
		Type of project: Visual C#
		Folder: SOURCE
		Name: AudioDeviceCmdlets
		Output type: Class Library

2. Install System.Management.Automation NuGet legacy package, which is packaged as part of Microsoft.PowerShell.5.1.ReferenceAssemblies  
    Project -> Manage NuGet Packages...

		Browse: Microsoft.PowerShell.5.1.ReferenceAssemblies
		Install: v1.0.0+

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

6. Import Cmdlet to PowerShell
	```PowerShell
	New-Item "$($profile | split-path)\Modules\AudioDeviceCmdlets" -Type directory -Force
	Copy-Item "C:\Path\to\AudioDeviceCmdlets.dll" "$($profile | split-path)\Modules\AudioDeviceCmdlets\AudioDeviceCmdlets.dll"
	Set-Location "$($profile | Split-Path)\Modules\AudioDeviceCmdlets"
	Get-ChildItem | Unblock-File
	Import-Module AudioDeviceCmdlets
	```


## Attribution

Based on code posted to Code Project by Ray Molenkamp with comments and suggestions by MadMidi  
http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control  
Based on code posted to GitHub by Chris Hunt  
https://github.com/cdhunt/WindowsAudioDevice-Powershell-Cmdlet  
