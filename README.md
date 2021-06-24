## Description
AudioDeviceCmdlets is a suite of PowerShell Cmdlets to control audio devices on Windows


## Features
Get list of all audio devices  
Get default audio device (playback/recording)  
Get volume and mute state of default audio device (playback/recording)  
Set default audio device (playback/recording)  
Set volume and mute state of default audio device (playback/recording)  


## Import Cmdlet to PowerShell
Run the script below for a hands-free installation. This will get the latest dll asset, create the module directory and install the module.
```powershell
# Setup
# Checks if the module is installed, if not, will download from v3.0 release and install the module.
function Get-LatestGitHubVersion {
    $repo = "frgnca/AudioDeviceCmdlets"

    $releases = "https://api.github.com/repos/$repo/releases"
    $releases
    [uri]$downloadURL = (Invoke-WebRequest $releases | ConvertFrom-Json)[0].assets[0].browser_download_url

    return $downloadURL
}
if (!(get-module -name AudioDeviceCmdlets)) {
    $modulePath = "$($profile | split-path)\Modules\AudioDeviceCmdlets"
    $fileName = "AudioDeviceCmdlets.dll"
    #dllURL = "https://github.com/frgnca/AudioDeviceCmdlets/releases/download/v3.0/AudioDeviceCmdlets.dll"
    $dllURL = Get-LatestGitHubVersion
    $dllDownloadPath = "$env:USERPROFILE\Downloads\$fileName"
    $dllDestinationPath = "$modulePath\$fileName"
    Invoke-WebRequest -Uri $dllURL -OutFile $dllDownloadPath
    if (!(Test-Path $modulePath)) {
        New-Item $modulePath -Type directory -Force
    }
    Copy-Item  $dllDownloadPath $dllDestinationPath
    Set-Location $modulePath
    Get-ChildItem | Unblock-File
    Import-Module AudioDeviceCmdlets -force
}
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
                -ID <string>              # Sets the default playback/recording device to the device with the ID corresponding to the given <string>
                -Index <int>              # Sets the default playback/recording device to the device with the Index corresponding to the given <int>
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
