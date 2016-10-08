Basic command-line audio device control from Powershell including Nuget Package Manager Console.

Features: Set Volume and toggle Mute on the Default Playback Device. Get a list of devices and set the Default Audio Device.

## Install.

### From WMF5+

    Install-Module -Name AudioDeviceCmdlets
    
### Manual Install

* Download https://github.com/cdhunt/WindowsAudioDevice-Powershell-Cmdlet/blob/master/AudioDeviceCmdlets.zip
* `New-Item "$($profile | split-path)\Modules\AudioDeviceCmdlets" -Type directory -Force`
* Copy CoreAudioApi.dll, AudioDeviceCmdlets.dll and AudioDeviceCmdlets.dll-Help.xml to the above location.
* Open a PowerShell console *As Administrator*.

```powershell
Set-Location "$($profile | Split-Path)\Modules\AudioDeviceCmdlets"
Get-ChildItem | Unblock-File
```

* Import the binary module. This can go into your profile.
        
```powershell
Import-Module AudioDeviceCmdlets
```

* You may need to set the execution policy.

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
```

## Suggested Aliases. I may set these in the module in the future.

```powershell
New-Alias -Name Mute -Value Set-DefaultAudioDeviceMute
New-Alias -Name Vol -Value set-DefaultAudioDeviceVolume
```

## Exposed Cmdlets
* Get-DefaultAudioDevice
* Get-AudioDeviceList
* Set-DefaultAudioDevice [-Index] &lt;Int&gt;
* Set-DefaultAudioDevice [-Name] &lt;String&gt;
* Set-DefaultAudioDevice [-InputObject] &lt;AudioDevice&gt;
* Set-DefaultAudioDeviceVolume -Volume &lt;float&gt;
* Get-DefaultAudioDeviceVolume
* Set-DefaultAudioDeviceMute [-Force] &lt;Bool&gt;
* Write-DefaultAudioDeviceValue [-StreamValue]

## Attribution
Based on work done by Ray M. <a href="http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control">hosted</a> on The Code Project
