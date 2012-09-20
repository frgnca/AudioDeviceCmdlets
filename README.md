
## Install
    Import-Module AudioDeviceCmdlets.dll

## Suggested Aliases. I may set these in the module in the future.
    New-Alias -Name mute -Value Set-DefaultAudioDeviceMute
    New-Alias -Name vol -Value set-DefaultAudioDeviceVolume

## Exposed Cmdlets
* Get-DefaultAudioDevice
* Get-AudioDeviceList
* Set-DefaultAudioDevice -Index <int>
* Set-DefaultAudioDeviceVolume -Volume <float>
* Set-DefaultAudioDeviceMute
* Write-DefaultAudioDeviceValue


Based on work done by Ray M. <a href="http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control">hosted</a> on The Code Project
