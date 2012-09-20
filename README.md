
PS C:\>Import-Module AudioDeviceCmdlets.dll

Get-DefaultAudioDevice
Get-AudioDeviceList
Set-DefaultAudioDevice -Index <int>
Set-DefaultAudioDeviceVolume -Volume <float>

Based on work done by Ray M. <a href="http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control">hosted</a> on The Code Project
