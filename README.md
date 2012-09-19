Import-Module AudioDeviceCmdlets.dll

Get-DefaultAudioDevice
Get-AudioDeviceList
Set-DefaultAudioDevice -Index <int>

Based on http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control