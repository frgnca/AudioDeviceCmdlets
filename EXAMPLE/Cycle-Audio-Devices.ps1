# This script cycles through all audio devices


# Get a list of all enabled devices
[System.Collections.ArrayList]$audioDevicesList = Get-AudioDevice -List

# Get the default playback device as <AudioDevice>
$defaultPlaybackDevice = Get-AudioDevice -Playback

#Remove non-playback devices from audioDevicesList
for ( $index = 0; $index -lt $audioDevicesList.count; $index++ ) {
   if ($audioDevicesList[$index].type -ne "Playback") {
      $audioDevicesList.Remove($audioDevicesList[$index])
   }
}

#Loop through audioDevicesList
for ( $index = 0; $index -lt $audioDevicesList.count; $index++ ) {
   #Find the position of the default audio device
   if ( $audioDevicesList[$index].name -eq $defaultPlaybackDevice.name) {
      #If it's the last device set the first audio device as the default
      if ( $index -eq $audioDevicesList.count - 1) {
         Set-AudioDevice -ID $audioDevicesList[0].id
      }
      #Otherwhise set the next audio device on the list as default
      else {
         Set-AudioDevice -ID $audioDevicesList[$index + 1].id
      }
   }
}
