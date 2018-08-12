/*
  Copyright (c) 2016-2018 Francois Gendron <fg@frgn.ca>
  MIT License

  AudioDeviceCmdlets.cs
  AudioDeviceCmdlets is a suite of PowerShell Cmdlets to control audio devices on Windows
  https://github.com/frgnca/AudioDeviceCmdlets
*/

// To interact with MMDevice
using CoreAudioApi;
// To act as a PowerShell Cmdlet
using System.Management.Automation;

namespace AudioDeviceCmdlets
{
    // Class to interact with a MMDevice as an object with attributes
    public class AudioDevice
    {
        // Order in which this MMDevice appeared from MMDeviceEnumerator
        public int Index;
        // Default (for its Type) is either true or false
        public bool Default;
        // Type is either "Playback" or "Recording"
        public string Type;
        // Name of the MMDevice ex: "Speakers (Realtek High Definition Audio)"
        public string Name;
        // ID of the MMDevice ex: "{0.0.0.00000000}.{c4aadd95-74c7-4b3b-9508-b0ef36ff71ba}"
        public string ID;
        // The MMDevice itself
        public MMDevice Device;

        // To be created, a new AudioDevice needs an Index, and the MMDevice it will communicate with
        public AudioDevice(int Index, MMDevice BaseDevice, bool Default = false)
        {
            // Set this object's Index to the received integer
            this.Index = Index;

            // Set this object's Default to the received boolean
            this.Default = Default;

            // If the received MMDevice is a playback device
            if (BaseDevice.DataFlow == EDataFlow.eRender)
            {
                // Set this object's Type to "Playback"
                this.Type = "Playback";
            }
            // If not, if the received MMDevice is a recording device
            else if (BaseDevice.DataFlow == EDataFlow.eCapture)
            {
                // Set this object's Type to "Recording"
                this.Type = "Recording";
            }

            // Set this object's Name to that of the received MMDevice's FriendlyName
            this.Name = BaseDevice.FriendlyName;

            // Set this object's Device to the received MMDevice
            this.Device = BaseDevice;

            // Set this object's ID to that of the received MMDevice's ID
            this.ID = BaseDevice.ID;
        }
    }

    // Get Cmdlet
    [Cmdlet(VerbsCommon.Get, "AudioDevice")]
    public class GetAudioDevice : Cmdlet
    {
        // Parameter called to list all devices
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ShowDisabled")]
        public SwitchParameter List
        {
            get { return list; }
            set { list = value; }
        }
        private bool list;

        // Parameter receiving the ID of the device to get
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ID")]
        public string ID
        {
            get { return id; }
            set { id = value; }
        }
        private string id;

        // Parameter receiving the Index of the device to get
        [ValidateRange(1, 42)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Index")]
        public int? Index
        {
            get { return index; }
            set { index = value; }
        }
        private int? index;

        // Parameter called to list the default playback device
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Playback")]
        public SwitchParameter Playback
        {
            get { return playback; }
            set { playback = value; }
        }
        private bool playback;

        // Parameter called to list the default playback device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackMute")]
        public SwitchParameter PlaybackMute
        {
            get { return playbackmute; }
            set { playbackmute = value; }
        }
        private bool playbackmute;

        // Parameter called to list the default playback device's volume
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackVolume")]
        public SwitchParameter PlaybackVolume
        {
            get { return playbackvolume; }
            set { playbackvolume = value; }
        }
        private bool playbackvolume;

        // Parameter called to list the default recording device
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Recording")]
        public SwitchParameter Recording
        {
            get { return recording; }
            set { recording = value; }
        }
        private bool recording;

        // Parameter called to list the default recording device' mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingMute")]
        public SwitchParameter RecordingMute
        {
            get { return recordingmute; }
            set { recordingmute = value; }
        }
        private bool recordingmute;

        // Parameter called to list the default recording device' volume
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingVolume")]
        public SwitchParameter RecordingVolume
        {
            get { return recordingvolume; }
            set { recordingvolume = value; }
        }
        private bool recordingvolume;

        // Parameter called to consider disabled devices
        [Parameter(Mandatory = false, ParameterSetName = "ShowDisabled")]
        public SwitchParameter ShowDisabled
        {
            get { return showdisabled; }
            set { showdisabled = value; }
        }
        private bool showdisabled;

        // Cmdlet execution
        protected override void ProcessRecord()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);

            // If the List switch parameter was called
            if (list)
            {
                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is either, the same the default playback device's ID, or the same as the default recording device's ID
                    if (DeviceCollection[i].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).ID || DeviceCollection[i].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).ID)
                    {
                        // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and a default value of true
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], true));
                    }
                    else
                    {
                        // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i]));
                    }
                }
                
                // If the ShowDisabled parameter was called
                if (showdisabled)
                {
                    // The ShowDisabled parameter was called

                    // Get enabled DeviceCollection count
                    int enabledCount = DeviceCollection.Count;

                    // Get MMDeviceCollection of every disabled devices
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_UNPLUGGED);

                    // For every MMDevice in DeviceCollection
                    for (int i = 0; i < DeviceCollection.Count; i++)
                    {
                        // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself
                        WriteObject(new AudioDevice(i + 1 + enabledCount, DeviceCollection[i]));
                    }
                }

                // Stop checking for other parameters
                return;
            }

            // If the ID parameter received a value
            if (!string.IsNullOrEmpty(id))
            {
                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same as the string received by the ID parameter
                    if (string.Compare(DeviceCollection[i].ID, id, System.StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        // If this MMDevice's ID is either, the same the default playback device's ID, or the same as the default recording device's ID
                        if (DeviceCollection[i].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).ID || DeviceCollection[i].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).ID)
                        {
                            // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and a default value of true
                            WriteObject(new AudioDevice(i + 1, DeviceCollection[i], true));
                        }
                        else
                        {
                            // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself
                            WriteObject(new AudioDevice(i + 1, DeviceCollection[i]));
                        }

                        // Stop checking for other parameters
                        return;
                    }
                }

                // Throw an exception about the received ID not being found
                throw new System.ArgumentException("No AudioDevice with that ID");
            }

            // If the Index parameter received a value
            if (index != null)
            {
                // If the Index is valid
                if (index.Value >= 1 && index.Value <= DeviceCollection.Count)
                {
                    // If the ID of the MMDevice associated with the Index is the same as, either the ID of the default playback device, or the ID of the default recording device
                    if (DeviceCollection[index.Value - 1].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).ID || DeviceCollection[index.Value - 1].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).ID)
                    {
                        // Output the result of the creation of a new AudioDevice while assining it the an index, and the MMDevice itself, and a default value of true
                        WriteObject(new AudioDevice(index.Value, DeviceCollection[index.Value - 1], true));

                        // Stop checking for other parameters
                        return;
                    }
                    else
                    {
                        // Output the result of the creation of a new AudioDevice while assining it the an index, and the MMDevice itself, and a default value of false
                        WriteObject(new AudioDevice(index.Value, DeviceCollection[index.Value - 1], false));

                        // Stop checking for other parameters
                        return;
                    }
                }
                else
                {
                    // Throw an exception about the received Index not being found
                    throw new System.ArgumentException("No AudioDevice with that Index");
                }
            }

            // If the Playback switch parameter was called
            if (playback)
            {
                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same the default playback device's ID
                    if (DeviceCollection[i].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).ID)
                    {
                        // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and a default value of true
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], true));
                    }
                }
                
                // Stop checking for other parameters
                return;
            }

            // If the PlaybackMute switch parameter was called
            if (playbackmute)
            {
                // Output the mute state of the default playback device
                WriteObject(DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute);

                // Stop checking for other parameters
                return;
            }

            // If the PlaybackVolume switch parameter was called
            if(playbackvolume)
            {
                // Output the current volume level of the default playback device
                WriteObject(string.Format("{0}%", DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.MasterVolumeLevelScalar * 100));

                // Stop checking for other parameters
                return;
            }

            // If the Recording switch parameter was called
            if (recording)
            {
                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same the default recording device's ID
                    if (DeviceCollection[i].ID == DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).ID)
                    {
                        // Output the result of the creation of a new AudioDevice while assining it an index, and the MMDevice itself, and a default value of true
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], true));
                    }
                }
                
                // Stop checking for other parameters
                return;
            }

            // If the RecordingMute switch parameter was called
            if (recordingmute)
            {
                // Output the mute state of the default recording device
                WriteObject(DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute);

                // Stop checking for other parameters
                return;
            }

            // If the RecordingVolume switch parameter was called
            if (recordingvolume)
            {
                // Output the current volume level of the default recording device
                WriteObject(string.Format("{0}%", DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.MasterVolumeLevelScalar * 100));

                // Stop checking for other parameters
                return;
            }
        }
    }

    // Set Cmdlet
    [Cmdlet(VerbsCommon.Set, "AudioDevice")]
    public class SetAudioDevice : Cmdlet
    {
        // Parameter receiving the AudioDevice to set as default
        [Parameter(Mandatory = true, ParameterSetName = "InputObject", ValueFromPipeline = true)]
        public AudioDevice InputObject
        {
            get { return inputObject; }
            set { inputObject = value; }

        }
        private AudioDevice inputObject;

        // Parameter receiving the ID of the device to set as default
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ID")]
        public string ID
        {
            get { return id; }
            set { id = value; }
        }
        private string id;

        // Parameter receiving the Index of the device to set as default
        [ValidateRange(1, 42)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Index")]
        public int? Index
        {
            get { return index; }
            set { index = value; }
        }
        private int? index;

        // Parameter called to set the default playback device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackMute")]
        public bool? PlaybackMute
        {
            get { return playbackmute; }
            set { playbackmute = value; }
        }
        private bool? playbackmute;

        // Parameter called to toggle the default playback device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackMuteToggle")]
        public SwitchParameter PlaybackMuteToggle
        {
            get { return playbackmutetoggle; }
            set { playbackmutetoggle = value; }
        }
        private SwitchParameter playbackmutetoggle;

        // Parameter receiving the volume level to set to the defaut playback device
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackVolume")]
        public float? PlaybackVolume
        {
            get { return playbackvolume; }
            set { playbackvolume = value; }
        }
        private float? playbackvolume;

        // Parameter called to set the default recording device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingMute")]
        public bool? RecordingMute
        {
            get { return recordingmute; }
            set { recordingmute = value; }
        }
        private bool? recordingmute;

        // Parameter called to toggle the default recording device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingMuteToggle")]
        public SwitchParameter RecordingMuteToggle
        {
            get { return recordingmutetoggle; }
            set { recordingmutetoggle = value; }
        }
        private SwitchParameter recordingmutetoggle;

        // Parameter receiving the volume level to set to the defaut recording device
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingVolume")]
        public float? RecordingVolume
        {
            get { return recordingvolume; }
            set { recordingvolume = value; }
        }
        private float? recordingvolume;

        // Cmdlet execution
        protected override void ProcessRecord()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            // Create a MMDeviceCollection of every devices that are enabled
            MMDeviceCollection DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);

            // If the InputObject parameter received a value
            if (inputObject != null)
            {
                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same as the ID of the MMDevice received by the InputObject parameter
                    if (DeviceCollection[i].ID == inputObject.ID)
                    {
                        // Create a new audio PolicyConfigClient
                        PolicyConfigClient client = new PolicyConfigClient();
                        // Using PolicyConfigClient, set the given device as the default playback communication device
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);
                        // Using PolicyConfigClient, set the given device as the default playback device
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);

                        // Output the result of the creation of a new AudioDevice while assining it the an index, and the MMDevice itself, and a default value of true
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], true));

                        // Stop checking for other parameters
                        return;
                    }
                }

                // Throw an exception about the received device not being found
                throw new System.ArgumentException("No such enabled AudioDevice found");
            }

            // If the ID parameter received a value
            if (!string.IsNullOrEmpty(id))
            {
                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same as the string received by the ID parameter
                    if (string.Compare(DeviceCollection[i].ID, id, System.StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        // Create a new audio PolicyConfigClient
                        PolicyConfigClient client = new PolicyConfigClient();
                        // Using PolicyConfigClient, set the given device as the default communication device (for its type)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);
                        // Using PolicyConfigClient, set the given device as the default device (for its type)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);

                        // Output the result of the creation of a new AudioDevice while assining it the index, and the MMDevice itself, and a default value of true
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], true));

                        // Stop checking for other parameters
                        return;
                    }
                }

                // Throw an exception about the received ID not being found
                throw new System.ArgumentException("No enabled AudioDevice found with that ID");
            }

            // If the Index parameter received a value
            if (index != null)
            {
                // If the Index is valid
                if (index.Value >= 1 && index.Value <= DeviceCollection.Count)
                {
                    // Create a new audio PolicyConfigClient
                    PolicyConfigClient client = new PolicyConfigClient();
                    // Using PolicyConfigClient, set the given device as the default communication device (for its type)
                    client.SetDefaultEndpoint(DeviceCollection[index.Value - 1].ID, ERole.eCommunications);
                    // Using PolicyConfigClient, set the given device as the default device (for its type)
                    client.SetDefaultEndpoint(DeviceCollection[index.Value - 1].ID, ERole.eMultimedia);

                    // Output the result of the creation of a new AudioDevice while assining it the index, and the MMDevice itself, and a default value of true
                    WriteObject(new AudioDevice(index.Value, DeviceCollection[index.Value - 1], true));

                    // Stop checking for other parameters
                    return;
                }
                else
                {
                    // Throw an exception about the received Index not being found
                    throw new System.ArgumentException("No enabled AudioDevice found with that Index");
                }
            }

            // If the PlaybackMute parameter received a value
            if (playbackmute != null)
            {
                // Set the mute state of the default playback device to that of the boolean value received by the Cmdlet
                DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = (bool)playbackmute;
            }

            // If the PlaybackMuteToggle paramter was called
            if (playbackmutetoggle)
            {
                // Toggle the mute state of the default playback device
                DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = !DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute;
            }

            // If the PlaybackVolume parameter received a value
            if(playbackvolume != null)
            {
                // Set the volume level of the default playback device to that of the float value received by the PlaybackVolume parameter
                DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.MasterVolumeLevelScalar = (float)playbackvolume / 100.0f;
            }

            // If the RecordingMute parameter received a value
            if (recordingmute != null)
            {
                // Set the mute state of the default recording device to that of the boolean value received by the Cmdlet
                DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute = (bool)recordingmute;
            }

            // If the RecordingMuteToggle paramter was called
            if (recordingmutetoggle)
            {
                // Toggle the mute state of the default recording device
                DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute = !DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute;
            }

            // If the RecordingVolume parameter received a value
            if (recordingvolume != null)
            {
                // Set the volume level of the default recording device to that of the float value received by the RecordingVolume parameter
                DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.MasterVolumeLevelScalar = (float)recordingvolume / 100.0f;
            }
        }
    }

    // Write Cmdlet
    [Cmdlet(VerbsCommunications.Write, "AudioDevice")]
    public class WriteAudioDevice : Cmdlet
    {
        // Parameter called to output audiometer result of the default playback device as a progress bar
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackMeter")]
        public SwitchParameter PlaybackMeter
        {
            get { return playbackmeter; }
            set { playbackmeter = value; }
        }
        private bool playbackmeter;

        // Parameter called to output audiometer result of the default playback device as a stream of values
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackStream")]
        public SwitchParameter PlaybackStream
        {
            get { return playbackstream; }
            set { playbackstream = value; }
        }
        private bool playbackstream;

        // Parameter called to output audiometer result of the default recording device as a progress bar
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingMeter")]
        public SwitchParameter RecordingMeter
        {
            get { return recordingmeter; }
            set { recordingmeter = value; }
        }
        private bool recordingmeter;

        // Parameter called to output audiometer result of the default recording device as a stream of values
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingStream")]
        public SwitchParameter RecordingStream
        {
            get { return recordingstream; }
            set { recordingstream = value; }
        }
        private bool recordingstream;

        // Cmdlet execution
        protected override void ProcessRecord()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();

            // If the PlaybackMeter parameter was called
            if (playbackmeter)
            {
                // Create a new progress bar to output current audiometer result of the default playback device
                ProgressRecord pr = new ProgressRecord(0, DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).FriendlyName, "Peak Value");
                // Set the progress bar to zero
                pr.PercentComplete = 0;

                // Loop until interruption ex: CTRL+C
                do
                {
                    // Set progress bar to current audiometer result
                    pr.PercentComplete = System.Convert.ToInt32(DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue * 100);

                    // Write current audiometer result as a progress bar
                    WriteProgress(pr);

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the PlaybackStream parameter was called
            if (playbackstream)
            {
                // Loop until interruption ex: CTRL+C
                do
                {
                    // Write current audiometer result as a value
                    WriteObject(System.Convert.ToInt32(DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue * 100));

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the RecordingMeter parameter was called
            if (recordingmeter)
            {
                // Create a new progress bar to output current audiometer result of the default recording device
                ProgressRecord pr = new ProgressRecord(0, DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).FriendlyName, "Peak Value");
                // Set the progress bar to zero
                pr.PercentComplete = 0;

                // Loop until interruption ex: CTRL+C
                do
                {
                    // Set progress bar to current audiometer result
                    pr.PercentComplete = System.Convert.ToInt32(DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue * 100);

                    // Write current audiometer result as a progress bar
                    WriteProgress(pr);

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the RecordingStream parameter was called
            if (recordingstream)
            {
                // Loop until interruption ex: CTRL+C
                do
                {
                    // Write current audiometer result as a value
                    WriteObject(System.Convert.ToInt32(DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue * 100));

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }
        }
    }
}
