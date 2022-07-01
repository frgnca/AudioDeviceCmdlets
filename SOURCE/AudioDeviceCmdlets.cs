/*
  Copyright (c) 2016-2022 Francois Gendron <fg@frgn.ca>
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
        // DefaultCommunication (for its Type) is either true or false
        public bool DefaultCommunication;
        // Type is either "Playback" or "Recording"
        public string Type;
        // Name of the MMDevice ex: "Speakers (Realtek High Definition Audio)"
        public string Name;
        // ID of the MMDevice ex: "{0.0.0.00000000}.{c4aadd95-74c7-4b3b-9508-b0ef36ff71ba}"
        public string ID;
        // The MMDevice itself
        public MMDevice Device;

        // To be created, a new AudioDevice needs an Index, and the MMDevice it will communicate with
        public AudioDevice(int Index, MMDevice BaseDevice, bool Default = false, bool DefaultCommunication = false)
        {
            // Set this object's Index to the received integer
            this.Index = Index;

            // Set this object's Default to the received boolean
            this.Default = Default;

            // Set this object's DefaultCommunication to the received boolean
            this.DefaultCommunication = DefaultCommunication;

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

    // Class to get information on a MMDevice towards the creation of a corresponding AudioDevice
    public class AudioDeviceCreationToolkit
    {
        // The MMDeviceEnumerator
        public MMDeviceEnumerator DevEnum;

        // To be created, a new AudioDeviceCreationToolkit needs a MMDeviceEnumerator it will use to compare the ID its methods receive
        public AudioDeviceCreationToolkit(MMDeviceEnumerator DevEnum)
        {
            // Set this object's DeviceEnumerator to the received MMDeviceEnumerator
            this.DevEnum = DevEnum;
        }

        // Method to find out, in a collection of all enabled MMDevice, the Index of a MMDevice, given its ID
        public int FindIndex(string ID)
        {
            MMDeviceCollection DeviceCollection = null;
            try
            {
                // Enumerate all enabled devices in a collection
                DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
            }
            catch
            {
                // Error
                throw new System.Exception("Error in method AudioDeviceCreationToolkit.FindIndex(string ID) - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
            }

            // For each device in the collection
            for (int i = 0; i < DeviceCollection.Count; i++)
            {
                // If the received ID is the same as this device's ID
                if(DeviceCollection[i].ID == ID)
                {
                    // Return this device's Index
                    return (i + 1);
                }
            }

            // Error
            throw new System.Exception("Error in method AudioDeviceCreationToolkit.FindIndex(string ID) - No MMDevice with the given ID was found in the collection of all enabled MMDevice");
        }

        // Method to find out if a MMDevice is the default MMDevice of its type, given its ID
        public bool IsDefault(string ID)
        {
            string PlaybackID = "";
            try
            {
                // Get the ID of the default playback device
                PlaybackID = (DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia)).ID;
            }
            catch { }

            // If the received ID is the same as the default playback device's ID
            if(ID == PlaybackID)
            {
                return (true);
            }

            string RecordingID = "";
            try
            {
                // Get the ID of the default recording device
                RecordingID = (DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia)).ID;
            }
            catch { }

            // If the received ID is the same as the default recording device's ID
            if (ID == RecordingID)
            {
                return (true);
            }

            return (false);
        }

        // Method to find out if a MMDevice is the default communication MMDevice of its type, given its ID
        public bool IsDefaultCommunication(string ID)
        {
            string PlaybackCommunicationID = "";
            try
            {
                // Get the ID of the default communication playback device
                PlaybackCommunicationID = (DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications)).ID;
            }
            catch { }

            // If the received ID is the same as the default communication playback device's ID
            if (ID == PlaybackCommunicationID)
            {
                return (true);
            }

            string RecordingCommunicationID = "";
            try
            {
                // Get the ID of the default communication recording device
                RecordingCommunicationID = (DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications)).ID;
            }
            catch { }

            // If the received ID is the same as the default communication recording device's ID
            if (ID == RecordingCommunicationID)
            {
                return (true);
            }

            return (false);
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

        // Parameter called to list the default communication playback device
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunication")]
        public SwitchParameter PlaybackCommunication
        {
            get { return playbackcommunication; }
            set { playbackcommunication = value; }
        }
        private bool playbackcommunication;

        // Parameter called to list the default communication playback device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationMute")]
        public SwitchParameter PlaybackCommunicationMute
        {
            get { return playbackcommunicationmute; }
            set { playbackcommunicationmute = value; }
        }
        private bool playbackcommunicationmute;

        // Parameter called to list the default communication playback device's volume
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationVolume")]
        public SwitchParameter PlaybackCommunicationVolume
        {
            get { return playbackcommunicationvolume; }
            set { playbackcommunicationvolume = value; }
        }
        private bool playbackcommunicationvolume;

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

        // Parameter called to list the default communication recording device
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunication")]
        public SwitchParameter RecordingCommunication
        {
            get { return recordingcommunication; }
            set { recordingcommunication = value; }
        }
        private bool recordingcommunication;

        // Parameter called to list the default communication recording device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationMute")]
        public SwitchParameter RecordingCommunicationMute
        {
            get { return recordingcommunicationmute; }
            set { recordingcommunicationmute = value; }
        }
        private bool recordingcommunicationmute;

        // Parameter called to list the default communication recording device's volume
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationVolume")]
        public SwitchParameter RecordingCommunicationVolume
        {
            get { return recordingcommunicationvolume; }
            set { recordingcommunicationvolume = value; }
        }
        private bool recordingcommunicationvolume;

        // Parameter called to list the default recording device
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Recording")]
        public SwitchParameter Recording
        {
            get { return recording; }
            set { recording = value; }
        }
        private bool recording;

        // Parameter called to list the default recording device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingMute")]
        public SwitchParameter RecordingMute
        {
            get { return recordingmute; }
            set { recordingmute = value; }
        }
        private bool recordingmute;

        // Parameter called to list the default recording device's volume
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingVolume")]
        public SwitchParameter RecordingVolume
        {
            get { return recordingvolume; }
            set { recordingvolume = value; }
        }
        private bool recordingvolume;

        // Parameter called to display version and credit info
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Version")]
        public SwitchParameter Version
        {
            get { return version; }
            set { version = value; }
        }
        private bool version;

        // Cmdlet execution
        protected override void ProcessRecord()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();

            // If the List switch parameter was called
            if (list)
            {
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDeviceCollection DeviceCollection = null;
                try
                {
                    // Enumerate all enabled devices in a collection
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
                }
                catch
                {
                    // Error
                    throw new System.Exception("Error in parameter List - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
                }

                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                    WriteObject(new AudioDevice(i + 1, DeviceCollection[i], Toolkit.IsDefault(DeviceCollection[i].ID), Toolkit.IsDefaultCommunication(DeviceCollection[i].ID)));
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
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDeviceCollection DeviceCollection = null;
                try
                {
                    // Enumerate all enabled devices in a collection
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
                }
                catch
                {
                    // Error
                    throw new System.Exception("Error in parameter ID - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
                }

                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same as the string received by the ID parameter
                    if (string.Compare(DeviceCollection[i].ID, id, System.StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], Toolkit.IsDefault(DeviceCollection[i].ID), Toolkit.IsDefaultCommunication(DeviceCollection[i].ID)));

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
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDeviceCollection DeviceCollection = null;
                try
                {
                    // Enumerate all enabled devices in a collection
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
                }
                catch
                {
                    // Error
                    throw new System.Exception("Error in parameter Index - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
                }

                // If the Index is valid
                if (index.Value >= 1 && index.Value <= DeviceCollection.Count)
                {
                    // Use valid Index as iterative
                    int i = index.Value - 1;

                    // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                    WriteObject(new AudioDevice(i + 1, DeviceCollection[i], Toolkit.IsDefault(DeviceCollection[i].ID), Toolkit.IsDefaultCommunication(DeviceCollection[i].ID)));

                    // Stop checking for other parameters
                    return;
                }
                else
                {
                    // Throw an exception about the received Index not being found
                    throw new System.ArgumentException("No AudioDevice with that Index");
                }
            }

            // If the PlaybackCommunication switch parameter was called
            if (playbackcommunication)
            {
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDevice Device = null;
                try
                {
                    // Get the default communication playback device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }

                // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                WriteObject(new AudioDevice(Toolkit.FindIndex(Device.ID), Device, Toolkit.IsDefault(Device.ID), Toolkit.IsDefaultCommunication(Device.ID)));

                // Stop checking for other parameters
                return;
            }

            // If the PlaybackCommunicationMute switch parameter was called
            if (playbackcommunicationmute)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default communication playback device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }

                // Output the mute state of the default communication playback device
                WriteObject(Device.AudioEndpointVolume.Mute);

                // Stop checking for other parameters
                return;
            }

            // If the PlaybackCommunicationVolume switch parameter was called
            if (playbackcommunicationvolume)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default communication playback device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }

                // Output the current volume level of the default communication playback device
                WriteObject(string.Format("{0}%", Device.AudioEndpointVolume.MasterVolumeLevelScalar * 100));

                // Stop checking for other parameters
                return;
            }

            // If the Playback switch parameter was called
            if (playback)
            {
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDevice Device = null;
                try
                {
                    // Get the default playback device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }

                // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                WriteObject(new AudioDevice(Toolkit.FindIndex(Device.ID), Device, Toolkit.IsDefault(Device.ID), Toolkit.IsDefaultCommunication(Device.ID)));

                // Stop checking for other parameters
                return;
            }

            // If the PlaybackMute switch parameter was called
            if (playbackmute)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default playback device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }

                // Output the mute state of the default playback device
                WriteObject(Device.AudioEndpointVolume.Mute);

                // Stop checking for other parameters
                return;
            }

            // If the PlaybackVolume switch parameter was called
            if(playbackvolume)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default playback device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }

                // Output the current volume level of the default playback device
                WriteObject(string.Format("{0}%", Device.AudioEndpointVolume.MasterVolumeLevelScalar * 100));

                // Stop checking for other parameters
                return;
            }

            // If the RecordingCommunication switch parameter was called
            if (recordingcommunication)
            {
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDevice Device = null;
                try
                {
                    // Get the default communication recording device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }

                // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                WriteObject(new AudioDevice(Toolkit.FindIndex(Device.ID), Device, Toolkit.IsDefault(Device.ID), Toolkit.IsDefaultCommunication(Device.ID)));

                // Stop checking for other parameters
                return;
            }

            // If the RecordingCommunicationMute switch parameter was called
            if (recordingcommunicationmute)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default communication recording device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }

                // Output the mute state of the default communication recording device
                WriteObject(Device.AudioEndpointVolume.Mute);

                // Stop checking for other parameters
                return;
            }

            // If the RecordingCommunicationVolume switch parameter was called
            if (recordingcommunicationvolume)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default communication recording device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }

                // Output the current volume level of the default communication recording device
                WriteObject(string.Format("{0}%", Device.AudioEndpointVolume.MasterVolumeLevelScalar * 100));

                // Stop checking for other parameters
                return;
            }

            // If the Recording switch parameter was called
            if (recording)
            {
                // Create a AudioDeviceCreationToolkit
                AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                MMDevice Device = null;
                try
                {
                    // Get the default recording device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }

                // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                WriteObject(new AudioDevice(Toolkit.FindIndex(Device.ID), Device, Toolkit.IsDefault(Device.ID), Toolkit.IsDefaultCommunication(Device.ID)));

                // Stop checking for other parameters
                return;
            }

            // If the RecordingMute switch parameter was called
            if (recordingmute)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default recording device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }

                // Output the mute state of the default recording device
                WriteObject(Device.AudioEndpointVolume.Mute);

                // Stop checking for other parameters
                return;
            }

            // If the RecordingVolume switch parameter was called
            if (recordingvolume)
            {
                MMDevice Device = null;
                try
                {
                    // Get the default recording device
                    Device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }

                // Output the current volume level of the default recording device
                WriteObject(string.Format("{0}%", Device.AudioEndpointVolume.MasterVolumeLevelScalar * 100));

                // Stop checking for other parameters
                return;
            }

            // If the Version parameter was called
            if (version)
            {
                // Version text
                string text = @"
  AudioDeviceCmdlets v3.1.0.2

  Copyright (c) 2016-2022 Francois Gendron <fg@frgn.ca>
  MIT License

  Thank you for considering a donation
  Bitcoin     (BTC) 3AffczXX4Jb2iN8QWQhHQAsj9AqGFXgYUF
  BitcoinCash (BCH) qraf6a3fklta7xkvwkh49zqn6mgnm2eyz589rkfvl3
  Ethereum    (ETH) 0xE4EA2A2356C04c8054Db452dCBd6f958F74722dE
";

                // Write version text
                WriteObject(text);

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

        // Parameter called to set the default communication playback device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationMute")]
        public bool? PlaybackCommunicationMute
        {
            get { return playbackcommunicationmute; }
            set { playbackcommunicationmute = value; }
        }
        private bool? playbackcommunicationmute;

        // Parameter called to toggle the default communication playback device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationMuteToggle")]
        public SwitchParameter PlaybackCommunicationMuteToggle
        {
            get { return playbackcommunicationmutetoggle; }
            set { playbackcommunicationmutetoggle = value; }
        }
        private SwitchParameter playbackcommunicationmutetoggle;

        // Parameter receiving the volume level to set to the default communication playback device
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationVolume")]
        public float? PlaybackCommunicationVolume
        {
            get { return playbackcommunicationvolume; }
            set { playbackcommunicationvolume = value; }
        }
        private float? playbackcommunicationvolume;

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

        // Parameter receiving the volume level to set to the default playback device
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackVolume")]
        public float? PlaybackVolume
        {
            get { return playbackvolume; }
            set { playbackvolume = value; }
        }
        private float? playbackvolume;

        // Parameter called to set the default communication recording device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationMute")]
        public bool? RecordingCommunicationMute
        {
            get { return recordingcommunicationmute; }
            set { recordingcommunicationmute = value; }
        }
        private bool? recordingcommunicationmute;

        // Parameter called to toggle the default communication recording device's mute state
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationMuteToggle")]
        public SwitchParameter RecordingCommunicationMuteToggle
        {
            get { return recordingcommunicationmutetoggle; }
            set { recordingcommunicationmutetoggle = value; }
        }
        private SwitchParameter recordingcommunicationmutetoggle;

        // Parameter receiving the volume level to set to the default communication recording device
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationVolume")]
        public float? RecordingCommunicationVolume
        {
            get { return recordingcommunicationvolume; }
            set { recordingcommunicationvolume = value; }
        }
        private float? recordingcommunicationvolume;

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

        // Parameter receiving the volume level to set to the default recording device
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingVolume")]
        public float? RecordingVolume
        {
            get { return recordingvolume; }
            set { recordingvolume = value; }
        }
        private float? recordingvolume;

        // Parameter called to only set device as default and not default communication
        [Parameter(Mandatory = false, ParameterSetName = "InputObject")]
        [Parameter(Mandatory = false, ParameterSetName = "ID")]
        [Parameter(Mandatory = false, ParameterSetName = "Index")]
        public SwitchParameter DefaultOnly
        {
            get { return defaultOnly; }
            set { defaultOnly = value; }
        }
        private SwitchParameter defaultOnly;

        // Parameter called to only set device as default communication and not default
        [Parameter(Mandatory = false, ParameterSetName = "InputObject")]
        [Parameter(Mandatory = false, ParameterSetName = "ID")]
        [Parameter(Mandatory = false, ParameterSetName = "Index")]
        public SwitchParameter CommunicationOnly
        {
            get { return communicationOnly; }
            set { communicationOnly = value; }
        }
        private SwitchParameter communicationOnly;

        // Parameter called to display version and credit info
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Version")]
        public SwitchParameter Version
        {
            get { return version; }
            set { version = value; }
        }
        private bool version;

        // Cmdlet execution
        protected override void ProcessRecord()
        {
            if (defaultOnly.ToBool() && communicationOnly.ToBool())
                throw new System.ArgumentException("Impossible to do both DefaultOnly and CommunicatioOnly at the same time.");

            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();

            // If the InputObject parameter received a value
            if (inputObject != null)
            {
                MMDeviceCollection DeviceCollection = null;
                try
                {
                    // Enumerate all enabled devices in a collection
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
                }
                catch
                {
                    // Error
                    throw new System.Exception("Error in parameter InputObject - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
                }

                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same as the ID of the MMDevice received by the InputObject parameter
                    if (DeviceCollection[i].ID == inputObject.ID)
                    {
                        // To use during creation of corresponding AudioDevice, assuming it is impossible to do both DefaultOnly and CommunicatioOnly at the same time
                        bool DefaultState;
                        bool CommunicationState;

                        // Create a new audio PolicyConfigClient
                        PolicyConfigClient client = new PolicyConfigClient();

                        // Create a AudioDeviceCreationToolkit
                        AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                        // Unless the DefaultOnly parameter was called
                        if (!defaultOnly.ToBool())
                        {
                            // The DefaultOnly parameter was not called

                            // Using PolicyConfigClient, set the given device as the default communication device (for its type)
                            client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);

                            // Set default communication state to use
                            CommunicationState = true;
                        }
                        else
                        {
                            // The DefaultOnly parameter was called

                            // Set default communication state to use
                            CommunicationState = Toolkit.IsDefaultCommunication(DeviceCollection[i].ID);
                        }

                        // Unless the CommunicationOnly parameter was called
                        if (!communicationOnly.ToBool())
                        {
                            // The CommunicationOnly parameter was not called

                            // Using PolicyConfigClient, set the given device as the default device (for its type)
                            client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);

                            // Set default state to use
                            DefaultState = true;
                        }
                        else
                        {
                            // The CommunicationOnly parameter was called

                            // Set default state to use
                            DefaultState = Toolkit.IsDefault(DeviceCollection[i].ID);
                        }

                        // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], DefaultState, CommunicationState));

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
                MMDeviceCollection DeviceCollection = null;
                try
                {
                    // Enumerate all enabled devices in a collection
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
                }
                catch
                {
                    // Error
                    throw new System.Exception("Error in parameter ID - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
                }

                // For every MMDevice in DeviceCollection
                for (int i = 0; i < DeviceCollection.Count; i++)
                {
                    // If this MMDevice's ID is the same as the string received by the ID parameter
                    if (string.Compare(DeviceCollection[i].ID, id, System.StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        // To use during creation of corresponding AudioDevice, assuming it is impossible to do both DefaultOnly and CommunicatioOnly at the same time
                        bool DefaultState;
                        bool CommunicationState;

                        // Create a new audio PolicyConfigClient
                        PolicyConfigClient client = new PolicyConfigClient();

                        // Create a AudioDeviceCreationToolkit
                        AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                        // Unless the DefaultOnly parameter was called
                        if (!defaultOnly.ToBool())
                        {
                            // The DefaultOnly parameter was not called

                            // Using PolicyConfigClient, set the given device as the default communication device (for its type)
                            client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);

                            // Set default communication state to use
                            CommunicationState = true;
                        }
                        else
                        {
                            // The DefaultOnly parameter was called

                            // Set default communication state to use
                            CommunicationState = Toolkit.IsDefaultCommunication(DeviceCollection[i].ID);
                        }

                        // Unless the CommunicationOnly parameter was called
                        if (!communicationOnly.ToBool())
                        {
                            // The CommunicationOnly parameter was not called

                            // Using PolicyConfigClient, set the given device as the default device (for its type)
                            client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);

                            // Set default state to use
                            DefaultState = true;
                        }
                        else
                        {
                            // The CommunicationOnly parameter was called

                            // Set default state to use
                            DefaultState = Toolkit.IsDefault(DeviceCollection[i].ID);
                        }

                        // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                        WriteObject(new AudioDevice(i + 1, DeviceCollection[i], DefaultState, CommunicationState));

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
                MMDeviceCollection DeviceCollection = null;
                try
                {
                    // Enumerate all enabled devices in a collection
                    DeviceCollection = DevEnum.EnumerateAudioEndPoints(EDataFlow.eAll, EDeviceState.DEVICE_STATE_ACTIVE);
                }
                catch
                {
                    // Error
                    throw new System.Exception("Error in parameter Index - Failed to create the collection of all enabled MMDevice using MMDeviceEnumerator");
                }

                // If the Index is valid
                if (index.Value >= 1 && index.Value <= DeviceCollection.Count)
                {
                    // Use valid Index as iterative
                    int i = index.Value - 1;

                    // To use during creation of corresponding AudioDevice, assuming it is impossible to do both DefaultOnly and CommunicatioOnly at the same time
                    bool DefaultState;
                    bool CommunicationState;

                    // Create a new audio PolicyConfigClient
                    PolicyConfigClient client = new PolicyConfigClient();

                    // Create a AudioDeviceCreationToolkit
                    AudioDeviceCreationToolkit Toolkit = new AudioDeviceCreationToolkit(DevEnum);

                    // Unless the DefaultOnly parameter was called
                    if (!defaultOnly.ToBool())
                    {
                        // The DefaultOnly parameter was not called

                        // Using PolicyConfigClient, set the given device as the default communication device (for its type)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eCommunications);

                        // Set default communication state to use
                        CommunicationState = true;
                    }
                    else
                    {
                        // The DefaultOnly parameter was called

                        // Set default communication state to use
                        CommunicationState = Toolkit.IsDefaultCommunication(DeviceCollection[i].ID);
                    }

                    // Unless the CommunicationOnly parameter was called
                    if (!communicationOnly.ToBool())
                    {
                        // The CommunicationOnly parameter was not called

                        // Using PolicyConfigClient, set the given device as the default device (for its type)
                        client.SetDefaultEndpoint(DeviceCollection[i].ID, ERole.eMultimedia);

                        // Set default state to use
                        DefaultState = true;
                    }
                    else
                    {
                        // The CommunicationOnly parameter was called

                        // Set default state to use
                        DefaultState = Toolkit.IsDefault(DeviceCollection[i].ID);
                    }

                    // Output the result of the creation of a new AudioDevice, while assining it its index, the MMDevice itself, its default state, and its default communication state
                    WriteObject(new AudioDevice(i + 1, DeviceCollection[i], DefaultState, CommunicationState));

                    // Stop checking for other parameters
                    return;
                }
                else
                {
                    // Throw an exception about the received Index not being found
                    throw new System.ArgumentException("No enabled AudioDevice found with that Index");
                }
            }

            // If the PlaybackCommunicationMute parameter received a value
            if (playbackcommunicationmute != null)
            {
                try
                {
                    // Set the mute state of the default communication playback device to that of the boolean value received by the Cmdlet
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).AudioEndpointVolume.Mute = (bool)playbackcommunicationmute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }                
            }

            // If the PlaybackCommunicationMuteToggle parameter was called
            if (playbackcommunicationmutetoggle)
            {
                try
                {
                    // Toggle the mute state of the default communication playback device
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).AudioEndpointVolume.Mute = !DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).AudioEndpointVolume.Mute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }
            }

            // If the PlaybackCommunicationVolume parameter received a value
            if(playbackcommunicationvolume != null)
            {
                try
                {
                    // Set the volume level of the default communication playback device to that of the float value received by the PlaybackCommunicationVolume parameter
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).AudioEndpointVolume.MasterVolumeLevelScalar = (float)playbackcommunicationvolume / 100.0f;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }
            }

            // If the PlaybackMute parameter received a value
            if (playbackmute != null)
            {
                try
                {
                    // Set the mute state of the default playback device to that of the boolean value received by the Cmdlet
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = (bool)playbackmute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }
            }

            // If the PlaybackMuteToggle parameter was called
            if (playbackmutetoggle)
            {
                try
                {
                    // Toggle the mute state of the default playback device
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute = !DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.Mute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }
            }

            // If the PlaybackVolume parameter received a value
            if(playbackvolume != null)
            {
                try
                {
                    // Set the volume level of the default playback device to that of the float value received by the PlaybackVolume parameter
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioEndpointVolume.MasterVolumeLevelScalar = (float)playbackvolume / 100.0f;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }
            }

            // If the RecordingCommunicationMute parameter received a value
            if (recordingcommunicationmute != null)
            {
                try
                {
                    // Set the mute state of the default communication recording device to that of the boolean value received by the Cmdlet
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).AudioEndpointVolume.Mute = (bool)recordingcommunicationmute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }
            }

            // If the RecordingCommunicationMuteToggle parameter was called
            if (recordingcommunicationmutetoggle)
            {
                try
                {
                    // Toggle the mute state of the default communication recording device
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).AudioEndpointVolume.Mute = !DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).AudioEndpointVolume.Mute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }
            }

            // If the RecordingCommunicationVolume parameter received a value
            if (recordingcommunicationvolume != null)
            {
                try
                {
                    // Set the volume level of the default communication recording device to that of the float value received by the RecordingCommunicationVolume parameter
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).AudioEndpointVolume.MasterVolumeLevelScalar = (float)recordingcommunicationvolume / 100.0f;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }
            }

            // If the RecordingMute parameter received a value
            if (recordingmute != null)
            {
                try
                {
                    // Set the mute state of the default recording device to that of the boolean value received by the Cmdlet
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute = (bool)recordingmute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }
            }

            // If the RecordingMuteToggle parameter was called
            if (recordingmutetoggle)
            {
                try
                {
                    // Toggle the mute state of the default recording device
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute = !DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.Mute;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }
            }

            // If the RecordingVolume parameter received a value
            if (recordingvolume != null)
            {
                try
                {
                    // Set the volume level of the default recording device to that of the float value received by the RecordingVolume parameter
                    DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioEndpointVolume.MasterVolumeLevelScalar = (float)recordingvolume / 100.0f;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }
            }

            // If the Version parameter was called
            if (version)
            {
                // Version text
                string text = @"
  AudioDeviceCmdlets v3.1.0.2

  Copyright (c) 2016-2022 Francois Gendron <fg@frgn.ca>
  MIT License

  Thank you for considering a donation
  Bitcoin     (BTC) 3AffczXX4Jb2iN8QWQhHQAsj9AqGFXgYUF
  BitcoinCash (BCH) qraf6a3fklta7xkvwkh49zqn6mgnm2eyz589rkfvl3
  Ethereum    (ETH) 0xE4EA2A2356C04c8054Db452dCBd6f958F74722dE
";

                // Write version text
                WriteObject(text);

                // Stop checking for other parameters
                return;
            }
        }
    }

    // Write Cmdlet
    [Cmdlet(VerbsCommunications.Write, "AudioDevice")]
    public class WriteAudioDevice : Cmdlet
    {
        // Parameter called to output audiometer result of the default communication playback device as a progress bar
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationMeter")]
        public SwitchParameter PlaybackCommunicationMeter
        {
            get { return playbackcommunicationmeter; }
            set { playbackcommunicationmeter = value; }
        }
        private bool playbackcommunicationmeter;

        // Parameter called to output audiometer result of the default communication playback device as a stream of values
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "PlaybackCommunicationStream")]
        public SwitchParameter PlaybackCommunicationStream
        {
            get { return playbackcommunicationstream; }
            set { playbackcommunicationstream = value; }
        }
        private bool playbackcommunicationstream;

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

        // Parameter called to output audiometer result of the default communication recording device as a progress bar
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationMeter")]
        public SwitchParameter RecordingCommunicationMeter
        {
            get { return recordingcommunicationmeter; }
            set { recordingcommunicationmeter = value; }
        }
        private bool recordingcommunicationmeter;

        // Parameter called to output audiometer result of the default communication recording device as a stream of values
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "RecordingCommunicationStream")]
        public SwitchParameter RecordingCommunicationStream
        {
            get { return recordingcommunicationstream; }
            set { recordingcommunicationstream = value; }
        }
        private bool recordingcommunicationstream;

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

        // Parameter called to display version and credit info
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Version")]
        public SwitchParameter Version
        {
            get { return version; }
            set { version = value; }
        }
        private bool version;

        // Cmdlet execution
        protected override void ProcessRecord()
        {
            // Create a new MMDeviceEnumerator
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();

            // If the PlaybackCommunicationMeter parameter was called
            if (playbackcommunicationmeter)
            {
                string FriendlyName = null;
                try
                {
                    // Get the name of the default communication playback device
                    FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).FriendlyName;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                }
                // Create a new progress bar to output current audiometer result of the default communication playback device
                ProgressRecord pr = new ProgressRecord(0, FriendlyName, "Peak Value");

                // Set the progress bar to zero
                pr.PercentComplete = 0;

                // Loop until interruption ex: CTRL+C
                do
                {
                    float MasterPeakValue;
                    try
                    {
                        // Get the name of the default communication playback device
                        FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).FriendlyName;

                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                    }
                    // Set progress bar title
                    pr.Activity = FriendlyName;

                    // Set progress bar to current audiometer result
                    pr.PercentComplete = System.Convert.ToInt32(MasterPeakValue * 100);

                    // Write current audiometer result as a progress bar
                    WriteProgress(pr);

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the PlaybackCommunicationStream parameter was called
            if (playbackcommunicationstream)
            {
                // Loop until interruption ex: CTRL+C
                do
                {
                    float MasterPeakValue;
                    try
                    {
                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eCommunications).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No playback AudioDevice found with the default communication role");
                    }
                    // Write current audiometer result as a value
                    WriteObject(System.Convert.ToInt32(MasterPeakValue * 100));
                    
                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the PlaybackMeter parameter was called
            if (playbackmeter)
            {
                string FriendlyName = null;
                try
                {
                    // Get the name of the default playback device
                    FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).FriendlyName;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No playback AudioDevice found with the default role");
                }
                // Create a new progress bar to output current audiometer result of the default playback device
                ProgressRecord pr = new ProgressRecord(0, FriendlyName, "Peak Value");

                // Set the progress bar to zero
                pr.PercentComplete = 0;

                // Loop until interruption ex: CTRL+C
                do
                {
                    float MasterPeakValue;
                    try
                    {
                        // Get the name of the default playback device
                        FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).FriendlyName;

                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No playback AudioDevice found with the default role");
                    }
                    // Set progress bar title
                    pr.Activity = FriendlyName;

                    // Set progress bar to current audiometer result
                    pr.PercentComplete = System.Convert.ToInt32(MasterPeakValue * 100);

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
                    float MasterPeakValue;
                    try
                    {
                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No playback AudioDevice found with the default role");
                    }
                    // Write current audiometer result as a value
                    WriteObject(System.Convert.ToInt32(MasterPeakValue * 100));

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the RecordingCommunicationMeter parameter was called
            if (recordingcommunicationmeter)
            {
                string FriendlyName = null;
                try
                {
                    // Get the name of the default communication recording device
                    FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).FriendlyName;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                }
                // Create a new progress bar to output current audiometer result of the default communication recording device
                ProgressRecord pr = new ProgressRecord(0, FriendlyName, "Peak Value");

                // Set the progress bar to zero
                pr.PercentComplete = 0;

                // Loop until interruption ex: CTRL+C
                do
                {
                    float MasterPeakValue;
                    try
                    {
                        // Get the name of the default communication recording device
                        FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).FriendlyName;

                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                    }
                    // Set progress bar title
                    pr.Activity = FriendlyName;

                    // Set progress bar to current audiometer result
                    pr.PercentComplete = System.Convert.ToInt32(MasterPeakValue * 100);

                    // Write current audiometer result as a progress bar
                    WriteProgress(pr);

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the RecordingCommunicationStream parameter was called
            if (recordingcommunicationstream)
            {
                // Loop until interruption ex: CTRL+C
                do
                {
                    float MasterPeakValue;
                    try
                    {
                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No recording AudioDevice found with the default communication role");
                    }
                    // Write current audiometer result as a value
                    WriteObject(System.Convert.ToInt32(MasterPeakValue * 100));

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the RecordingMeter parameter was called
            if (recordingmeter)
            {
                string FriendlyName = null;
                try
                {
                    // Get the name of the default recording device
                    FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).FriendlyName;
                }
                catch
                {
                    // Throw an exception about the device not being found
                    throw new System.ArgumentException("No recording AudioDevice found with the default role");
                }
                // Create a new progress bar to output current audiometer result of the default recording device
                ProgressRecord pr = new ProgressRecord(0, FriendlyName, "Peak Value");

                // Set the progress bar to zero
                pr.PercentComplete = 0;

                // Loop until interruption ex: CTRL+C
                do
                {
                    float MasterPeakValue;
                    try
                    {
                        // Get the name of the default recording device
                        FriendlyName = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).FriendlyName;

                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No recording AudioDevice found with the default role");
                    }
                    // Set progress bar title
                    pr.Activity = FriendlyName;

                    // Set progress bar to current audiometer result
                    pr.PercentComplete = System.Convert.ToInt32(MasterPeakValue * 100);

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
                    float MasterPeakValue;
                    try
                    {
                        // Get current audio meter master peak value
                        MasterPeakValue = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia).AudioMeterInformation.MasterPeakValue;
                    }
                    catch
                    {
                        // Throw an exception about the device not being found
                        throw new System.ArgumentException("No recording AudioDevice found with the default role");
                    }
                    // Write current audiometer result as a value
                    WriteObject(System.Convert.ToInt32(MasterPeakValue * 100));

                    // Wait 100 milliseconds
                    System.Threading.Thread.Sleep(100);
                }
                // Loop interrupted ex: CTRL+C
                while (!Stopping);
            }

            // If the Version parameter was called
            if (version)
            {
                // Version text
                string text = @"
  AudioDeviceCmdlets v3.1.0.2

  Copyright (c) 2016-2022 Francois Gendron <fg@frgn.ca>
  MIT License

  Thank you for considering a donation
  Bitcoin     (BTC) 3AffczXX4Jb2iN8QWQhHQAsj9AqGFXgYUF
  BitcoinCash (BCH) qraf6a3fklta7xkvwkh49zqn6mgnm2eyz589rkfvl3
  Ethereum    (ETH) 0xE4EA2A2356C04c8054Db452dCBd6f958F74722dE
";

                // Write version text
                WriteObject(text);

                // Stop checking for other parameters
                return;
            }
        }
    }
}
