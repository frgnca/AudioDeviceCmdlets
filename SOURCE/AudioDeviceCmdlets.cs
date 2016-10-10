/*
The MIT License (MIT)
Copyright (c) 2016 Francois Gendron

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using System;
using System.Threading;
using System.Management.Automation;
using CoreAudioApi;

namespace AudioDeviceCmdlets
{
    public class AudioDevice
    {
        public string Type;
        public int Index;
        public string Name;
        public MMDevice Device;

        public AudioDevice(string Type, int Index, MMDevice BaseDevice)
        {
            this.Type = Type;
            this.Index = Index;
            this.Name = BaseDevice.FriendlyName;
            this.Device = BaseDevice;
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}", Index, Name);
        }
    }

    // Cmdlet to get a list of all audio devices
    [Cmdlet(VerbsCommon.Get, "AudioDeviceList")]
    public class GetAudioDeviceList : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection playbackdevices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDeviceCollection recordingdevices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice DefaultPlaybackDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            MMDevice DefaultRecordingDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            for (int i = 0; i < playbackdevices.Count; i++)
            {
                WriteObject(new AudioDevice("Playback", i, playbackdevices[i]));
            }

            for (int i = 0; i < recordingdevices.Count; i++)
            {
                WriteObject(new AudioDevice("Recording", i, recordingdevices[i]));
            }
        }
    }

    // Cmdlet to get a list of all default audio devices
    [Cmdlet(VerbsCommon.Get, "AudioDeviceDefaultList")]
    public class GetAudioDeviceDefaultList : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection playbackdevices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDeviceCollection recordingdevices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice DefaultPlaybackDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            MMDevice DefaultRecordingDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            for (int i = 0; i < playbackdevices.Count; i++)
            {
                if (playbackdevices[i].ID == DefaultPlaybackDevice.ID)
                {
                    WriteObject(new AudioDevice("Playback", i, playbackdevices[i]));
                    break;
                }
            }

            for (int i = 0; i < recordingdevices.Count; i++)
            {
                if (recordingdevices[i].ID == DefaultRecordingDevice.ID)
                {
                    WriteObject(new AudioDevice("Recording", i, recordingdevices[i]));
                    break;
                }
            }
        }
    }

    // Cmdlet to get a list of all playback audio devices
    [Cmdlet(VerbsCommon.Get, "AudioDevicePlaybackList")]
    public class GetAudioDevicePlaybackList : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);

            for (int i = 0; i < devices.Count; i++)
            {
                WriteObject(new AudioDevice("Playback", i, devices[i]));
            }
        }
    }

    // Cmdlet to get the default playback audio device
    [Cmdlet(VerbsCommon.Get, "AudioDevicePlaybackDefault")]
    public class GetAudioDevicePlaybackDefault : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice DefaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            for (int i = 0; i < devices.Count; i++)
            {
                if (devices[i].ID == DefaultDevice.ID)
                {
                    WriteObject(new AudioDevice("Playback", i, devices[i]));
                    return;
                }
            }
        }
    }

    // Cmdlet to set the default playback audio device
    [Cmdlet(VerbsCommon.Set, "AudioDevicePlaybackDefault")]
    public class SetAudioDevicePlaybackDefault : Cmdlet
    {
        [ValidateRange(0, 9)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName="Index")]
        public int Index
        {
            get { return index; }
            set { index = value; }
        }
        private int index;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Name")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        private string name;

        [Parameter(Mandatory=true, ParameterSetName="InputObject", ValueFromPipeline=true)]
        public AudioDevice InputObject 
        {
            get { return inputObject;}
            set { inputObject = value; }
        
        }
        private AudioDevice inputObject;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            
            PolicyConfigClient client = new PolicyConfigClient();

            if (!string.IsNullOrEmpty(name))
            {
                for (int i = 0; i < devices.Count; i++)
                {
                    if (string.Compare(devices[i].FriendlyName, name, StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        index = i;
                        break;
                    }
                }
            }

            if (inputObject != null)
            {
                for (int i = 0; i < devices.Count; i++)
                {
                    if (devices[i].ID == inputObject.Device.ID)
                    {
                        index = i;
                        break;
                    }
                }
            }

            client.SetDefaultEndpoint(devices[index].ID, ERole.eCommunications);
            client.SetDefaultEndpoint(devices[index].ID, ERole.eMultimedia);

            WriteObject(new AudioDevice("Playback", index, devices[index]));
        }
    }

    // Cmdlet to get the mute state of the default playback audio device
    [Cmdlet(VerbsCommon.Get, "AudioDevicePlaybackDefaultMute")]
    public class GetAudioDevicePlaybackDefaultMute : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
 
            WriteObject(defaultDevice.AudioEndpointVolume.Mute);
        }
}

    // Cmdlet to set the mute state of the default playback audio device
    [Cmdlet(VerbsCommon.Set, "AudioDevicePlaybackDefaultMute")]
    public class SetAudioDevicePlaybackDefaultMute : Cmdlet
    {
        [Parameter(Position = 0)]
        public bool? Mute
        {
            get { return mute; }
            set { mute = value; }
        }
        private bool? mute;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            if(mute != null)
            {
                defaultDevice.AudioEndpointVolume.Mute = (bool)mute;
            }
            else
            {
                defaultDevice.AudioEndpointVolume.Mute = !defaultDevice.AudioEndpointVolume.Mute;
            }
        }
    }

    // Cmdlet to get the volume level of the default playback audio device
    [Cmdlet(VerbsCommon.Get, "AudioDevicePlaybackDefaultVolume")]
    public class GetAudioDevicePlaybackDefaultVolume : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            WriteObject(string.Format("{0}%", defaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100));
        }
    }

    // Cmdlet to set the volume level of the default playback audio device
    [Cmdlet(VerbsCommon.Set, "AudioDevicePlaybackDefaultVolume")]
    public class SetAudioDevicePlaybackDefaultVolume : Cmdlet
    {      
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0)]
        public float Volume
        {
            get { return volume; }
            set { volume = value; }
        }
        private float volume;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            defaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volume / 100.0f;
        }
    }

    // Cmdlet to get a list of all recording audio devices
    [Cmdlet(VerbsCommon.Get, "AudioDeviceRecordingList")]
    public class GetAudioDeviceRecordingList : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);

            for (int i = 0; i < devices.Count; i++)
            {
                WriteObject(new AudioDevice("Recording", i, devices[i]));
            }
        }
    }

    // Cmdlet to get the default recording audio device
    [Cmdlet(VerbsCommon.Get, "AudioDeviceRecordingDefault")]
    public class GetAudioDeviceRecordingDefault : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice DefaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            for (int i = 0; i < devices.Count; i++)
            {
                if (devices[i].ID == DefaultDevice.ID)
                {
                    WriteObject(new AudioDevice("Recording", i, devices[i]));
                    return;
                }
            }
        }
    }

    // Cmdlet to set the default recording audio device
    [Cmdlet(VerbsCommon.Set, "AudioDeviceRecordingDefault")]
    public class SetAudioDeviceRecordingDefault : Cmdlet
    {
        [ValidateRange(0, 9)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Index")]
        public int Index
        {
            get { return index; }
            set { index = value; }
        }
        private int index;

        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Name")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        private string name;

        [Parameter(Mandatory = true, ParameterSetName = "InputObject", ValueFromPipeline = true)]
        public AudioDevice InputObject
        {
            get { return inputObject; }
            set { inputObject = value; }

        }
        private AudioDevice inputObject;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);

            PolicyConfigClient client = new PolicyConfigClient();

            if (!string.IsNullOrEmpty(name))
            {
                for (int i = 0; i < devices.Count; i++)
                {
                    if (string.Compare(devices[i].FriendlyName, name, StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        index = i;
                        break;
                    }
                }
            }

            if (inputObject != null)
            {
                for (int i = 0; i < devices.Count; i++)
                {
                    if (devices[i].ID == inputObject.Device.ID)
                    {
                        index = i;
                        break;
                    }
                }
            }

            client.SetDefaultEndpoint(devices[index].ID, ERole.eCommunications);
            client.SetDefaultEndpoint(devices[index].ID, ERole.eMultimedia);

            WriteObject(new AudioDevice("Recording", index, devices[index]));
        }
    }

    // Cmdlet to get the mute state of the default recording audio device
    [Cmdlet(VerbsCommon.Get, "AudioDeviceRecordingDefaultMute")]
    public class GetAudioDeviceRecordingDefaultMute : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            WriteObject(defaultDevice.AudioEndpointVolume.Mute);
        }
    }

    // Cmdlet to set the mute state of the default recording audio device
    [Cmdlet(VerbsCommon.Set, "AudioDeviceRecordingDefaultMute")]
    public class SetAudioDeviceRecordingDefaultMute : Cmdlet
    {
        [Parameter(Position = 0)]
        public bool? Mute
        {
            get { return mute; }
            set { mute = value; }
        }
        private bool? mute;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            if (mute != null)
            {
                defaultDevice.AudioEndpointVolume.Mute = (bool)mute;
            }
            else
            {
                defaultDevice.AudioEndpointVolume.Mute = !defaultDevice.AudioEndpointVolume.Mute;
            }
        }
    }

    // Cmdlet to get the volume level of the default recording audio device
    [Cmdlet(VerbsCommon.Get, "AudioDeviceRecordingDefaultVolume")]
    public class GetAudioDeviceRecordingDefaultVolume : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            WriteObject(string.Format("{0}%", defaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100));
        }
    }

    // Cmdlet to set the volume level of the default recording audio device
    [Cmdlet(VerbsCommon.Set, "AudioDeviceRecordingDefaultVolume")]
    public class SetAudioDeviceRecordingDefaultVolume : Cmdlet
    {
        [ValidateRange(0, 100.0f)]
        [Parameter(Mandatory = true, Position = 0)]
        public float Volume
        {
            get { return volume; }
            set { volume = value; }
        }
        private float volume;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia);

            defaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volume / 100.0f;
        }
    }


    [Cmdlet(VerbsCommunications.Write, "AudioDeviceDefaultValue")]
    public class WriteAudioDeviceDefaultValue : Cmdlet
    {
        [Parameter(Position=0)]
        public SwitchParameter StreamValue
        {
            get { return streamValue; }
            set { streamValue = value; }
        }
        private bool streamValue;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            ProgressRecord pr = new ProgressRecord(0, defaultDevice.FriendlyName, "Peak Value");
            pr.PercentComplete = 0;

            if (streamValue)
                WriteObject(0);
            else
                WriteProgress(pr);
            
            do{
                pr.PercentComplete = System.Convert.ToInt32( defaultDevice.AudioMeterInformation.MasterPeakValue * 100);

                if (streamValue)
                    WriteObject(System.Convert.ToInt32(defaultDevice.AudioMeterInformation.MasterPeakValue * 100));
                else
                    WriteProgress(pr);
                
                Thread.Sleep(100);

            } while (!Stopping);
        }
    }
}
