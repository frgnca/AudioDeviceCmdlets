using System;
using System.Management.Automation;
using CoreAudioApi;
using System.Collections.Generic;
using System.Threading;

// Based on code posed to Code Project
// http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control
// by Ray Molenkamp
// and comments and suggestions by MadMidi

namespace AudioDeviceCmdlets
{
    public class AudioDevice
    {
        public int Index;
        public string DeviceFriendlyname;
        public MMDevice Device;

        public AudioDevice(int Index, MMDevice BaseDevice)
        {
            this.Index = Index;
            this.DeviceFriendlyname = BaseDevice.FriendlyName;
            this.Device = BaseDevice;
        }

        public override string ToString()
        {
            return string.Format("{0}, {1}", Index, DeviceFriendlyname);
        }
    }

    [Cmdlet(VerbsCommon.Get, "DefaultAudioDevice")]
    public class GetDefaultAudioDevice : Cmdlet
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
                    WriteObject(new AudioDevice(i, devices[i]));
                    return;
                }
            }
        }
    }

    [Cmdlet(VerbsCommon.Set, "DefaultAudioDevice")]
    public class SetDefaultAudioDevice : Cmdlet
    {
        [Alias("DeviceIndex")]
        [ValidateRange(0, 9)]
        [Parameter(Mandatory = true, Position = 0, ParameterSetName="Index")]
        public int Index
        {
            get { return index; }
            set { index = value; }
        }
        private int index;

        [Alias("DeviceName", "FriendlyName")]
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

            WriteObject(new AudioDevice(index, devices[index]));
        }
    }

    [Cmdlet(VerbsCommon.Get, "AudioDeviceList")]
    public class GetAudioDeviceList : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);

            for (int i = 0; i < devices.Count; i++)
            {
                WriteObject(new AudioDevice(i, devices[i]));
            }
        }
    }
    
    [Cmdlet(VerbsCommon.Get, "DefaultAudioDeviceVolume")]
    [Alias("vol","volume")]
    public class GetDefaultAudioDeviceVolume : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            WriteObject(string.Format("{0}%", defaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100));
        }
    }

    [Cmdlet(VerbsCommon.Set, "DefaultAudioDeviceVolume")]
    [Alias("setvol","setvolume")]
    public class SetDefaultAudioDeviceVolume : Cmdlet
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

    [Cmdlet(VerbsCommon.Get, "DefaultAudioDeviceMute")]
    [Alias("getmute")]
    public class GetDefaultAudioDeviceMute : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            WriteObject(defaultDevice.AudioEndpointVolume.Mute);
        }
    }

    [Cmdlet(VerbsCommon.Set, "DefaultAudioDeviceMute")]
    [Alias("mute")]
    public class SetDefaultAudioDeviceMute : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            MMDevice defaultDevice = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);

            defaultDevice.AudioEndpointVolume.Mute = !defaultDevice.AudioEndpointVolume.Mute;
        }
    }

    [Cmdlet(VerbsCommunications.Write, "DefaultAudioDeviceValue")]
    public class WriteDefaultAudioDeviceValue : Cmdlet
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
