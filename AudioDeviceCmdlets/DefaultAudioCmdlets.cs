using System.Management.Automation;
using CoreAudioApi;
using System.Collections.Generic;

// Based on code posed to Code Project
// http://www.codeproject.com/Articles/18520/Vista-Core-Audio-API-Master-Volume-Control
// by Ray Molenkamp
// and comments and suggestions by MadMidi

namespace AudioDeviceCmdlets
{
    class ReturnObject
    {
        public int Index;
        public MMDevice Device;
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
                    WriteObject(new ReturnObject { Index = i, Device = devices[i] });
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
        [Parameter(Mandatory = true, Position = 0)]
        public int Index
        {
            get { return index; }
            set { index = value; }
        }
        private int index;

        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);
            
            PolicyConfigClient client = new PolicyConfigClient();

            client.SetDefaultEndpoint(devices[index].ID, ERole.eCommunications);
            client.SetDefaultEndpoint(devices[index].ID, ERole.eMultimedia);

            WriteObject(devices[index]);
        }
    }

    [Cmdlet(VerbsCommon.Get, "AudioDeviceList")]
    public class GetAudioDeviceList : Cmdlet
    {
        protected override void ProcessRecord()
        {
            MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
            MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eRender, EDeviceState.DEVICE_STATE_ACTIVE);

            List<ReturnObject> resultObjectList = new List<ReturnObject>();

            for (int i = 0; i < devices.Count; i++)
            {
                resultObjectList.Add(new ReturnObject{Index = i, Device = devices[i]});
            }

            WriteObject(resultObjectList);
        }
    }

    [Cmdlet(VerbsCommon.Set, "DefaultAudioDeviceVolume")]
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

}