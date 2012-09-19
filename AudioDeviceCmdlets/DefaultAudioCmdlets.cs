using System.Management.Automation;
using CoreAudioApi;
using System.Collections.Generic;

namespace AudioDeviceCmdlets
{
    class ReturnObject
    {
        public int Index;
        public MMDevice Device;
    }

    [Cmdlet(VerbsCommon.Get, "DefaultAudioDevice")]
    class GetDefaultAudioDevice : Cmdlet
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

    [Cmdlet(VerbsCommon.Get, "AudioDeviceList")]
    class GetAudioDeviceList : Cmdlet
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
}