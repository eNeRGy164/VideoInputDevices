using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Hompus.VideoInputDevices
{
    public class SystemDeviceEnumerator : IDisposable
    {
        private bool disposed;
        private ICreateDevEnum _systemDeviceEnumerator;

        public SystemDeviceEnumerator()
        {
            var comType = Type.GetTypeFromCLSID(new Guid("62BE5D10-60EB-11D0-BD3B-00A0C911CE86"));
            _systemDeviceEnumerator = (ICreateDevEnum)Activator.CreateInstance(comType);
        }

        public IReadOnlyDictionary<int, string> ListVideoInputDevice()
        {
            var videoInputDeviceClass = new Guid("{860BB310-5D01-11D0-BD3B-00A0C911CE86}");

            var hresult = _systemDeviceEnumerator.CreateClassEnumerator(ref videoInputDeviceClass, out var enumMoniker, 0);
            if (hresult != 0)
            {
                throw new ApplicationException("No devices of the category");
            }

            var moniker = new IMoniker[1];
            var list = new Dictionary<int, string>();

            while (true)
            {
                hresult = enumMoniker.Next(1, moniker, IntPtr.Zero);
                if ((hresult != 0) || (moniker[0] == null))
                {
                    break;
                }

                var device = new VideoInputDevice(moniker[0]);
                list.Add(list.Count, device.Name);

                // Release COM object
                Marshal.ReleaseComObject(moniker[0]);
                moniker[0] = null;
            }

            return list;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    Marshal.ReleaseComObject(_systemDeviceEnumerator);
                    _systemDeviceEnumerator = null;
                }

                disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
