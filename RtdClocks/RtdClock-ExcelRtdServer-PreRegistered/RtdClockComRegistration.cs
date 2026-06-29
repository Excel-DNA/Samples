using System;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace RtdClock_ExcelRtdServer_PreRegistered
{
    public static class RtdClockComRegistration
    {
        [ExcelCommand(MenuName = "RTD Clock", MenuText = "Register RTD COM Server")]
        public static void RegisterRtdClockComServer()
        {
            int result = ComServer.DllRegisterServer();
            if (result != 0)
            {
                throw new InvalidOperationException($"COM server registration failed with HRESULT 0x{result:X8}.");
            }
        }

        [ExcelCommand(MenuName = "RTD Clock", MenuText = "Unregister RTD COM Server")]
        public static void UnregisterRtdClockComServer()
        {
            int result = ComServer.DllUnregisterServer();
            if (result != 0)
            {
                throw new InvalidOperationException($"COM server unregistration failed with HRESULT 0x{result:X8}.");
            }
        }
    }
}
