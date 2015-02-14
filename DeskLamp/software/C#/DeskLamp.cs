using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;

namespace DeskLamp {
    /// <summary>
    /// Interface to the DMXControl DeskLamp
    /// </summary>
    public sealed class DeskLampInstance : IDisposable {
        private const uint DIGCF_PRESENT = 0x2;
        private const uint DIGCF_DEVICEINTERFACE = 0x10;
        private const uint FILE_SHARE_READ = 0x1;
        private const uint FILE_SHARE_WRITE = 0x2;
        private const uint OPEN_EXISTING = 3;

        private const ushort VENDOR_ID1 = 0x16C0; // Original DeskLamp
        private const ushort PRODUCT_ID1 = 0x5DF;
        private const ushort VENDOR_ID = 0x16D0;  // DMXConrol Projects "official" DeskLamp
        private const ushort PRODUCT_ID = 0x831;
        private const string VENDOR_NAME = "www.dmxcontrol.de";
        private const string PRODUCT_NAME = "DeskLamp";

        private IntPtr _HIDHandle = IntPtr.Zero;
        private byte _brightness = 0;
        private byte _strobe = 0;
        private String _id = "";
        private int _version = 0;

        private bool _enabled = false;

        /// <summary>
        /// Creates a new Desklamp Instance
        /// </summary>
        public DeskLampInstance() : this(false, "") {
        }

        /// <summary>
        /// Creates a new Desklamp Instance
        /// </summary>
        /// <param name="id">The ID if the DeskLamp</param>
        public DeskLampInstance(String id) : this (true, id) {
        }

        /// <summary>
        /// Creates a new Desklamp Instance
        /// </summary>
        /// <param name="enabled">Initial value of the Enabled Property</param>
        /// <param name="id">The ID if the DeskLamp</param>
        public DeskLampInstance(bool enabled, String id) {
            this.ID = id;
            this.Enabled = enabled;
        }

        ~DeskLampInstance() {
            this.Dispose();
        }

        /// <summary>
        /// Checks whether the Desklamp is Available
        /// </summary>
        public bool IsAvailable {
            get { return OpenIfRequiredCheckAvailable(); }
        }

        private bool OpenIfRequiredCheckAvailable() {
            try {
                if (System.Environment.OSVersion.Platform == PlatformID.Unix || System.Environment.OSVersion.Platform == PlatformID.MacOSX) {
                    return false; // No Desklamp Support under Linux/MacOS yet
                }
                if (this._HIDHandle == IntPtr.Zero) {
                    this._HIDHandle = OpenHIDDevice(this.ID, null, out _version);
                }
                return this._HIDHandle != IntPtr.Zero;
            } finally {
                if (!this.Enabled) {
                    Close();
                }
            }
        }

        /// <summary>
        /// The ID of the Desklamp
        /// </summary>
        public String ID {
            get { return this._id; }
            set {
                if (value != this._id) {
                    this._id = value;
                    Close();
                    WriteBrightness(this.Brightness);
                }
            }
        }

        public int Version {
            get { return this._version; }
        }

        /// <summary>
        /// Enables or disables the Desklamp
        /// </summary>
        public bool Enabled {
            get { return this._enabled; }
            set {
                if (!value && this._enabled) { //Switch from on to off
                    WriteBrightnessVerified(0);
                }
                this._enabled = value;
                if (!this._enabled) {
                    Close();
                } else {
                    WriteBrightnessVerified(this.Brightness);
                }
            }
        }

        /// <summary>
        /// Verifys Available by writing the Brightness Value
        /// </summary>
        public bool VerifyAvailable {
            get { return WriteBrightnessVerified(this.Brightness); }
        }

        /// <summary>
        /// Gets or Sets the Brightness of the Desklamp
        /// </summary>
        public byte Brightness {
            get {
                if (_version < 2) {
                    return this._brightness;
                }

                byte[] sendBuffer = new byte[] { 
                    6, // Get Dimmer
                    0
                };
                if (!DeskLampInterface.HidD_GetFeature(this._HIDHandle, sendBuffer, sendBuffer.Length)) {
                    Close();
                    return 0;
                }
                return sendBuffer[1];
            }
            set {
                this._brightness = value;
                WriteBrightnessVerified(value);
            }
        }

        private bool WriteBrightnessVerified(byte brightness) {
            bool b = WriteBrightness(brightness);
            if (!b && Enabled) { //Retry once
                b = WriteBrightness(brightness);
            }
            return b;
        }

        private bool WriteBrightness(byte brighness) {
            if (!Enabled || !IsAvailable) {
                return false;
            }

            if (_version == 1) {
                byte[] sendBuffer = new byte[] { 1, brighness };
                uint bytesWritten;
                System.Threading.NativeOverlapped ovl = new System.Threading.NativeOverlapped();
                if (!DeskLampInterface.WriteFile(this._HIDHandle, sendBuffer, (uint)sendBuffer.Length, out bytesWritten, ref ovl)) {
                    Close();
                    return false;
                }
                return true;
            } else if (_version == 2) {
                byte[] sendBuffer = new byte[] { 
                    2, // Report ID - Set Dimmer
                    brighness,
                    0, // Report length is 5, therefore fill with zeroes
                    0,
                    0
                };
                uint bytesWritten;
                System.Threading.NativeOverlapped ovl = new System.Threading.NativeOverlapped();
                if (!DeskLampInterface.WriteFile(this._HIDHandle, sendBuffer, (uint)sendBuffer.Length, out bytesWritten, ref ovl)) {
                    Close();
                    return false;
                }
                return true;
            }
            return false;
        }

        public bool IsRGB {
            get {
                if (_version < 2) {
                    return false;
                }
                byte[] sendBuffer = new byte[] { 
                    8, // Get Colormode
                    0
                };
                if (!DeskLampInterface.HidD_GetFeature(this._HIDHandle, sendBuffer, sendBuffer.Length)) {
                    Close();
                    return false;
                }
                return (sendBuffer[1] == 0); // Colormode 0 = RGB
            }
        }

        public bool ExternalUSBConnected {
            get {
                if (_version < 2) {
                    return false;
                }
                byte[] sendBuffer = new byte[] { 
                    9, // Get Ext-USB
                    0
                };
                if (!DeskLampInterface.HidD_GetFeature(this._HIDHandle, sendBuffer, sendBuffer.Length)) {
                    Close();
                    return false;
                }
                return (sendBuffer[1] == 1);
            }
        }

        public byte Strobe {
            get { return _strobe; }
            set {
                if (_version < 2) {
                    return;
                }

                this._strobe = value;
                byte[] sendBuffer = new byte[] { 
                    4, // Report ID - Set Strobe
                    _strobe,
                    0, // Report length is 5, therefore fill with zeroes
                    0,
                    0
                };
                uint bytesWritten;
                System.Threading.NativeOverlapped ovl = new System.Threading.NativeOverlapped();
                if (!DeskLampInterface.WriteFile(this._HIDHandle, sendBuffer, (uint)sendBuffer.Length, out bytesWritten, ref ovl)) {
                    Close();
                }
            }
        }

        public Color Color {
            get {
                if (_version < 2) {
                    return Color.White;
                }
                byte[] sendBuffer = new byte[] { 
                    7, // Get RGB
                    0,
                    0,
                    0
                };
                if (!DeskLampInterface.HidD_GetFeature(this._HIDHandle, sendBuffer, sendBuffer.Length)) {
                    Close();
                    return Color.White;
                }
                return Color.FromArgb(sendBuffer[1], sendBuffer[2], sendBuffer[3]);
            }
            set {
                if (_version < 2) {
                    return;
                }

                byte[] sendBuffer = new byte[] { 
                    3, // Report ID - Set RGB
                    value.R,
                    value.G, 
                    value.B,
                    0 // Report length is 5, therefore fill with zeroes
                };
                uint bytesWritten;
                System.Threading.NativeOverlapped ovl = new System.Threading.NativeOverlapped();
                if (!DeskLampInterface.WriteFile(this._HIDHandle, sendBuffer, (uint)sendBuffer.Length, out bytesWritten, ref ovl)) {
                    Close();
                }
            }
        }

        private void Close() {
            if (this._HIDHandle != IntPtr.Zero) {
                DeskLampInterface.CloseHandle(this._HIDHandle);
                this._HIDHandle = IntPtr.Zero;
            }
        }

        public static List<String> GetAvailableDeskLamps() {
            List<String> ret = new List<String>();
            int version;
            OpenHIDDevice(null, ret, out version);
            return ret;
        }

        private static IntPtr OpenHIDDevice(String id, List<String> devices, out int version) {
            version = 0;
            Guid hidGuid = new Guid();

            DeskLampInterface.SP_DEVICE_INTERFACE_DATA deviceInterfaceData = new DeskLampInterface.SP_DEVICE_INTERFACE_DATA();
            DeskLampInterface.SP_DEVINFO_DATA deviceInfoData = new DeskLampInterface.SP_DEVINFO_DATA(),
                                              dummyInfo = new DeskLampInterface.SP_DEVINFO_DATA();
            DeskLampInterface.SP_DEVICE_INTERFACE_DETAIL_DATA deviceInterfaceDetailData = new DeskLampInterface.SP_DEVICE_INTERFACE_DETAIL_DATA();

            DeskLampInterface.HIDD_ATTRIBUTES deviceAttributes = new DeskLampInterface.HIDD_ATTRIBUTES();

            /******************************************************************************
             * HidD_GetHidGuid
             * Get the GUID for all system HIDs.
             * Returns: the GUID in HidGuid.
             * The routine doesn't return a value in Result
             * but the routine is declared as a function for consistency with the other API calls.
             ******************************************************************************/
            DeskLampInterface.HidD_GetHidGuid(ref hidGuid);

            /******************************************************************************
             * SetupDiGetClassDevs
             * Returns: a handle to a device information set for all installed devices.
             * Requires: the HidGuid returned in GetHidGuid.
             ******************************************************************************/
            IntPtr deviceInfoList = DeskLampInterface.SetupDiGetClassDevs(ref hidGuid, null, IntPtr.Zero, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
            IntPtr tmp = IntPtr.Zero, handle = IntPtr.Zero;
            uint deviceCnt = 0;
            if (devices != null) {
                devices.Clear();
            }
            for (uint memberIndex = 0; ; memberIndex++) {
                if (tmp != IntPtr.Zero) {
                    DeskLampInterface.CloseHandle(tmp);
                    tmp = IntPtr.Zero;
                }

                deviceInterfaceData.cbSize = (uint)Marshal.SizeOf(deviceInterfaceData);

                /******************************************************************************
                 * SetupDiEnumDeviceInterfaces
                 * On return, DeviceInterfaceData contains the handle to a
                 * SP_DEVICE_INTERFACE_DATA structure for a detected device.
                 * Requires:
                 * the DeviceInfoList returned in SetupDiGetClassDevs.
                 * the HidGuid returned in GetHidGuid.
                 * An index to specify a device.
                 ******************************************************************************/
                if (!DeskLampInterface.SetupDiEnumDeviceInterfaces(deviceInfoList, IntPtr.Zero, ref hidGuid, memberIndex, ref deviceInterfaceData)) {
                    break;
                }

                deviceInfoData.cbSize = (uint)Marshal.SizeOf(deviceInfoData);
                uint needed, detailData;

                dummyInfo.cbSize = (uint)Marshal.SizeOf(dummyInfo);
                /******************************************************************************
                 * SetupDiGetDeviceInterfaceDetail
                 * Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure
                 * containing information about a device.
                 * To retrieve the information, call this function twice.
                 * The first time returns the size of the structure in Needed.
                 * The second time returns a pointer to the data in DeviceInfoList.
                 * Requires:
                 * A DeviceInfoList returned by SetupDiGetClassDevs and
                 * an SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces.
                 *******************************************************************************/
                DeskLampInterface.SetupDiGetDeviceInterfaceDetail(deviceInfoList, ref deviceInterfaceData,
                        IntPtr.Zero, 0, out needed, IntPtr.Zero);
                detailData = needed;

                if (IntPtr.Size == 8) { // for 64 bit operating systems
                    deviceInterfaceDetailData.cbSize = 8;
                } else {
                    deviceInterfaceDetailData.cbSize = 4 + (uint)Marshal.SystemDefaultCharSize; // for 32 bit systems
                }
                DeskLampInterface.SetupDiGetDeviceInterfaceDetail(deviceInfoList, ref deviceInterfaceData,
                    ref deviceInterfaceDetailData, detailData, out needed, IntPtr.Zero);

                /******************************************************************************
                 * CreateFile
                 * Returns: a handle that enables reading and writing to the device.
                 * Requires:
                 * The DevicePathName returned by SetupDiGetDeviceInterfaceDetail.
                 ******************************************************************************/
                tmp = DeskLampInterface.CreateFile(deviceInterfaceDetailData.DevicePath, FileAccess.Read | FileAccess.Write,
                     FileShare.Read | FileShare.Write, IntPtr.Zero, FileMode.Open, 0, IntPtr.Zero);
                deviceAttributes.Size = (uint)Marshal.SizeOf(deviceAttributes);

                /******************************************************************************
                 * HidD_GetAttributes
                 * Requests information from the device.
                 * Requires: The handle returned by CreateFile.
                 * Returns: an HIDD_ATTRIBUTES structure containing
                 * the Vendor ID, Product ID, and Product Version Number.
                 * Use this information to determine if the detected device
                 * is the one we're looking for.
                 ******************************************************************************/
                DeskLampInterface.HidD_GetAttributes(tmp, ref deviceAttributes);

                if ((deviceAttributes.VendorID == VENDOR_ID && deviceAttributes.ProductID == PRODUCT_ID) || (deviceAttributes.VendorID == VENDOR_ID1 && deviceAttributes.ProductID == PRODUCT_ID1)) {
                    StringBuilder buffer = new StringBuilder();
                    if (DeskLampInterface.HidD_GetManufacturerString(tmp, buffer, 512)) {
                        if (VENDOR_NAME.Equals(buffer.ToString())) {
                            buffer = new StringBuilder();
                            if (DeskLampInterface.HidD_GetProductString(tmp, buffer, 512)) {
                                if (PRODUCT_NAME.Equals(buffer.ToString())) {
                                    String thisID;
                                    if (deviceAttributes.ProductID == PRODUCT_ID) {
                                        buffer = new StringBuilder();
                                        if (DeskLampInterface.HidD_GetSerialNumberString(tmp, buffer, 512)) {
                                            thisID = buffer.ToString();
                                        } else {
                                            thisID = deviceCnt.ToString();
                                        }
                                        version = 2;
                                    } else {
                                        thisID = deviceCnt.ToString();
                                        version = 1;
                                    }

                                    if (devices != null) {
                                        devices.Add(thisID);
                                    }

                                    if (thisID.Equals(id)) {
                                        handle = tmp;
                                        tmp = IntPtr.Zero;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            DeskLampInterface.SetupDiDestroyDeviceInfoList(deviceInfoList);

            return handle;
        }

        #region IDisposable Member

        public void Dispose() {
            Close();
        }

        #endregion
    }

    static class DeskLampInterface
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        [DllImport("hid.dll", SetLastError = true)]
        public static extern void HidD_GetHidGuid(ref Guid hidGuid);

        [DllImport("setupapi.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SetupDiGetClassDevs(
                                              ref Guid ClassGuid,
                                              [MarshalAs(UnmanagedType.LPTStr)] string Enumerator,
                                              IntPtr hwndParent,
                                              uint Flags
                                             );

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetupDiEnumDeviceInterfaces(
                                       IntPtr hDevInfo,
                                       IntPtr devInfo,
                                       ref Guid interfaceClassGuid,
                                       uint memberIndex,
                                       ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData
                                    );

        [DllImport(@"setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetupDiGetDeviceInterfaceDetail(
                                       IntPtr hDevInfo,
                                       ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
                                       IntPtr deviceInterfaceDetailData,
                                       UInt32 deviceInterfaceDetailDataSize,
                                       out UInt32 requiredSize,
                                       IntPtr deviceInfoData
                                    );

        [DllImport(@"setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetupDiGetDeviceInterfaceDetail(
                                       IntPtr hDevInfo,
                                       ref SP_DEVICE_INTERFACE_DATA deviceInterfaceData,
                                       ref SP_DEVICE_INTERFACE_DETAIL_DATA deviceInterfaceDetailData,
                                       UInt32 deviceInterfaceDetailDataSize,
                                       out UInt32 requiredSize,
                                       IntPtr deviceInfoData
                                    );

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateFile(
            string fileName,
            [MarshalAs(UnmanagedType.U4)] FileAccess fileAccess,
            [MarshalAs(UnmanagedType.U4)] FileShare fileShare,
            IntPtr securityAttributes,
            [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
            int flags,
            IntPtr template);

        [DllImport("hid.dll", SetLastError = true)]
        public static extern bool HidD_GetAttributes(IntPtr HidDeviceObject, ref HIDD_ATTRIBUTES Attributes);

        [DllImport("hid.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool HidD_GetManufacturerString(IntPtr HidDeviceObject, StringBuilder Buffer, int BufferLength);

        [DllImport("hid.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool HidD_GetProductString(IntPtr HidDeviceObject, StringBuilder Buffer, int BufferLength);

        [DllImport("hid.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool HidD_GetSerialNumberString(IntPtr HidDeviceObject, StringBuilder Buffer, int BufferLength);

        [DllImport("hid.dll", SetLastError = true)]
        public static extern bool HidD_GetFeature(IntPtr HidDeviceObject, Byte[] lpReportBuffer, Int32 ReportBufferLength);

        [DllImport("setupapi.dll", SetLastError = true)]
        public static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("kernel32.dll")]
        public static extern bool WriteFile(IntPtr hFile, byte[] lpBuffer, uint nNumberOfBytesToWrite, out uint lpNumberOfBytesWritten,
           [In] ref System.Threading.NativeOverlapped lpOverlapped);

       
        [StructLayout(LayoutKind.Sequential)]
        internal struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public IntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SP_DEVICE_INTERFACE_DATA
        {
            public uint cbSize;
            public Guid InterfaceClassGuid;
            public uint Flags;
            public IntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct SP_DEVICE_INTERFACE_DETAIL_DATA
        {
            public uint cbSize;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string DevicePath;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct HIDD_ATTRIBUTES
        {
            public uint Size;
            public ushort VendorID;
            public ushort ProductID;
            public ushort VersionNumber;
        }
    }
}
