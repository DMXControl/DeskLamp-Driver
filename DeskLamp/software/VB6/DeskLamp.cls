VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeskLamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal Count As Long) As Long
'Declared as a function for consistency,
'but returns nothing. (Ignore the returned value.)
Private Declare Function HidD_GetHidGuid Lib "hid.dll" (ByRef HidGuid As GUID) As Long
Private Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByRef ClassGuid As GUID, ByVal Enumerator As String, ByVal hWndParent As Long, ByVal Flags As Long) As Long
Private Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
Private Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As Long, ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function HidD_GetAttributes Lib "hid.dll" (ByVal HidDeviceObject As Long, ByRef Attributes As HIDD_ATTRIBUTES) As Long
Private Declare Function HidD_GetManufacturerString Lib "hid.dll" (ByVal HidDeviceObject As Long, ByVal UBuff As String, ByVal UBuffLen As Long) As Long
Private Declare Function HidD_GetProductString Lib "hid.dll" (ByVal HidDeviceObject As Long, ByVal UBuff As String, ByVal UBuffLen As Long) As Long
Private Declare Function HidD_GetSerialNumberString Lib "hid.dll" (ByVal HidDeviceObject As Long, ByVal UBuff As String, ByVal UBuffLen As Long) As Long
Private Declare Function HidD_GetFeature Lib "hid.dll" (ByVal HidDeviceObject As Long, ByRef ReportBuffer As Byte, ByVal ReportBufferLength As Long) As Byte
Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type SP_DEVICE_INTERFACE_DATA
  cbSize As Long
  InterfaceClassGuid As GUID
  Flags As Long
  Reserved As Long
End Type

Private Type SP_DEVICE_INTERFACE_DETAIL_DATA
  cbSize As Long
  DevicePath As Byte
End Type

Private Type SP_DEVINFO_DATA
  cbSize As Long
  ClassGuid As GUID
  DevInst As Long
  Reserved As Long
End Type

Private Type HIDD_ATTRIBUTES
  Size As Long
  VendorID As Integer
  ProductID As Integer
  VersionNumber As Integer
End Type

Private Const DIGCF_PRESENT = &H2
Private Const DIGCF_DEVICEINTERFACE = &H10
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1

Private Const VendorID = &H16D0
Private Const VendorID1 = &H16C0
Private Const ProductID = &H831
Private Const ProductID1 = &H5DF
Private Const VendorName = "www.dmxcontrol.de"
Private Const ProductName = "DeskLamp"

Private Const OUT_REPORT_LENGTH As Long = 5

Private mHIDHandle As Long
Private mBrightness As Byte
Private mVersion As Byte
Public Property Get Brightness() As Byte
  Dim SendBuffer(0 To 7) As Byte
  Dim Result As Long
        
  If mVersion = 1 Then Brightness = mBrightness: Exit Property
    
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    SendBuffer(0) = 6 'Get Dimmer
    Result = HidD_GetFeature(mHIDHandle, SendBuffer(0), CLng(2))
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    Else
      Brightness = SendBuffer(1)
      mBrightness = Brightness
    End If
  End If
End Property

Public Property Get ExternUSB() As Boolean
  Dim SendBuffer(0 To 7) As Byte
  Dim Result As Long

  If mVersion = 1 Then ExternUSB = False: Exit Property

  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    SendBuffer(0) = 9 'Get Ext-USB
    Result = HidD_GetFeature(mHIDHandle, SendBuffer(0), CLng(2))
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    Else
      ExternUSB = (SendBuffer(1) = 1)
    End If
  End If
End Property

Public Property Get RGBEnabled() As Boolean
  Dim SendBuffer(0 To 7) As Byte
  Dim Result As Long
  
  If mVersion = 1 Then RGBEnabled = False: Exit Property
    
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    SendBuffer(0) = 8 'Get Colormode
    Result = HidD_GetFeature(mHIDHandle, SendBuffer(0), CLng(2))
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    Else
      RGBEnabled = (SendBuffer(1) = 0) ' Colormode 0 = RGB
    End If
  End If
End Property
Function OpenHIDDevice() As Long
  Dim Result As Long
  Dim HidGuid As GUID
  Dim Security As SECURITY_ATTRIBUTES
  Dim DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
  Dim DeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
  Dim DeviceInfoData As SP_DEVINFO_DATA
  Dim DeviceAttributes As HIDD_ATTRIBUTES
  Dim DeviceInfoList As Long
  Dim MemberIndex As Long
  Dim Needed As Long
  Dim DetailData As Long
  Dim DetailDataBuffer() As Byte
  Dim DevicePathName As String
  Dim Handle As Long
  Dim Buffer As String
  
  OpenHIDDevice = INVALID_HANDLE_VALUE
  
  'Values for SECURITY_ATTRIBUTES structure:
  Security.lpSecurityDescriptor = 0
  Security.bInheritHandle = True
  Security.nLength = Len(Security)

  '******************************************************************************
  'HidD_GetHidGuid
  'Get the GUID for all system HIDs.
  'Returns: the GUID in HidGuid.
  'The routine doesn't return a value in Result
  'but the routine is declared as a function for consistency with the other API calls.
  '******************************************************************************
  Result = HidD_GetHidGuid(HidGuid)

  '******************************************************************************
  'SetupDiGetClassDevs
  'Returns: a handle to a device information set for all installed devices.
  'Requires: the HidGuid returned in GetHidGuid.
  '******************************************************************************
  DeviceInfoList = SetupDiGetClassDevs(HidGuid, vbNullString, 0, (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))

  '******************************************************************************
  'SetupDiEnumDeviceInterfaces
  'On return, DeviceInterfaceData contains the handle to a
  'SP_DEVICE_INTERFACE_DATA structure for a detected device.
  'Requires:
  'the DeviceInfoList returned in SetupDiGetClassDevs.
  'the HidGuid returned in GetHidGuid.
  'An index to specify a device.
  '******************************************************************************

  'Begin with 0 and increment until no more devices are detected.
  Handle = INVALID_HANDLE_VALUE
  mVersion = 0
  MemberIndex = 0
  Do
    'Close remaining handle
    If (Handle <> INVALID_HANDLE_VALUE) Then
      CloseHandle Handle
      Handle = INVALID_HANDLE_VALUE
    End If

    'The cbSize element of the DeviceInterfaceData structure must be set to
    'the structure's size in bytes. The size is 28 bytes.
    DeviceInterfaceData.cbSize = LenB(DeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces(DeviceInfoList, 0, HidGuid, MemberIndex, DeviceInterfaceData)
    If Result = 0 Then Exit Do
    
    '******************************************************************************
    'SetupDiGetDeviceInterfaceDetail
    'Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure
    'containing information about a device.
    'To retrieve the information, call this function twice.
    'The first time returns the size of the structure in Needed.
    'The second time returns a pointer to the data in DeviceInfoList.
    'Requires:
    'A DeviceInfoList returned by SetupDiGetClassDevs and
    'an SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces.
    '*******************************************************************************
    DeviceInfoData.cbSize = Len(DeviceInfoData)
    Result = SetupDiGetDeviceInterfaceDetail(DeviceInfoList, DeviceInterfaceData, 0, 0, Needed, 0)
    DetailData = Needed
                    
    'Store the structure's size.
    DeviceInterfaceDetailData.cbSize = Len(DeviceInterfaceDetailData)
        
    'Use a byte array to allocate memory for
    'the DeviceInterfaceDetailData structure
    ReDim DetailDataBuffer(Needed)
       
    'Store cbSize in the first four bytes of the array.
    Call RtlMoveMemory(DetailDataBuffer(0), DeviceInterfaceDetailData, 4)
        
    'Call SetupDiGetDeviceInterfaceDetail again.
    'This time, pass the address of the first element of DetailDataBuffer
    'and the returned required buffer size in DetailData.
    Result = SetupDiGetDeviceInterfaceDetail(DeviceInfoList, DeviceInterfaceData, VarPtr(DetailDataBuffer(0)), DetailData, Needed, 0)
        
    'Convert the byte array to a string.
    DevicePathName = CStr(DetailDataBuffer())
    'Convert to Unicode.
    DevicePathName = StrConv(DevicePathName, vbUnicode)
    'Strip cbSize (4 bytes) from the beginning.
    DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
                
    '******************************************************************************
    'CreateFile
    'Returns: a handle that enables reading and writing to the device.
    'Requires:
    'The DevicePathName returned by SetupDiGetDeviceInterfaceDetail.
    '******************************************************************************
    Handle = CreateFile(DevicePathName, GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE), Security, OPEN_EXISTING, 0&, 0)
                    
    'Now we can find out if it's the device we're looking for.
        
    '******************************************************************************
    'HidD_GetAttributes
    'Requests information from the device.
    'Requires: The handle returned by CreateFile.
    'Returns: an HIDD_ATTRIBUTES structure containing
    'the Vendor ID, Product ID, and Product Version Number.
    'Use this information to determine if the detected device
    'is the one we're looking for.
    '******************************************************************************
        
    'Set the Size property to the number of bytes in the structure.
    DeviceAttributes.Size = LenB(DeviceAttributes)
    Result = HidD_GetAttributes(Handle, DeviceAttributes)
                    
    'Find out if the device matches the one we're looking for.
    If ((DeviceAttributes.VendorID = VendorID) And (DeviceAttributes.ProductID = ProductID)) Or ((DeviceAttributes.VendorID = VendorID1) And (DeviceAttributes.ProductID = ProductID1)) Then
      'At least VID and PID match...
      Buffer = Space(512)
      If HidD_GetManufacturerString(Handle, Buffer, 512) Then
        Buffer = StrConv(Buffer, vbFromUnicode)
        Buffer = Left(Buffer, InStrRev(Buffer, Chr(0)) - 1)
        If Buffer = VendorName Then
          Buffer = Space(512)
          If HidD_GetProductString(Handle, Buffer, 512) Then
            Buffer = StrConv(Buffer, vbFromUnicode)
            Buffer = Left(Buffer, InStrRev(Buffer, Chr(0)) - 1)
            If Buffer = ProductName Then
              If (DeviceAttributes.ProductID = ProductID) Then
                If HidD_GetSerialNumberString(Handle, Buffer, 512) Then
                    Buffer = StrConv(Buffer, vbFromUnicode)
                    'Buffer = Left(Buffer, InStrRev(Buffer, Chr(0)) - 1)
                    'MsgBox "Serial number: " & Buffer
                End If
               mVersion = 2
              ElseIf (DeviceAttributes.ProductID = ProductID1) Then
               mVersion = 1
              End If
              Exit Do
            End If
          End If
        End If
      End If
    End If
    
    'Keep looking until we find the device or there are no more left to examine.
    MemberIndex = MemberIndex + 1
  Loop

  'Free the memory reserved for the DeviceInfoList returned by SetupDiGetClassDevs.
  Result = SetupDiDestroyDeviceInfoList(DeviceInfoList)

  If (Handle <> INVALID_HANDLE_VALUE) Then
    OpenHIDDevice = Handle
  End If
End Function

Public Function isAvailable() As Boolean
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  isAvailable = (mHIDHandle <> INVALID_HANDLE_VALUE)
End Function

Public Property Let Brightness(Brightness As Byte)
  Dim SendBuffer(0 To 7) As Byte
  Dim NumberOfBytesWritten As Long
  Dim Result As Long
    
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
   If mVersion = 1 Then
    'The first byte is the Report ID
    SendBuffer(0) = 1

    'The next bytes are data
    SendBuffer(1) = Brightness

    NumberOfBytesWritten = 0

    Result = WriteFile(mHIDHandle, SendBuffer(0), CLng(2), NumberOfBytesWritten, 0)
   Else
    'The first byte is the Report ID
    SendBuffer(0) = 2 'Set Dimmer

    'The next bytes are data
    SendBuffer(1) = Brightness
    'SendBuffer(1) = &HCA
    'SendBuffer(2) = &HFE
    'SendBuffer(3) = &HBA
    'SendBuffer(4) = &HBE
    
    NumberOfBytesWritten = 0

    Result = WriteFile(mHIDHandle, SendBuffer(0), OUT_REPORT_LENGTH, NumberOfBytesWritten, 0)
   End If
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    Else
      mBrightness = Brightness
    End If
  End If
End Property
Public Property Let Strobe(Strobe As Byte)
  Dim SendBuffer(0 To 7) As Byte
  Dim NumberOfBytesWritten As Long
  Dim Result As Long
    
  If mVersion = 1 Then Exit Property
  
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    'The first byte is the Report ID
    SendBuffer(0) = 4 'Set Strobe

    'The next bytes are data
    SendBuffer(1) = Strobe
    
    NumberOfBytesWritten = 0

    Result = WriteFile(mHIDHandle, SendBuffer(0), OUT_REPORT_LENGTH, NumberOfBytesWritten, 0)
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    End If
  End If
End Property

Public Property Let RGB(Color As OLE_COLOR)
  Dim SendBuffer(0 To 7) As Byte
  Dim NumberOfBytesWritten As Long
  Dim Result As Long
  
  If mVersion = 1 Then Exit Property
    
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    'The first byte is the Report ID
    SendBuffer(0) = 3 'Set RGB

    'The next bytes are data
    SendBuffer(1) = CByte(Color And &HFF&)
    SendBuffer(2) = CByte((Color And &HFF00&) \ &H100)
    SendBuffer(3) = CByte((Color And &HFF0000) \ &H10000)
    
    NumberOfBytesWritten = 0

    Result = WriteFile(mHIDHandle, SendBuffer(0), OUT_REPORT_LENGTH, NumberOfBytesWritten, 0)
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    End If
  End If
End Property

Public Property Get RGB() As OLE_COLOR
  Dim SendBuffer(0 To 7) As Byte
  Dim Result As Long
    
  If mVersion = 1 Then RGB = vbWhite: Exit Property
    
  If mHIDHandle = INVALID_HANDLE_VALUE Then
    mHIDHandle = OpenHIDDevice
  End If
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    SendBuffer(0) = 7 'Get RGB
    Result = HidD_GetFeature(mHIDHandle, SendBuffer(0), CLng(4))
    If Result = 0 Then
      CloseHandle mHIDHandle
      mHIDHandle = INVALID_HANDLE_VALUE
    Else
      RGB = CLng(SendBuffer(1)) + 256 * (CLng(SendBuffer(2)) + 256 * CLng(SendBuffer(3)))
    End If
  End If
End Property

Private Sub Class_Initialize()
  mHIDHandle = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
  If mHIDHandle <> INVALID_HANDLE_VALUE Then
    CloseHandle mHIDHandle
  End If
End Sub


