VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPIDeviceDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TAPI Device Description Object (for client use, through the DeviceList object)

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************


Option Explicit


Private mlDeviceIndex As Long

Private msAPIVersion As String

'devcaps variables
Private msLineName As String
Private msProviderInfo As String
Private msSwitchInfo As String
Private mlPermanentLineID As Long
Private mlStringFormat As Long
Private mlNumAddresses As Long
Private mlMaxDataRate As Long
Private mlBearerModes As Long
Private mlAddressModes As Long
Private mlMediaModes As Long
Private mlGenerateToneMaxNumFreq As Long
Private mlGenerateToneModes As Long
Private mlNumTerminals As Long

Private mlExtensionID0 As Long
Private mlExtensionID1 As Long
Private mlExtensionID2 As Long
Private mlExtensionID3 As Long



Public Property Get DeviceIndex() As Long
    DeviceIndex = mlDeviceIndex
End Property
Public Property Get DeviceAPIVersion() As String
    DeviceAPIVersion = msAPIVersion
End Property
Public Property Get LineName() As String
    LineName = msLineName
End Property
Public Property Get ProviderInfo() As String
    ProviderInfo = msProviderInfo
End Property
Public Property Get SwitchInfo() As String
    SwitchInfo = msSwitchInfo
End Property
Public Property Get PermanentLineID() As Long
    PermanentLineID = mlPermanentLineID
End Property
Public Property Get StringFormat() As Long
    StringFormat = mlStringFormat
End Property
Public Property Get NumAddresses() As Long
    NumAddresses = mlNumAddresses
End Property
Public Property Get MaxDataRate() As Long
    MaxDataRate = mlMaxDataRate
End Property
Public Property Get BearerModes() As Long
    BearerModes = mlBearerModes
End Property
Public Property Get AddressModes() As Long
    AddressModes = mlAddressModes
End Property
Public Property Get MediaModes() As Long
    MediaModes = mlMediaModes
End Property
Public Property Get GenerateToneMaxNumFreq() As Long
    GenerateToneMaxNumFreq = mlGenerateToneMaxNumFreq
End Property
Public Property Get GenerateToneModes() As Long
    GenerateToneModes = mlGenerateToneModes
End Property
Public Property Get NumTerminals() As Long
    NumTerminals = mlNumTerminals
End Property

Public Property Get DeviceSupportsVoiceCalls() As Boolean
    If mlBearerModes And LINEBEARERMODE_VOICE Then
        If mlMediaModes And LINEMEDIAMODE_INTERACTIVEVOICE Then
            DeviceSupportsVoiceCalls = True
        End If
    End If
End Property

Friend Sub CopyDeviceInfo(ByVal oDev As CTAPIDevice)

    On Error Resume Next
    
    With oDev
        mlDeviceIndex = .DeviceIndex
        msAPIVersion = MakeVerStr(.DeviceAPIVersion)
        msLineName = .LineName
        msProviderInfo = .ProviderInfo
        msSwitchInfo = .SwitchInfo
        mlPermanentLineID = .PermanentLineID
        mlStringFormat = .StringFormat
        mlNumAddresses = .NumAddresses
        mlMaxDataRate = .MaxDataRate
        mlBearerModes = .BearerModes
        mlAddressModes = .AddressModes
        mlMediaModes = .MediaModes
        mlGenerateToneMaxNumFreq = .GenerateToneMaxNumFreq
        mlGenerateToneModes = .GenerateToneModes
        mlNumTerminals = .NumTerminals
        
        mlExtensionID0 = .ExtensionID0
        mlExtensionID1 = .ExtensionID1
        mlExtensionID2 = .ExtensionID2
        mlExtensionID3 = .ExtensionID3
        
    End With

End Sub


Public Sub PaintDeviceIcon(ByVal hdc As Long, _
                                Optional ByVal lLeft As Long = 0, _
                                Optional ByVal lTop As Long = 0, _
                                Optional ByVal lWidth As Long = 0, _
                                Optional ByVal lHeight As Long = 0)
    On Error Resume Next
    
    If Not goTAPIApp Is Nothing Then _
        goTAPIApp.TAPIDevice(mlDeviceIndex).PaintDeviceIcon hdc, lLeft, lTop, lWidth, lHeight
     
End Sub

Public Function OpenConfigDialog(Optional ByVal hOwnerHwnd As Long = 0&, _
    Optional ByVal sDeviceClass As String = "") As Boolean
    
    On Error Resume Next
    
    If Not goTAPIApp Is Nothing Then _
        OpenConfigDialog = goTAPIApp.TAPIDevice(mlDeviceIndex).OpenConfigDialog(hOwnerHwnd, sDeviceClass)
        
End Function

'this is called from the App class because this fn needs these first 2 parms...
'
Public Function OpenDialingPropDialog(ByVal hOwnerHwnd As Long, _
    ByVal sPhoneNumber As String) As Boolean
    
    On Error Resume Next
    
    If Not goTAPIApp Is Nothing Then _
        OpenDialingPropDialog = goTAPIApp.OpenDialingPropDialog(mlDeviceIndex, hOwnerHwnd, sPhoneNumber)
   
End Function


