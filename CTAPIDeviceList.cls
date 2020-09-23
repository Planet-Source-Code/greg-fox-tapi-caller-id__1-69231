VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPIDeviceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TAPI Device List Interface (Public)

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************


Option Explicit


Private mcDevs As Collection
Private msErrorString As String
Private msTAPIVersion As String

Private Sub Class_Initialize()

    On Error Resume Next
    
    Dim oDevDesc As CTAPIDeviceDesc
    Dim lCount As Long
    Dim i As Long

    Set mcDevs = New Collection

    If goTAPIApp Is Nothing Then
        msErrorString = "Unable to initialize the TAPI system"
    
    Else
        lCount = goTAPIApp.TAPIDeviceCount
        If lCount = 0& Then
            msErrorString = "No TAPI devices defined"
        Else
            For i = 0 To lCount - 1
                Set oDevDesc = New CTAPIDeviceDesc
                oDevDesc.CopyDeviceInfo goTAPIApp.TAPIDeviceByIndex(i)
                mcDevs.Add oDevDesc, CStr(oDevDesc.DeviceIndex)
            Next
        End If
        
        msTAPIVersion = MakeVerStr(goTAPIApp.APIVersion)
    End If

End Sub

Public Property Get ErrorString() As String
    ErrorString = msErrorString
End Property

Public Property Get TAPIVersion() As String
    TAPIVersion = msTAPIVersion
End Property

Public Property Get TAPIDeviceCount() As Long
    TAPIDeviceCount = mcDevs.Count
End Property

'this uses numerical index (0-based) - it is the same as the DeviceIndex
Public Property Get TAPIDevice(ByVal lDeviceIndex As Long) As CTAPIDeviceDesc
    Set TAPIDevice = mcDevs.Item(lDeviceIndex + 1)
End Property




