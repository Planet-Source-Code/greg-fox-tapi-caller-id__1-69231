VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPICID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TAPI CallerID Interface

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************


Option Explicit


Public Event IncomingCIDMsg( _
    ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String, _
    ByVal sCallerID As String, _
    ByVal sCallerName As String)
    
Public Event CallEnded( _
    ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String)
    
Private WithEvents moTAPIApp As CTAPIApp
Attribute moTAPIApp.VB_VarHelpID = -1
Private mlDeviceIndex As Long               '-1 if not open

Private msErrorString As String
Private msTAPIVersion As String

    
    
Private Sub Class_Initialize()
       
    On Error Resume Next
    
    mlDeviceIndex = -1&
    
    If goTAPIApp Is Nothing Then
        msErrorString = "Unable to initialize the TAPI system"
    Else
        Set moTAPIApp = goTAPIApp
        msTAPIVersion = MakeVerStr(moTAPIApp.APIVersion)
    End If
       
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    CloseCIDLine
    Set moTAPIApp = Nothing
End Sub

Public Property Get ErrorString() As String
    ErrorString = msErrorString
End Property

Public Property Get TAPIVersion() As String
    TAPIVersion = msTAPIVersion
End Property



Public Sub OpenCIDLine(ByVal lDeviceIndex As Long)

    On Error Resume Next
    
    mlDeviceIndex = lDeviceIndex
    
    moTAPIApp.OpenLine mlDeviceIndex, LINECALLPRIVILEGE_MONITOR Or LINECALLPRIVILEGE_OWNER, _
        LINEMEDIAMODE_DATAMODEM
    
    moTAPIApp.RegisterRequestRecipient mlDeviceIndex, LINEREQUESTMODE_MAKECALL
    
    
    moTAPIApp.SetPriority mlDeviceIndex, LINEMEDIAMODE_DATAMODEM, LINEREQUESTMODE_MAKECALL
    moTAPIApp.SetPriority mlDeviceIndex, LINEMEDIAMODE_INTERACTIVEVOICE, LINEREQUESTMODE_MAKECALL

End Sub

Public Sub CloseCIDLine()

    On Error Resume Next

    If mlDeviceIndex >= 0& Then

        moTAPIApp.ClearPriority mlDeviceIndex, LINEMEDIAMODE_DATAMODEM, LINEREQUESTMODE_MAKECALL
        moTAPIApp.ClearPriority mlDeviceIndex, LINEMEDIAMODE_INTERACTIVEVOICE, LINEREQUESTMODE_MAKECALL
    
        moTAPIApp.UnregisterRequestRecipient mlDeviceIndex, LINEREQUESTMODE_MAKECALL
    
        moTAPIApp.CloseLine mlDeviceIndex
        
        mlDeviceIndex = -1&
        
    End If

End Sub

Private Sub moTAPIApp_CallEnded(ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String)

    RaiseEvent CallEnded(lTAPIDeviceIndex, sTAPIDeviceName)

End Sub

Private Sub moTAPIApp_IncomingCIDMsg(ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String, _
    ByVal sCallerID As String, _
    ByVal sCallerName As String)
    
    RaiseEvent IncomingCIDMsg(lTAPIDeviceIndex, sTAPIDeviceName, sCallerID, sCallerName)

End Sub
