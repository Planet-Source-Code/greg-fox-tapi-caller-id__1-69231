VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPIDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' collection of Devices

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit

Private mcDevices As Collection


Private Sub Class_Initialize()
    Set mcDevices = New Collection
End Sub
Private Sub Class_Terminate()
    DeleteAll
    Set mcDevices = Nothing
End Sub
Public Property Get Count() As Long
    Count = mcDevices.Count
End Property
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcDevices.[_NewEnum]
End Function

'for when the Device is identified by its DeviceIndex (now is the same as the numerical index)
Public Function Item(ByVal lDeviceIndex As Long) As CTAPIDevice
Attribute Item.VB_UserMemId = 0

    On Error GoTo Item_Error
    
    Set Item = mcDevices.Item(CStr(lDeviceIndex))
    
Item_Exit:
    Exit Function
    
Item_Error:
    #If LevelA = 1 Then
        Stop
    #End If
    Set Item = Nothing
    Resume Item_Exit

End Function

'0-based index
Public Function ItemByIndex(ByVal lIndex As Long) As CTAPIDevice

    On Error GoTo Item_Error
    
    Set ItemByIndex = mcDevices.Item(lIndex + 1)
    
Item_Exit:
    Exit Function
    
Item_Error:
    #If LevelA = 1 Then
        Stop
    #End If
    Set ItemByIndex = Nothing
    Resume Item_Exit

End Function

Public Sub AddExisting(ByRef oDev As CTAPIDevice)

    If oDev.DeviceIndex < 0 Then oDev.DeviceIndex = mcDevices.Count     'protection

    mcDevices.Add oDev, CStr(oDev.DeviceIndex)      '<-- treat device index as the key
           
End Sub
'
Public Sub Delete(ByVal lDeviceIndex As Long)

    On Error Resume Next
    
    Dim oDev As CTAPIDevice
    
    Set oDev = mcDevices.Item(CStr(lDeviceIndex))
    
    If Not (oDev Is Nothing) Then
        oDev.Release
        mcDevices.Remove CStr(lDeviceIndex)
    End If
    
End Sub

Public Sub DeleteAll()

    On Error Resume Next

    Dim oDev As CTAPIDevice
    
    For Each oDev In mcDevices
        oDev.Release
        mcDevices.Remove CStr(oDev.DeviceIndex)
    Next
    
End Sub

