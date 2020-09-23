VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPICalls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' collection of Calls

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit

Private mcCalls As Collection

Private mlDeviceIndex As Long               'parent index
Private mhOpenLine As Long                  'parent device handle



'parent device
Public Property Let DeviceIndex(ByVal newValue As Long)
    mlDeviceIndex = newValue
    
    Dim oCall As CTAPICall
    
    For Each oCall In mcCalls
        oCall.DeviceIndex = mlDeviceIndex
    Next

End Property
Public Property Get DeviceIndex() As Long
    DeviceIndex = mlDeviceIndex
End Property

Public Property Let OpenLineHandle(ByVal newValue As Long)
    mhOpenLine = newValue
    
    Dim oCall As CTAPICall
    
    For Each oCall In mcCalls
        oCall.OpenLineHandle = mhOpenLine
    Next

End Property
Public Property Get OpenLineHandle() As Long
    OpenLineHandle = mhOpenLine
End Property


Private Sub Class_Initialize()
    Set mcCalls = New Collection
End Sub
Private Sub Class_Terminate()
    DeleteAll
    Set mcCalls = Nothing
End Sub
Public Property Get Count() As Long
    Count = mcCalls.Count
End Property
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcCalls.[_NewEnum]
End Function

Public Function Item(ByVal lCallHandle As Long) As CTAPICall
Attribute Item.VB_UserMemId = 0

    On Error GoTo Item_Error
    
    Set Item = mcCalls.Item(CStr(lCallHandle))
    
Item_Exit:
    Exit Function
    
Item_Error:
    #If LevelA = 1 Then
        Stop
    #End If
    Set Item = Nothing
    Resume Item_Exit

End Function

Public Function AddGetCall(ByVal lCallHandle As Long) As CTAPICall

    On Error Resume Next
    
    Dim oCall As CTAPICall
    
    Set oCall = mcCalls.Item(CStr(lCallHandle))
    
    If oCall Is Nothing Then
        Set oCall = New CTAPICall
        oCall.OpenCallHandle = lCallHandle
        oCall.DeviceIndex = mlDeviceIndex
        oCall.OpenLineHandle = mhOpenLine
        mcCalls.Add oCall, CStr(lCallHandle)      '<-- treat OpenCallHandle as the key
    End If
    
    Set AddGetCall = oCall
    
    Set oCall = Nothing

End Function


'0-based index
Public Function ItemByIndex(ByVal lIndex As Long) As CTAPICall

    On Error GoTo Item_Error
    
    Set ItemByIndex = mcCalls.Item(lIndex + 1)
    
Item_Exit:
    Exit Function
    
Item_Error:
    #If LevelA = 1 Then
        Stop
    #End If
    Set ItemByIndex = Nothing
    Resume Item_Exit

End Function

Public Sub AddExisting(ByRef oCall As CTAPICall)

    If oCall.OpenCallHandle <> 0 Then

        oCall.DeviceIndex = mlDeviceIndex
        oCall.OpenLineHandle = mhOpenLine
        mcCalls.Add oCall, CStr(oCall.OpenCallHandle)      '<-- treat OpenCallHandle as the key
    
    End If
    
End Sub
'
Public Sub Delete(ByVal lCallHandle As Long)

    On Error Resume Next
    
    Dim oCall As CTAPICall
    
    Set oCall = mcCalls.Item(CStr(lCallHandle))
    
    If Not (oCall Is Nothing) Then
        oCall.Release
        mcCalls.Remove CStr(lCallHandle)   '<-- need to use lCallHandle because oCall.OpenCallHandle is now 0
    End If
    
End Sub

Public Sub DeleteAll()

    On Error Resume Next

    Dim oCall As CTAPICall
    Dim lHandle As Long
    
    For Each oCall In mcCalls
        lHandle = oCall.OpenCallHandle
        oCall.Release
        mcCalls.Remove CStr(lHandle)   '<-- need to use lHandle because oCall.OpenCallHandle is now 0
    Next
    
End Sub

