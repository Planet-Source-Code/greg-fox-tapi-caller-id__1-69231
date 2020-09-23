VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPICall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TAPI Call (managed by the Device Obj)

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit


Private mlDeviceIndex As Long               'parent index
Private mhOpenLine As Long                  'parent device handle

Private mhOpenCall As Long                  'call handle
Private moCalled As CTAPICallNode
Private moCaller As CTAPICallNode

Public Property Let OpenCallHandle(ByVal newValue As Long)
    mhOpenCall = newValue
End Property
Public Property Get OpenCallHandle() As Long
    OpenCallHandle = mhOpenCall
End Property


Public Property Let DeviceIndex(ByVal newValue As Long)
    mlDeviceIndex = newValue
End Property
Public Property Get DeviceIndex() As Long
    DeviceIndex = mlDeviceIndex
End Property

Public Property Let OpenLineHandle(ByVal newValue As Long)
    mhOpenLine = newValue
End Property
Public Property Get OpenLineHandle() As Long
    OpenLineHandle = mhOpenLine
End Property

Public Property Get Called() As CTAPICallNode
    Set Called = moCalled
End Property
Public Property Get Caller() As CTAPICallNode
    Set Caller = moCaller
End Property


Private Sub Class_Terminate()
    Release
End Sub

Public Sub Release()
    DeallocateTAPICall
    Set moCalled = Nothing
    Set moCaller = Nothing
End Sub

Public Sub DeallocateTAPICall()

    Dim rc As Long
    
    If mhOpenCall <> 0 Then
        rc = lineDeallocateCall(mhOpenCall)
        mhOpenCall = 0
    End If

End Sub

Private Sub Class_Initialize()
    Set moCalled = New CTAPICallNode
    Set moCaller = New CTAPICallNode
End Sub


'Because the LINE_CALLINFO msg is sent multiple times (even with the same changed-flag!), and we simply
'update all the interesting properties at once, we do a comparison and utilize "InfoChanged"
'flags to signal when a true "changed" event should be raised to the client.

Public Sub UpdateCallInfo()


    Dim uCallInfo As LINECALLINFO
    Dim rc As Long
    
    Dim sCallerID As String
    Dim sCallerName As String
    Dim sCallerStatus As String
    
    Dim sCalledID As String
    Dim sCalledName As String
    Dim sCalledStatus As String
    
    
    moCaller.InfoChanged = False
    moCalled.InfoChanged = False
    

    If mhOpenCall <> 0 Then
    
        uCallInfo.dwTotalSize = Len(uCallInfo)
        rc = lineGetCallInfo(mhOpenCall, uCallInfo)
    
    'Called
        If uCallInfo.dwCalledIDSize <> 0 Then _
            sCalledID = GetCallInfoString(uCallInfo.mem, uCallInfo.dwCalledIDOffset, uCallInfo.dwCalledIDSize)
        If uCallInfo.dwCalledIDNameSize <> 0 Then _
            sCalledName = GetCallInfoString(uCallInfo.mem, uCallInfo.dwCalledIDNameOffset, uCallInfo.dwCalledIDNameSize)
        sCalledStatus = GetCallNodeStatus(uCallInfo.dwCalledIDFlags)
    
        If StrComp(sCalledID, moCalled.ID, vbBinaryCompare) <> 0 Then
             moCalled.ID = sCalledID
             moCalled.InfoChanged = True
        End If
        If StrComp(sCalledName, moCalled.Name, vbBinaryCompare) <> 0 Then
             moCalled.Name = sCalledName
             moCalled.InfoChanged = True
        End If
        If StrComp(sCalledStatus, moCalled.Status, vbBinaryCompare) <> 0 Then
             moCalled.Status = sCalledStatus
             moCalled.InfoChanged = True
        End If
    
    
    'Caller
        If uCallInfo.dwCallerIDSize <> 0 Then _
            sCallerID = GetCallInfoString(uCallInfo.mem, uCallInfo.dwCallerIDOffset, uCallInfo.dwCallerIDSize)
        If uCallInfo.dwCallerIDNameSize <> 0 Then _
            sCallerName = GetCallInfoString(uCallInfo.mem, uCallInfo.dwCallerIDNameOffset, uCallInfo.dwCallerIDNameSize)
        sCallerStatus = GetCallNodeStatus(uCallInfo.dwCallerIDFlags)

        If StrComp(sCallerID, moCaller.ID, vbBinaryCompare) <> 0 Then
             moCaller.ID = sCallerID
             moCaller.InfoChanged = True
        End If
        If StrComp(sCallerName, moCaller.Name, vbBinaryCompare) <> 0 Then
             moCaller.Name = sCallerName
             moCaller.InfoChanged = True
        End If
        If StrComp(sCallerStatus, moCaller.Status, vbBinaryCompare) <> 0 Then
             moCaller.Status = sCallerStatus
             moCaller.InfoChanged = True
        End If

    End If

End Sub


