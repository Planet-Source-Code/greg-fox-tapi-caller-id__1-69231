Attribute VB_Name = "mTAPIMain"
' TAPI Callback and Sub Main()

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit

Global goTAPIApp As CTAPIApp            'Singleton - Main TAPI interface object


Public Sub LineCallbackProc(ByVal hDevice As Long, _
                                ByVal dwMsg As Long, _
                                ByVal dwCallbackInstance As Long, _
                                ByVal dwParam1 As Long, _
                                ByVal dwParam2 As Long, _
                                ByVal dwParam3 As Long)
                                
    On Error Resume Next
    
    Dim oDevice As CTAPIDevice
    Dim objTemp As CTAPIDevice
    
    'Debug.Print "LineCallbackProc - Msg:" & Hex$(dwMsg)


    'we have told tapi to use the ObjPtr() of the CTAPIDevice instance in dwCallbackInstance,
    'so I pass along a reference to a CTAPIDevice object for convenience
    
    If dwCallbackInstance = 0 Then
        Set oDevice = New CTAPIDevice               'So we don't need to test for Is Nothing in case of error
    Else
        
        CopyMemory objTemp, dwCallbackInstance, 4   'Turn pointer into illegal, uncounted reference
        Set oDevice = objTemp                       'Assign to legal reference
        CopyMemory objTemp, 0&, 4                   'Destroy the illegal reference
        
    End If

    goTAPIApp.LineProcHandler hDevice, dwMsg, oDevice, dwParam1, dwParam2, dwParam3

End Sub


Public Function GetTAPIStructString(ByVal ptrTapistruct As Long, ByVal offset As Long, ByVal length As Long) As String
    
    Dim buffer() As Byte
    
    If length >= 0 Then       'ck for erroneous input
    
        If offset Then
            ReDim buffer(0 To length - 1)
            CopyMemory buffer(0), ByVal ptrTapistruct + offset, length
            GetTAPIStructString = StrConv(buffer, vbUnicode)
        End If
    End If
    
End Function


Public Function GetCallInfoString(ByRef mem As String, ByVal offset As Long, ByVal size As Long) As String
    GetCallInfoString = Trim$(Replace(Replace(Mid(mem, offset + 1 - LINECALLINFO_FIXEDSIZE, size - 1), Chr(0), " "), "|", " "))
End Function


Public Function GetCallNodeStatus(ByVal lStatus As Long) As String
    Select Case lStatus
        Case LINECALLPARTYID_BLOCKED: GetCallNodeStatus = "Blocked"
        Case LINECALLPARTYID_OUTOFAREA: GetCallNodeStatus = "OutOfArea"
        Case LINECALLPARTYID_NAME: GetCallNodeStatus = "Name"
        Case LINECALLPARTYID_ADDRESS: GetCallNodeStatus = "Address"
        Case LINECALLPARTYID_PARTIAL: GetCallNodeStatus = "Partial"
        Case LINECALLPARTYID_UNAVAIL: GetCallNodeStatus = "Unavail"
        Case Else: GetCallNodeStatus = "Unknown"
    End Select
End Function

Public Function MakeVerStr(ByVal lVersion As Long) As String
    Dim v As String
    Dim s As String
    Dim l As Long
    v = Hex$(lVersion)
    l = Len(v)
    If l > 4 Then
        s = CStr(Val(left$(v, l - 4))) & "."
    End If
    If l > 2 Then
        s = s & CStr(Val(Mid$(v, l - 3, 2))) & "."
    End If
    If l > 0 Then
        s = s & CStr(Val(Right$(v, 2)))
    End If
    MakeVerStr = s
End Function

Public Sub Main()

    Dim bSuccess As Boolean

    Set goTAPIApp = New CTAPIApp        'Initialize the TAPI class (use default v1.3 - 3.0)
    
    bSuccess = goTAPIApp.Initialize()
    
    If Not bSuccess Then
        
        're-negotiate lower version
        goTAPIApp.LowAPI = &H10003 ' 1.3 = &H00010003
        goTAPIApp.HiAPI = &H10004 '  1.4 = &H00010004
        bSuccess = goTAPIApp.Initialize()
    
    End If
      
End Sub
