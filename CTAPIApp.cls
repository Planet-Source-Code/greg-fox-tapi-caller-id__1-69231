VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPIApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TAPI App (Main Class) - Singleton - Globally referenced as goTAPIApp by Public client-facing classes

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

'Notes:

'1.  When I phone, again, before the LINECALLSTATE_IDLE msg, TAPI uses the SAME call number for the
'    second call (100AD) !

'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100E1, Msg: $17, P1: $0, P2: $100AD, P3: $4 - LINE_APPNEWCALL
'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100AD, Msg: $2, P1: $2, P2: $0, P3: $0 - LINECALLSTATE_OFFERING
'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100AD, Msg: $1, P1: $8000, P2: $0, P3: $0 - LINE_CALLINFO
'  >> Caller ID: 8561224444, Name: FOX GREG, Status: Unknown
'  >> Called ID: , Name: , Status: Unavail
'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100AD, Msg: $1, P1: $8000, P2: $0, P3: $0 - LINE_CALLINFO
'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100AD, Msg: $1, P1: $8000, P2: $0, P3: $0 - LINE_CALLINFO
'  >> Caller ID: 8562223333, Name: DATA MGMT ASSOC, Status: Unknown
'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100AD, Msg: $1, P1: $8000, P2: $0, P3: $0 - LINE_CALLINFO
'CTAPIApp.LineProcHandler(0)(100E1): Dev: $100AD, Msg: $2, P1: $1, P2: $0, P3: $0 - LINECALLSTATE_IDLE




Option Explicit


Private mhLineApp As Long               'handle to TAPI
Private mlApplicationVersion As Long    'some TAPI calls need this instead of negotiated API version
Private mlLowAPI As Long                'default TAPI 1.3 (&H00010003)
Private mlHighAPI As Long               'default TAPI 3.0 (&H00030000)
Private mlLastTAPIError As Long
Private mbInitialized As Boolean            'result of .Create()

Private mcDevices As CTAPIDevices


'Events
Public Event IncomingCIDMsg( _
    ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String, _
    ByVal sCallerID As String, _
    ByVal sCallerName As String)
    
Public Event CallEnded( _
    ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String)


Private Sub Class_Initialize()

    #If LevelA = 1 Then                     '<-- This is our convention for specifying component code-base levels
        Debug.Print "CTAPIApp.Class_Init"   '    ie: LevelA is when running in the IDE
    #End If
    
    mlLowAPI = &H10003           'default TAPI 1.3 - 3.0
    mlHighAPI = &H30000
    
    Set mcDevices = New CTAPIDevices
    
End Sub

Private Sub Class_Terminate()
    Release
    #If LevelA = 1 Then
        Debug.Print "CTAPIApp.Class_Terminate"
    #End If
End Sub

Public Sub Release()
    
    On Error Resume Next
    
    If Not (mcDevices Is Nothing) Then
        mcDevices.DeleteAll
        Set mcDevices = Nothing
    End If
    
    If mhLineApp <> 0 Then
        lineShutdown mhLineApp
        mhLineApp = 0
    End If
    
End Sub


Public Property Get LowAPI() As Long
    LowAPI = mlLowAPI
End Property
Public Property Let LowAPI(ByVal ver As Long)
    mlLowAPI = ver
End Property

Public Property Get HiAPI() As Long
    HiAPI = mlHighAPI
End Property
Public Property Let HiAPI(ByVal ver As Long)
    mlHighAPI = ver
End Property


Public Property Get LastError() As Long
    LastError = mlLastTAPIError
    mlLastTAPIError = 0&
End Property

Public Function ErrorString(ByVal lErrCode As Long) As String
    ErrorString = GetLineErrString(lErrCode)
End Function

Public Property Get Initialized() As Boolean        'the result of Initialize()
    Initialized = mbInitialized
End Property

Public Function Initialize() As Boolean

    Dim oDev As CTAPIDevice
    Dim rc As Long
    Dim lLineIndex As Long
    Dim lLineCount As Long
    
    
    Dim lip As LINEINITIALIZEEXPARAMS
    Dim lxid As LINEEXTENSIONID
    
    mbInitialized = False
    
    'init params
    lip.dwTotalSize = Len(lip)
    lip.dwOptions = LINEINITIALIZEEXOPTION_USEHIDDENWINDOW
    
    'initialize TAPI lines and get handle (mhLineApp)
    rc = lineInitializeEx(mhLineApp, App.hInstance, AddressOf mTAPIMain.LineCallbackProc, App.Title & vbNullString, _
                                 lLineCount, mlHighAPI, lip)
    If rc <> TAPI_SUCCESS Then
        mlLastTAPIError = rc
    Else
        If lLineCount > 0 Then
            For lLineIndex = 0 To lLineCount - 1
                Set oDev = New CTAPIDevice
                oDev.DeviceIndex = lLineIndex
                oDev.InitializeLine mhLineApp, mlLowAPI, mlHighAPI, Me
                mcDevices.AddExisting oDev
            Next
        End If
        Initialize = True
        mbInitialized = True
    End If
    
    #If LevelA = 1 Then
        Debug.Print "CTAPIApp.Create Result: " & IIf(mbInitialized, "True", "False")
    #End If

End Function


'for call back from each Device - so we can get the highest API version
'lDeviceIndex passed for logging
Friend Sub DeviceAPIVerNegotiated(ByVal lDeviceIndex As Long, ByVal lAPIVersion As Long)
    If (lAPIVersion > mlApplicationVersion) Then mlApplicationVersion = lAPIVersion
End Sub

Public Property Get APIVersion() As Long
    APIVersion = mlApplicationVersion
End Property


'uses string value of the index as the Key
Public Function TAPIDevice(ByVal lDeviceIndex As Long) As CTAPIDevice
    Set TAPIDevice = mcDevices.Item(lDeviceIndex)
End Function
Public Property Get TAPIDeviceCount() As Long
    TAPIDeviceCount = mcDevices.Count
End Property
Public Function TAPIDeviceByIndex(ByVal lIndex As Long) As CTAPIDevice
    Set TAPIDeviceByIndex = mcDevices.ItemByIndex(lIndex)
End Function

Public Function TAPIDeviceNegotiatedAPI(ByVal lDeviceIndex As Long) As Long

    On Error Resume Next

    TAPIDeviceNegotiatedAPI = mcDevices.Item(lDeviceIndex).DeviceAPIVersion
    
End Function

Public Function TAPIDeviceName(ByVal lDeviceIndex As Long) As String

    On Error Resume Next

    TAPIDeviceName = mcDevices.Item(lDeviceIndex).LineName
    
End Function

Public Function OpenDialingPropDialog(ByVal lDeviceIndex As Long, _
    Optional ByVal hOwnerHwnd As Long = 0&, _
    Optional ByVal sPhoneNumber As String = "") As Boolean

    On Error Resume Next
    
    OpenDialingPropDialog = mcDevices.Item(lDeviceIndex).OpenDialingPropDialog( _
        mhLineApp, mlApplicationVersion, hOwnerHwnd, sPhoneNumber)

End Function

Public Function OpenLine(ByVal lDeviceIndex As Long, _
    Optional ByVal lPrivilege As EnumTAPICallPrivileges = LINECALLPRIVILEGE_NONE, _
    Optional ByVal lMediaModes As EnumTAPIMediaModes = LINEMEDIAMODE_INTERACTIVEVOICE) As Boolean
    
    
    On Error Resume Next

    OpenLine = mcDevices.Item(lDeviceIndex).OpenLine(mhLineApp, lPrivilege, lMediaModes)

End Function

Public Sub CloseLine(ByVal lDeviceIndex As Long)

    On Error Resume Next

    mcDevices.Item(lDeviceIndex).CloseLine
    
End Sub

Public Function RegisterRequestRecipient(ByVal lDeviceIndex As Long, _
    ByVal lRequestMode As EnumTAPIRequestModes) As Boolean
    
    On Error Resume Next

    RegisterRequestRecipient = mcDevices.Item(lDeviceIndex).RegisterRequestRecipient(mhLineApp, lRequestMode)

End Function

Public Function UnregisterRequestRecipient(ByVal lDeviceIndex As Long, _
    ByVal lRequestMode As EnumTAPIRequestModes) As Boolean
    
    On Error Resume Next

    UnregisterRequestRecipient = mcDevices.Item(lDeviceIndex).UnregisterRequestRecipient(mhLineApp, lRequestMode)

End Function

'sets to True - use ClearPriority to set to False
Public Function SetPriority(ByVal lDeviceIndex As Long, _
    ByVal lPriorityMode As EnumTAPIMediaModes, _
    ByVal lRequestMode As EnumTAPIRequestModes) As Boolean

    On Error Resume Next

    SetPriority = mcDevices.Item(lDeviceIndex).SetPriority(mhLineApp, lPriorityMode, lRequestMode, True)

End Function

Public Function ClearPriority(ByVal lDeviceIndex As Long, _
    ByVal lPriorityMode As EnumTAPIMediaModes, _
    ByVal lRequestMode As EnumTAPIRequestModes) As Boolean

    On Error Resume Next

    ClearPriority = mcDevices.Item(lDeviceIndex).SetPriority(mhLineApp, lPriorityMode, lRequestMode, False)

End Function

Public Function GetPriority(ByVal lDeviceIndex As Long, _
    ByVal lPriorityMode As EnumTAPIMediaModes, _
    ByVal lRequestMode As EnumTAPIRequestModes, _
    ByRef bPriority As Boolean) As Boolean

    On Error Resume Next

    GetPriority = mcDevices.Item(lDeviceIndex).GetPriority(mhLineApp, lPriorityMode, lRequestMode, bPriority)

End Function


'------------------------------------------------------------

Friend Sub LineProcHandler(ByVal hDevice As Long, _
                                        ByVal dwMsg As Long, _
                                        ByVal oDevice As CTAPIDevice, _
                                        ByVal dwParam1 As Long, _
                                        ByVal dwParam2 As Long, _
                                        ByVal dwParam3 As Long)
                                        
    Dim sDesc As String
    Dim uCInfoState As EnumTAPICallInfoStates
    
    sDesc = "CTAPIApp.LineProcHandler(" & CStr(oDevice.DeviceIndex) & ")(" & _
        Hex$(oDevice.OpenLineHandle) & "): Dev: $" & Hex$(hDevice) & ", Msg: $" & Hex$(dwMsg) & _
        ", P1: $" & Hex$(dwParam1) & ", P2: $" & Hex$(dwParam2) & ", P3: $" & Hex$(dwParam3) & " - "
        
    
    Select Case dwMsg
        Case LINE_REPLY
            Debug.Print sDesc & "LINE_REPLY"
            
        Case LINE_APPNEWCALL:
            Debug.Print sDesc & "LINE_APPNEWCALL"
            oDevice.Calls.AddGetCall dwParam2

        Case LINE_CALLINFO:
            Debug.Print sDesc & "LINE_CALLINFO"
            uCInfoState = dwParam1
            
                                    'OR together the ones we are interested in...
            If (uCInfoState And (LINECALLINFOSTATE_CALLERID Or LINECALLINFOSTATE_CALLEDID)) <> 0 Then
                oDevice.Calls.AddGetCall(hDevice).UpdateCallInfo
                
                With oDevice.Calls.Item(hDevice).Caller
                    If .InfoChanged Then
                        #If LevelA = 1 Then
                            Debug.Print "  >> Caller ID: " & .ID & ", Name: " & .Name & ", Status: " & .Status
                        #End If
                        RaiseEvent IncomingCIDMsg(oDevice.DeviceIndex, oDevice.LineName, .ID, .Name)
                    End If
                End With
                        
            End If

        Case LINE_CALLSTATE
            Select Case dwParam1
                Case LINECALLSTATE_OFFERING
                    Debug.Print sDesc & "LINECALLSTATE_OFFERING"
                    oDevice.Calls.AddGetCall hDevice
                    
                Case LINECALLSTATE_DISCONNECTED
                    'remote party has disconnected from the call
                    Debug.Print sDesc & "LINECALLSTATE_DISCONNECTED"
                                    
                Case LINECALLSTATE_IDLE
                    'no call exists - the line is idle
                    Debug.Print sDesc & "LINECALLSTATE_IDLE"
                    
                    On Error Resume Next
                    oDevice.Calls.Delete hDevice        'also calls lineDeallocate
                    
                    RaiseEvent CallEnded(oDevice.DeviceIndex, oDevice.LineName)
                    
                 
                Case LINECALLSTATE_CONNECTED
                    Debug.Print sDesc & "LINECALLSTATE_CONNECTED"
                                       
                Case LINECALLSTATE_BUSY
                    Debug.Print sDesc & "LINECALLSTATE_BUSY"
                
                Case LINECALLSTATE_DIALTONE
                    'switch is ready to receive a dialed number
                    Debug.Print sDesc & "LINECALLSTATE_DIALTONE"
                
                Case LINECALLSTATE_RINGBACK
                    'the other station has been reached and is being alerted (ringing)
                    Debug.Print sDesc & "LINECALLSTATE_RINGBACK"
                
                Case LINECALLSTATE_DIALING
                    Debug.Print sDesc & "LINECALLSTATE_DIALING"
                
                Case LINECALLSTATE_PROCEEDING
                    Debug.Print sDesc & "LINECALLSTATE_PROCEEDING"
                
                Case LINECALLSTATE_SPECIALINFO
                    'network error occured
                    Debug.Print sDesc & "LINECALLSTATE_SPECIALINFO"
                    
                Case Else
                    Debug.Print sDesc & "Unrecognized CallSTATE (dwParam1): $" & Hex(dwParam1)
            End Select
            
        Case LINE_ADDRESSSTATE: Debug.Print sDesc & "LINE_ADDRESSSTATE"


        Case LINE_CLOSE: Debug.Print sDesc & "LINE_CLOSE"
        Case LINE_CREATE: Debug.Print sDesc & "LINE_CREATE"
        Case LINE_AGENTSPECIFIC: Debug.Print sDesc & "LINE_AGENTSPECIFIC"
        Case LINE_AGENTSTATUS: Debug.Print sDesc & "LINE_AGENTSTATUS"
        Case LINE_PROXYREQUEST: Debug.Print sDesc & "LINE_PROXYREQUEST"
        Case LINE_REMOVE: Debug.Print sDesc & "LINE_REMOVE"
        Case LINE_DEVSPECIFIC: Debug.Print sDesc & "LINE_DEVSPECIFIC"
        Case LINE_LINEDEVSTATE: Debug.Print sDesc & "LINE_LINEDEVSTATE"
        Case LINE_REQUEST: Debug.Print sDesc & "LINE_REQUEST"

        Case Else: Debug.Print sDesc & "Unrecognized dwMsg: $" & Hex(dwMsg)
        
    End Select

End Sub

