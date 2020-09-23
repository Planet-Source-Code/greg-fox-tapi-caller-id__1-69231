VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPIDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TAPI Device

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit


Private mlDeviceIndex As Long

Private mlAPIVersion As Long
Private mlLowAPI As Long
Private mlHighAPI As Long
Private mlLastTAPIError As Long
Private mhOpenLine As Long

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

'ext id vars
Private mlExtensionID0 As Long
Private mlExtensionID1 As Long
Private mlExtensionID2 As Long
Private mlExtensionID3 As Long

Private mcCalls As CTAPICalls



Private Sub Class_Initialize()
    mlDeviceIndex = -1          'ie, not set
    Set mcCalls = New CTAPICalls
End Sub

Public Property Let DeviceIndex(ByVal newValue As Long)
    mlDeviceIndex = newValue
End Property
Public Property Get DeviceIndex() As Long
    DeviceIndex = mlDeviceIndex
End Property

Public Property Get DeviceAPIVersion() As Long
    DeviceAPIVersion = mlAPIVersion
End Property

Public Property Let OpenLineHandle(ByVal newValue As Long)
    mhOpenLine = newValue
End Property
Public Property Get OpenLineHandle() As Long
    OpenLineHandle = mhOpenLine
End Property

Public Property Get Calls() As CTAPICalls
    Set Calls = mcCalls
End Property


'Dev Caps
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




'ext id
Public Property Get ExtensionID0() As Long
    ExtensionID0 = mlExtensionID0
End Property
Public Property Get ExtensionID1() As Long
    ExtensionID1 = mlExtensionID1
End Property
Public Property Get ExtensionID2() As Long
    ExtensionID2 = mlExtensionID2
End Property
Public Property Get ExtensionID3() As Long
    ExtensionID3 = mlExtensionID3
End Property


Public Function InitializeLine(ByVal hLineApp As Long, _
    ByVal lLowAPI As Long, _
    ByVal lHighAPI As Long, _
    ByRef oCallBack As CTAPIApp) As Boolean

    On Error GoTo InitLine_Error
    
    Dim bResult As Boolean
    
    mlLowAPI = lLowAPI: mlHighAPI = lHighAPI
    
    NegotiateAPIVersion hLineApp
    
    oCallBack.DeviceAPIVerNegotiated mlDeviceIndex, mlAPIVersion
    
    bResult = GetLineDevCaps(hLineApp)
    
    mcCalls.DeviceIndex = mlDeviceIndex
    mcCalls.OpenLineHandle = mhOpenLine
    
    #If LevelA = 1 Then
        Debug.Print "CTAPIDevice(" & CStr(mlDeviceIndex) & ").InitializeLine: " & msLineName
    #End If

    
InitLine_Exit:
    InitializeLine = bResult
    Exit Function
    
InitLine_Error:
    Resume InitLine_Exit
    
End Function

Private Sub NegotiateAPIVersion(ByVal hLineApp As Long)

    Dim rc As Long
    Dim uLxID As LINEEXTENSIONID
    
    rc = lineNegotiateAPIVersion(hLineApp, mlDeviceIndex, mlLowAPI, mlHighAPI, mlAPIVersion, uLxID)
    
    If rc <> TAPI_SUCCESS Then
        mlAPIVersion = 0&
    Else
        'also cache extension version (dev-specific features) if available
        mlExtensionID0 = uLxID.dwExtensionID0
        mlExtensionID1 = uLxID.dwExtensionID1
        mlExtensionID2 = uLxID.dwExtensionID2
        mlExtensionID3 = uLxID.dwExtensionID3
            'TODO! add call to lineNegotiateExtVersion()
    
    End If

End Sub

Private Function GetLineDevCaps(ByVal hLineApp As Long) As Boolean

    Dim rc As Long
    Dim uLineCaps As LINEDEVCAPS
    
    uLineCaps.dwTotalSize = Len(uLineCaps)      'init required fields in devcaps UDT
    
    rc = lineGetDevCaps(hLineApp, mlDeviceIndex, mlAPIVersion, 0&, uLineCaps)
    
    If rc <> TAPI_SUCCESS Then
        mlLastTAPIError = rc
    Else
        msLineName = GetTAPIStructString(VarPtr(uLineCaps), uLineCaps.dwLineNameOffset, uLineCaps.dwLineNameSize)
        msProviderInfo = GetTAPIStructString(VarPtr(uLineCaps), uLineCaps.dwProviderInfoOffset, uLineCaps.dwProviderInfoSize)
        msSwitchInfo = GetTAPIStructString(VarPtr(uLineCaps), uLineCaps.dwSwitchInfoOffset, uLineCaps.dwSwitchInfoSize)
        mlPermanentLineID = uLineCaps.dwPermanentLineID
        mlStringFormat = uLineCaps.dwStringFormat
        mlNumAddresses = uLineCaps.dwNumAddresses
        mlMaxDataRate = uLineCaps.dwMaxRate
        mlBearerModes = uLineCaps.dwBearerModes
        mlAddressModes = uLineCaps.dwAddressModes
        mlMediaModes = uLineCaps.dwMediaModes
        mlGenerateToneMaxNumFreq = uLineCaps.dwGenerateToneMaxNumFreq
        mlGenerateToneModes = uLineCaps.dwGenerateToneModes
        mlNumTerminals = uLineCaps.dwNumTerminals
        
        GetLineDevCaps = True
    End If
    
End Function


Public Function PaintDeviceIcon(ByVal hdc As Long, _
                                Optional ByVal lLeft As Long = 0, _
                                Optional ByVal lTop As Long = 0, _
                                Optional ByVal lWidth As Long = 0, _
                                Optional ByVal lHeight As Long = 0) As Boolean
    Dim hIcon As Long
    Dim rc As Long
    
    rc = lineGetIcon(mlDeviceIndex, 0&, hIcon)
    If rc <> TAPI_SUCCESS Then
        mlLastTAPIError = rc
    Else
        rc = DrawIconEx(hdc, lLeft, lTop, hIcon, lWidth, lHeight, 0&, 0&, DI_NORMAL)
                            '(function succeeds even when there is no icon in the TSP)
        If hIcon <> 0 Then   'function failed (check actual hIcon instead of return value)
            DestroyIcon hIcon
            PaintDeviceIcon = True
        End If
    End If
    
End Function


Public Sub Release()
    
    On Error Resume Next
    
    If Not (mcCalls Is Nothing) Then
        mcCalls.DeleteAll
        Set mcCalls = Nothing
    End If
    
    CloseLine
    
End Sub

Public Sub CloseLine()
    
    On Error Resume Next
    
    If mhOpenLine <> 0& Then
        lineClose mhOpenLine
        mhOpenLine = 0&
        mcCalls.OpenLineHandle = 0&
        
        #If LevelA = 1 Then
            Debug.Print "CTAPIDevice(" & CStr(mlDeviceIndex) & ").CloseLine"
        #End If
    End If
    
End Sub

Friend Function OpenLine(ByVal hLineApp As Long, _
    ByVal lPrivilege As Long, _
    ByVal lMediaModes As Long) As Boolean
    
    Dim uLineCallParam As LINECALLPARAMS
    Dim rc As Long
    
    If mhOpenLine = 0& Then
    
        uLineCallParam.dwTotalSize = Len(uLineCallParam)
        
        rc = lineOpen(hLineApp, _
                        mlDeviceIndex, _
                        mhOpenLine, _
                        mlAPIVersion, _
                        0&, _
                        ObjPtr(Me), _
                        lPrivilege, _
                        lMediaModes, _
                        ByVal 0&)
                        
        If rc <> TAPI_SUCCESS Then
            mlLastTAPIError = rc
        Else
            OpenLine = True
        End If
        
        #If LevelA = 1 Then
            Debug.Print "CTAPIDevice(" & CStr(mlDeviceIndex) & ").OpenLine Result: " & IIf(rc = TAPI_SUCCESS, "True", "False")
        #End If

    End If
    
End Function


Friend Function RegisterRequestRecipient(ByVal hLineApp As Long, _
    ByVal lRequestMode As Long) As Boolean
    
    Dim uLineCallParam As LINECALLPARAMS
    Dim rc As Long
    
    If mhOpenLine <> 0& Then
    
        rc = lineRegisterRequestRecipient(hLineApp, _
            App.hInstance, _
            lRequestMode, _
            True)
                        
        If rc <> TAPI_SUCCESS Then
            mlLastTAPIError = rc
        Else
            RegisterRequestRecipient = True
        End If
    
        #If LevelA = 1 Then
            Debug.Print "CTAPIDevice(" & CStr(mlDeviceIndex) & ").RegisterRequestRecipient Result: " & IIf(rc = TAPI_SUCCESS, "True", "False") & ", Mode: " & Hex$(lRequestMode)
        #End If
        
    End If
    
End Function

Friend Function UnregisterRequestRecipient(ByVal hLineApp As Long, _
    ByVal lRequestMode As Long) As Boolean
    
    Dim uLineCallParam As LINECALLPARAMS
    Dim rc As Long
    
    If mhOpenLine <> 0& Then
    
        rc = lineRegisterRequestRecipient(hLineApp, _
            App.hInstance, _
            lRequestMode, _
            False)
                        
        If rc <> TAPI_SUCCESS Then
            mlLastTAPIError = rc
        Else
            UnregisterRequestRecipient = True
        End If
        
        #If LevelA = 1 Then
            Debug.Print "CTAPIDevice(" & CStr(mlDeviceIndex) & ").UnregisterRequestRecipient Result: " & IIf(rc = TAPI_SUCCESS, "True", "False") & ", Mode: " & Hex$(lRequestMode)
        #End If

    End If

End Function
    
Friend Function SetPriority(ByVal hLineApp As Long, _
    ByVal lPriorityMode As Long, _
    ByVal lRequestMode As Long, _
    ByVal bPriority As Boolean) As Boolean
    
    Dim uLineExtID  As LINEEXTENSIONID
    Dim rc As Long
    
    If mhOpenLine <> 0& Then
    
        rc = lineSetAppPriority( _
            App.Title & vbNullString, _
            lPriorityMode, _
            uLineExtID, _
            lRequestMode, _
            "", _
            IIf(bPriority, 1, 0))
    
        If rc <> TAPI_SUCCESS Then
            mlLastTAPIError = rc
        Else
            SetPriority = True
        End If
    
        #If LevelA = 1 Then
            Debug.Print "CTAPIDevice(" & CStr(mlDeviceIndex) & ")." & IIf(bPriority, "Set", "Clear") & "Priority Result: " & IIf(rc = TAPI_SUCCESS, "True", "False") & ", Priority: " & Hex$(lPriorityMode)
        #End If
    
    End If
    
End Function

Friend Function GetPriority(ByVal hLineApp As Long, _
    ByVal lPriorityMode As Long, _
    ByVal lRequestMode As Long, _
    ByRef bPriority As Boolean) As Boolean
    
    
    Dim uLineExtID  As LINEEXTENSIONID
    Dim rc As Long
    Dim lPriority As Long
    
    If mhOpenLine <> 0& Then
    
        rc = lineSetAppPriority( _
            App.Title & vbNullString, _
            lPriorityMode, _
            uLineExtID, _
            lRequestMode, _
            "", _
            lPriority)
    
        If rc <> TAPI_SUCCESS Then
            mlLastTAPIError = rc
        Else
            bPriority = CBool(lPriority)
            GetPriority = True
        End If
    
    End If
        
End Function


Public Function OpenConfigDialog(Optional ByVal hOwnerHwnd As Long = 0&, _
    Optional ByVal sDeviceClass As String = "") As Boolean
    
    Dim rc As Long
    
    If Len(sDeviceClass) = 0 Then
        rc = lineConfigDialog(mlDeviceIndex, hOwnerHwnd, vbNullString)
    Else
        rc = lineConfigDialog(mlDeviceIndex, hOwnerHwnd, sDeviceClass)
    End If
    
    If rc <> TAPI_SUCCESS Then
        mlLastTAPIError = rc
    Else
        OpenConfigDialog = True
    End If
    
End Function

'this is called from the App class because this fn needs these first 2 parms...
'
Friend Function OpenDialingPropDialog(ByVal hLineApp As Long, _
    ByVal lAppVersion As Long, _
    ByVal hOwnerHwnd As Long, _
    ByVal sPhoneNumber As String) As Boolean
    
    Dim rc As Long
    If Len(sPhoneNumber) = 0 Then
        rc = lineTranslateDialog(hLineApp, mlDeviceIndex, lAppVersion, hOwnerHwnd, vbNullString)
    Else
        rc = lineTranslateDialog(hLineApp, mlDeviceIndex, lAppVersion, hOwnerHwnd, sPhoneNumber)
    End If
    
    If rc <> TAPI_SUCCESS Then
        mlLastTAPIError = rc
    Else
        OpenDialingPropDialog = True
    End If
    
End Function



