Attribute VB_Name = "mTAPITypes"
' TAPI Type Defs and Declarations

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit


Public Type LINEINITIALIZEEXPARAMS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwOptions As Long
    hEvent As Long          'union hEvent and Completion port
    dwCompletionKey As Long
End Type

Public Type LINEMONITORTONE
    dwAppSpecific As Long
    dwDuration As Long
    dwFrequency1 As Long
    dwFrequency2 As Long
    dwFrequency3 As Long
End Type

Public Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type

Public Type LINECALLINFO
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    hLine As Long
    dwLineDeviceID As Long
    dwAddressID As Long

    dwBearerMode As Long
    dwRate As Long
    dwMediaMode As Long

    dwAppSpecific As Long
    dwCallID As Long
    dwRelatedCallID As Long
    dwCallParamFlags As Long
    dwCallStates As Long

    dwMonitorDigitModes As Long
    dwMonitorMediaModes As Long
    DialParams As LINEDIALPARAMS

    dwOrigin As Long
    dwReason As Long
    dwCompletionID As Long
    dwNumOwners As Long
    dwNumMonitors As Long

    dwCountryCode As Long
    dwTrunk As Long

    dwCallerIDFlags As Long
    dwCallerIDSize As Long
    dwCallerIDOffset As Long
    dwCallerIDNameSize As Long
    dwCallerIDNameOffset As Long

    dwCalledIDFlags As Long
    dwCalledIDSize As Long
    dwCalledIDOffset As Long
    dwCalledIDNameSize As Long
    dwCalledIDNameOffset As Long

    dwConnectedIDFlags As Long
    dwConnectedIDSize As Long
    dwConnectedIDOffset As Long
    dwConnectedIDNameSize As Long
    dwConnectedIDNameOffset As Long

    dwRedirectionIDFlags As Long
    dwRedirectionIDSize As Long
    dwRedirectionIDOffset As Long
    dwRedirectionIDNameSize As Long
    dwRedirectionIDNameOffset As Long

    dwRedirectingIDFlags As Long
    dwRedirectingIDSize As Long
    dwRedirectingIDOffset As Long
    dwRedirectingIDNameSize As Long
    dwRedirectingIDNameOffset As Long

    dwAppNameSize As Long
    dwAppNameOffset As Long

    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long

    dwCalledPartySize As Long
    dwCalledPartyOffset As Long

    dwCommentSize As Long
    dwCommentOffset As Long

    dwDisplaySize As Long
    dwDisplayOffset As Long

    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long

    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long

    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long

    dwChargingInfoSize As Long
    dwChargingInfoOffset As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    
    mem As String * 2048        'this buffer not changed to Byte() array since there
                                'is an existing method to pluck the substring out of this string
                                '(and it doesn't error out like the other buffer(s) do as fixed-len strings)
End Type

Public Const LINECALLINFO_FIXEDSIZE = 296


Public Type LINECALLPARAMS                 '// DEFAULTS
    dwTotalSize As Long                    '// ---------
    dwBearerMode As Long                   '// voice
    dwMinRate As Long                      '// (3.1kHz)
    dwMaxRate As Long                      '// (3.1kHz)
    dwMediaMode As Long                    '// interactiveVoice
    dwCallParamFlags As Long               '// 0
    dwAddressMode As Long                  '// addressID
    dwAddressID As Long                    '// (any available)
    DialParams As LINEDIALPARAMS           '// (0, 0, 0, 0)
    dwOrigAddressSize As Long              '// 0
    dwOrigAddressOffset As Long
    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long
    dwCalledPartySize As Long              '// 0
    dwCalledPartyOffset As Long
    dwCommentSize As Long                  '// 0
    dwCommentOffset As Long
    dwUserUserInfoSize As Long             '// 0
    dwUserUserInfoOffset As Long
    dwHighLevelCompSize As Long            '// 0
    dwHighLevelCompOffset As Long
    dwLowLevelCompSize As Long             '// 0
    dwLowLevelCompOffset As Long
    dwDevSpecificSize As Long              '// 0
    dwDevSpecificOffset As Long

    dwPredictiveAutoTransferStates As Long                 '// TAPI v2.0
    dwTargetAddressSize As Long                            '// TAPI v2.0
    dwTargetAddressOffset As Long                          '// TAPI v2.0
    dwSendingFlowspecSize As Long                          '// TAPI v2.0
    dwSendingFlowspecOffset As Long                        '// TAPI v2.0
    dwReceivingFlowspecSize As Long                        '// TAPI v2.0
    dwReceivingFlowspecOffset As Long                      '// TAPI v2.0
    dwDeviceClassSize As Long                              '// TAPI v2.0
    dwDeviceClassOffset As Long                            '// TAPI v2.0
    dwDeviceConfigSize As Long                             '// TAPI v2.0
    dwDeviceConfigOffset As Long                           '// TAPI v2.0
    dwCallDataSize As Long                                 '// TAPI v2.0
    dwCallDataOffset As Long                               '// TAPI v2.0
    dwNoAnswerTimeout As Long                              '// TAPI v2.0
    dwCallingPartyIDSize As Long                           '// TAPI v2.0
    dwCallingPartyIDOffset As Long                         '// TAPI v2.0

    mem(0 To 2048) As Byte       ' padding

End Type


Public Type LINEDEVCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long

    dwSwitchInfoSize As Long
    dwSwitchInfoOffset As Long

    dwPermanentLineID As Long
    dwLineNameSize As Long
    dwLineNameOffset As Long
    dwStringFormat As Long

    dwAddressModes As Long
    dwNumAddresses As Long
    dwBearerModes As Long
    dwMaxRate As Long
    dwMediaModes As Long

    dwGenerateToneModes As Long
    dwGenerateToneMaxNumFreq As Long
    dwGenerateDigitModes As Long
    dwMonitorToneMaxNumFreq As Long
    dwMonitorToneMaxNumEntries As Long
    dwMonitorDigitModes As Long
    dwGatherDigitsMinTimeout As Long
    dwGatherDigitsMaxTimeout As Long

    dwMedCtlDigitMaxListSize As Long
    dwMedCtlMediaMaxListSize As Long
    dwMedCtlToneMaxListSize As Long
    dwMedCtlCallStateMaxListSize As Long

    dwDevCapFlags As Long
    dwMaxNumActiveCalls As Long
    dwAnswerMode As Long
    dwRingModes As Long
    dwLineStates As Long

    dwUUIAcceptSize As Long
    dwUUIAnswerSize As Long
    dwUUIMakeCallSize As Long
    dwUUIDropSize As Long
    dwUUISendUserInfoSize As Long
    dwUUICallInfoSize As Long

    MinDialParams As LINEDIALPARAMS
    MaxDialParams As LINEDIALPARAMS
    DefaultDialParams As LINEDIALPARAMS

    dwNumTerminals As Long
    dwTerminalCapsSize As Long
    dwTerminalCapsOffset As Long
    dwTerminalTextEntrySize As Long
    dwTerminalTextSize As Long
    dwTerminalTextOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    dwLineFeatures As Long                  ' TAPI v1.4

    dwSettableDevStatus As Long             ' TAPI v2.0
    dwDeviceClassesSize As Long             ' TAPI v2.0
    dwDeviceClassesOffset As Long           ' TAPI v2.0

    mem(0 To 2048) As Byte
    'with LINEERR_STRUCTURETOOSMALL you may need to enlarge the buffer
End Type


Public Type LPPROVIDERLIST
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    
    dwNumProviders As Long
    dwProviderListSize As Long
    dwProviderListOffset As Long

    mem(0 To 2048) As Byte
End Type

Public Type LINEADDRESSCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwLineDeviceID As Long

    dwAddressSize As Long
    dwAddressOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long

    dwAddressSharing As Long
    dwAddressStates As Long
    dwCallInfoStates As Long
    dwCallerIDFlags As Long
    dwCalledIDFlags As Long
    dwConnectedIDFlags As Long
    dwRedirectionIDFlags As Long
    dwRedirectingIDFlags As Long
    dwCallStates As Long
    dwDialToneModes As Long
    dwBusyModes As Long
    dwSpecialInfo As Long
    dwDisconnectModes As Long

    dwMaxNumActiveCalls As Long
    dwMaxNumOnHoldCalls As Long
    dwMaxNumOnHoldPendingCalls As Long
    dwMaxNumConference As Long
    dwMaxNumTransConf As Long

    dwAddrCapFlags As Long
    dwCallFeatures As Long
    dwRemoveFromConfCaps As Long
    dwRemoveFromConfState As Long
    dwTransferModes As Long
    dwParkModes As Long

    dwForwardModes As Long
    dwMaxForwardEntries As Long
    dwMaxSpecificEntries As Long
    dwMinFwdNumRings As Long
    dwMaxFwdNumRings As Long

    dwMaxCallCompletions As Long
    dwCallCompletionConds As Long
    dwCallCompletionModes As Long
    dwNumCompletionMessages As Long
    dwCompletionMsgTextEntrySize As Long
    dwCompletionMsgTextSize As Long
    dwCompletionMsgTextOffset As Long
    
    dwPredictiveAutoTransferStates As Long
    dwNumCallTreatments As Long
    dwCallTreatmentListSize As Long
    dwCallTreatmentListOffset As Long
    dwDeviceClassesSize As Long
    dwDeviceClassesOffset As Long
    dwMaxCallDataSize As Long
    dwCallFeatures2 As Long
    dwMaxNoAnswerTimeout As Long
    dwConnectedModes As Long
    dwOfferingModes As Long
    dwAvailableMediaModes As Long

    mem(0 To 2048) As Byte
End Type

Public Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type

Public Type LINETRANSLATECAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumLocations As Long
    dwLocationListSize As Long
    dwLocationListOffset As Long

    dwCurrentLocationID As Long

    dwNumCards As Long
    dwCardListSize As Long
    dwCardListOffset As Long

    dwCurrentPreferredCardID As Long

    mem(0 To 2048) As Byte
    
End Type


Type LINETERMCAPS
    dwTermDev As Long
    dwTermModes As Long
    dwTermSharing As Long
End Type


'==============================================

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                        (dest As Any, src As Any, ByVal length As Long)

Public Declare Function lineInitialize Lib "TAPI32.DLL" _
    (ByRef lphLineApp As Long, _
    ByVal hInstance As Long, _
    ByVal lpfnCallback As Long, _
    ByVal lpszAppName As String, _
    ByRef lpdwNumDevs As Long) As Long

Public Declare Function lineInitializeEx Lib "TAPI32.DLL" Alias "lineInitializeExA" _
    (ByRef lphLineApp As Long, _
    ByVal hInstance As Long, _
    ByVal lpfnCallback As Long, _
    ByVal lpszFriendlyAppName As String, _
    ByRef lpdwNumDevs As Long, _
    ByRef lpdwAPIVersion As Long, _
    ByRef lpLineInitializeExParams As LINEINITIALIZEEXPARAMS) As Long
    
Public Declare Function lineGetDevCaps Lib "TAPI32.DLL" Alias "lineGetDevCapsA" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtVersion As Long, _
    ByRef lpLineDevCaps As LINEDEVCAPS) As Long
 
Public Declare Function lineConfigDialog Lib "TAPI32.DLL" Alias "lineConfigDialogA" _
    (ByVal dwDeviceID As Long, _
    ByVal hwndOwner As Long, _
    ByVal lpszDeviceClass As String) As Long
    
Public Declare Function lineTranslateDialog Lib "TAPI32.DLL" Alias "lineTranslateDialogA" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal hwndOwner As Long, _
    ByVal lpszAddressIn As String) As Long

Public Declare Function lineShutdown Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long) As Long
    
Public Declare Function lineMakeCall Lib "TAPI32.DLL" Alias "lineMakeCallA" _
    (ByVal hLine As Long, _
    ByRef lphCall As Long, _
    ByVal lpszDestAddress As String, _
    ByVal dwCountryCode As Long, _
    ByRef lpCallParams As Any) As Long 'LINECALLPARAMS declared As Any so Null value can be passed
    
Public Declare Function lineDeallocateCall Lib "TAPI32.DLL" _
    (ByVal hCall As Long) As Long
    
Public Declare Function lineDrop Lib "TAPI32.DLL" _
    (ByVal hCall As Long, _
    ByVal lpsUserUserInfo As String, _
    ByVal dwSize As Long) As Long
    
Public Declare Function lineGetIcon Lib "TAPI32.DLL" Alias "lineGetIconA" _
    (ByVal dwDeviceID As Long, _
    ByVal lpszDeviceClass As Long, _
    ByRef lphIcon As Long) As Long
    

Public Declare Function lineNegotiateAPIVersion Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPILowVersion As Long, _
    ByVal dwAPIHighVersion As Long, _
    ByRef lpdwAPIVersion As Long, _
    ByRef lpExtensionID As LINEEXTENSIONID) As Long


Public Declare Function lineClose Lib "TAPI32.DLL" _
    (ByVal hLine As Long) As Long

Public Declare Function lineOpen Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByRef lphLine As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtVersion As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwPrivileges As Long, _
    ByVal dwMediaModes As Long, _
    ByRef lpCallParams As Any) As Long
    
    'LINECALLPARAMS declared As Any so NULL can be passed

Public Declare Function lineRegisterRequestRecipient Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long, _
    ByVal dwRegistrationInstance As Long, _
    ByVal dwRequestMode As Long, _
    ByVal bEnable As Long) As Long

Public Declare Function lineGetAppPriority Lib "TAPI32.DLL" _
    (ByVal lpszAppFilename As String, _
    ByVal dwMediaMode As Long, _
    ByRef lpExtensionID As LINEEXTENSIONID, _
    ByVal dwRequestMode As Long, _
    ByRef lpExtensionName As String, _
    ByRef lpdwPriority As Long) As Long
    
Public Declare Function lineSetAppPriority Lib "TAPI32.DLL" _
    (ByVal lpszAppFilename As String, _
    ByVal dwMediaMode As Long, _
    ByRef lpExtensionID As LINEEXTENSIONID, _
    ByVal dwRequestMode As Long, _
    ByVal lpszExtensionName As String, _
    ByVal dwPriority As Long) As Long

Public Declare Function lineGetCallInfo Lib "TAPI32.DLL" _
    (ByVal hCall As Long, _
    ByRef lpCallInfo As LINECALLINFO) As Long

Public Declare Function DrawIconEx Lib "user32.dll" _
                            (ByVal hdc As Long, ByVal left As Long, ByVal top As Long, ByVal hIcon As Long, _
                            ByVal width As Long, ByVal height As Long, ByVal step As Long, ByVal hBrush As Long, _
                            ByVal uFlags As Long) As Long
                            
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long  'BOOL



