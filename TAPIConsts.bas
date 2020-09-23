Attribute VB_Name = "mTAPIConsts"
' TAPI Constants

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit

Global Const TAPI_SUCCESS As Long = 0& 'declared for convenience
Global Const LINEERR_ALLOCATED As Long = &H80000001
Global Const LINEERR_BADDEVICEID As Long = &H80000002
Global Const LINEERR_BEARERMODEUNAVAIL As Long = &H80000003
Global Const LINEERR_CALLUNAVAIL As Long = &H80000005
Global Const LINEERR_COMPLETIONOVERRUN As Long = &H80000006
Global Const LINEERR_CONFERENCEFULL As Long = &H80000007
Global Const LINEERR_DIALBILLING As Long = &H80000008
Global Const LINEERR_DIALDIALTONE As Long = &H80000009
Global Const LINEERR_DIALPROMPT As Long = &H8000000A
Global Const LINEERR_DIALQUIET As Long = &H8000000B
Global Const LINEERR_INCOMPATIBLEAPIVERSION As Long = &H8000000C
Global Const LINEERR_INCOMPATIBLEEXTVERSION As Long = &H8000000D
Global Const LINEERR_INIFILECORRUPT As Long = &H8000000E
Global Const LINEERR_INUSE As Long = &H8000000F
Global Const LINEERR_INVALADDRESS As Long = &H80000010
Global Const LINEERR_INVALADDRESSID As Long = &H80000011
Global Const LINEERR_INVALADDRESSMODE As Long = &H80000012
Global Const LINEERR_INVALADDRESSSTATE As Long = &H80000013
Global Const LINEERR_INVALAPPHANDLE As Long = &H80000014
Global Const LINEERR_INVALAPPNAME As Long = &H80000015
Global Const LINEERR_INVALBEARERMODE As Long = &H80000016
Global Const LINEERR_INVALCALLCOMPLMODE As Long = &H80000017
Global Const LINEERR_INVALCALLHANDLE As Long = &H80000018
Global Const LINEERR_INVALCALLPARAMS As Long = &H80000019
Global Const LINEERR_INVALCALLPRIVILEGE As Long = &H8000001A
Global Const LINEERR_INVALCALLSELECT As Long = &H8000001B
Global Const LINEERR_INVALCALLSTATE As Long = &H8000001C
Global Const LINEERR_INVALCALLSTATELIST As Long = &H8000001D
Global Const LINEERR_INVALCARD As Long = &H8000001E
Global Const LINEERR_INVALCOMPLETIONID As Long = &H8000001F
Global Const LINEERR_INVALCONFCALLHANDLE As Long = &H80000020
Global Const LINEERR_INVALCONSULTCALLHANDLE As Long = &H80000021
Global Const LINEERR_INVALCOUNTRYCODE As Long = &H80000022
Global Const LINEERR_INVALDEVICECLASS As Long = &H80000023
Global Const LINEERR_INVALDEVICEHANDLE As Long = &H80000024
Global Const LINEERR_INVALDIGITLIST As Long = &H80000026
Global Const LINEERR_INVALDIGITMODE As Long = &H80000027
Global Const LINEERR_INVALDIGITS As Long = &H80000028
Global Const LINEERR_INVALEXTVERSION As Long = &H80000029
Global Const LINEERR_INVALGROUPID As Long = &H8000002A
Global Const LINEERR_INVALLINEHANDLE As Long = &H8000002B
Global Const LINEERR_INVALLINESTATE As Long = &H8000002C
Global Const LINEERR_INVALLOCATION As Long = &H8000002D
Global Const LINEERR_INVALMEDIALIST As Long = &H8000002E
Global Const LINEERR_INVALMEDIAMODE As Long = &H8000002F
Global Const LINEERR_INVALMESSAGEID As Long = &H80000030
Global Const LINEERR_INVALPARAM As Long = &H80000032
Global Const LINEERR_INVALPARKID As Long = &H80000033
Global Const LINEERR_INVALPARKMODE As Long = &H80000034
Global Const LINEERR_INVALPOINTER As Long = &H80000035
Global Const LINEERR_INVALPRIVSELECT As Long = &H80000036
Global Const LINEERR_INVALRATE As Long = &H80000037
Global Const LINEERR_INVALREQUESTMODE As Long = &H80000038
Global Const LINEERR_INVALTERMINALID As Long = &H80000039
Global Const LINEERR_INVALTERMINALMODE As Long = &H8000003A
Global Const LINEERR_INVALTIMEOUT As Long = &H8000003B
Global Const LINEERR_INVALTONE As Long = &H8000003C
Global Const LINEERR_INVALTONELIST As Long = &H8000003D
Global Const LINEERR_INVALTONEMODE As Long = &H8000003E
Global Const LINEERR_INVALTRANSFERMODE As Long = &H8000003F
Global Const LINEERR_LINEMAPPERFAILED As Long = &H80000040
Global Const LINEERR_NOCONFERENCE As Long = &H80000041
Global Const LINEERR_NODEVICE As Long = &H80000042
Global Const LINEERR_NODRIVER As Long = &H80000043
Global Const LINEERR_NOMEM As Long = &H80000044
Global Const LINEERR_NOREQUEST As Long = &H80000045
Global Const LINEERR_NOTOWNER As Long = &H80000046
Global Const LINEERR_NOTREGISTERED As Long = &H80000047
Global Const LINEERR_OPERATIONFAILED As Long = &H80000048
Global Const LINEERR_OPERATIONUNAVAIL As Long = &H80000049
Global Const LINEERR_RATEUNAVAIL As Long = &H8000004A
Global Const LINEERR_RESOURCEUNAVAIL As Long = &H8000004B
Global Const LINEERR_REQUESTOVERRUN As Long = &H8000004C
Global Const LINEERR_STRUCTURETOOSMALL As Long = &H8000004D
Global Const LINEERR_TARGETNOTFOUND As Long = &H8000004E
Global Const LINEERR_TARGETSELF As Long = &H8000004F
Global Const LINEERR_UNINITIALIZED As Long = &H80000050
Global Const LINEERR_USERUSERINFOTOOBIG As Long = &H80000051
Global Const LINEERR_REINIT As Long = &H80000052
Global Const LINEERR_ADDRESSBLOCKED As Long = &H80000053
Global Const LINEERR_BILLINGREJECTED As Long = &H80000054
Global Const LINEERR_INVALFEATURE As Long = &H80000055
Global Const LINEERR_NOMULTIPLEINSTANCE As Long = &H80000056

Global Const LINEFEATURE_DEVSPECIFIC As Long = &H1&
Global Const LINEFEATURE_DEVSPECIFICFEAT As Long = &H2&
Global Const LINEFEATURE_FORWARD As Long = &H4&
Global Const LINEFEATURE_MAKECALL As Long = &H8&
Global Const LINEFEATURE_SETMEDIACONTROL As Long = &H10&
Global Const LINEFEATURE_SETTERMINAL As Long = &H20&

Global Const LINECALLFEATURE_ACCEPT As Long = &H1&
Global Const LINECALLFEATURE_ADDTOCONF As Long = &H2&
Global Const LINECALLFEATURE_ANSWER As Long = &H4&
Global Const LINECALLFEATURE_BLINDTRANSFER As Long = &H8&
Global Const LINECALLFEATURE_COMPLETECALL As Long = &H10&
Global Const LINECALLFEATURE_COMPLETETRANSF As Long = &H20&
Global Const LINECALLFEATURE_DIAL As Long = &H40&
Global Const LINECALLFEATURE_DROP As Long = &H80&
Global Const LINECALLFEATURE_GATHERDIGITS As Long = &H100&
Global Const LINECALLFEATURE_GENERATEDIGITS As Long = &H200&
Global Const LINECALLFEATURE_GENERATETONE As Long = &H400&
Global Const LINECALLFEATURE_HOLD As Long = &H800&
Global Const LINECALLFEATURE_MONITORDIGITS As Long = &H1000&
Global Const LINECALLFEATURE_MONITORMEDIA As Long = &H2000&
Global Const LINECALLFEATURE_MONITORTONES As Long = &H4000&
Global Const LINECALLFEATURE_PARK As Long = &H8000&
Global Const LINECALLFEATURE_PREPAREADDCONF As Long = &H10000
Global Const LINECALLFEATURE_REDIRECT As Long = &H20000
Global Const LINECALLFEATURE_REMOVEFROMCONF As Long = &H40000
Global Const LINECALLFEATURE_SECURECALL As Long = &H80000
Global Const LINECALLFEATURE_SENDUSERUSER As Long = &H100000
Global Const LINECALLFEATURE_SETCALLPARAMS As Long = &H200000
Global Const LINECALLFEATURE_SETMEDIACONTROL As Long = &H400000
Global Const LINECALLFEATURE_SETTERMINAL As Long = &H800000
Global Const LINECALLFEATURE_SETUPCONF As Long = &H1000000
Global Const LINECALLFEATURE_SETUPTRANSFER As Long = &H2000000
Global Const LINECALLFEATURE_SWAPHOLD As Long = &H4000000
Global Const LINECALLFEATURE_UNHOLD As Long = &H8000000




'// These constants are mutually exclusive - there's no way to specify more
'// than one at a time (and it doesn't make sense, either) so they're
'// ordinal rather than bits.
'//
Global Const LINEINITIALIZEEXOPTION_USEHIDDENWINDOW   As Long = &H1&        '// TAPI v2.0
Global Const LINEINITIALIZEEXOPTION_USEEVENT    As Long = &H2&         '// TAPI v2.0
Global Const LINEINITIALIZEEXOPTION_USECOMPLETIONPORT    As Long = &H3&         '// TAPI v2.0

'// Messages for Phones and Lines

Global Const LINE_ADDRESSSTATE               As Long = 0&
Global Const LINE_CALLINFO                   As Long = 1&
Global Const LINE_CALLSTATE                  As Long = 2&
Global Const LINE_CLOSE                      As Long = 3&
Global Const LINE_DEVSPECIFIC                As Long = 4&
Global Const LINE_DEVSPECIFICFEATURE         As Long = 5&
Global Const LINE_GATHERDIGITS               As Long = 6&
Global Const LINE_GENERATE                   As Long = 7&
Global Const LINE_LINEDEVSTATE               As Long = 8&
Global Const LINE_MONITORDIGITS              As Long = 9&
Global Const LINE_MONITORMEDIA               As Long = 10&
Global Const LINE_MONITORTONE                As Long = 11&
Global Const LINE_REPLY                      As Long = 12&
Global Const LINE_REQUEST                    As Long = 13&
Global Const PHONE_BUTTON                    As Long = 14&
Global Const PHONE_CLOSE                     As Long = 15&
Global Const PHONE_DEVSPECIFIC               As Long = 16&
Global Const PHONE_REPLY                     As Long = 17&
Global Const PHONE_STATE                     As Long = 18&
Global Const LINE_CREATE                     As Long = 19&          ' TAPI v1.4
Global Const PHONE_CREATE                    As Long = 20&          ' TAPI v1.4

Global Const LINE_AGENTSPECIFIC              As Long = 21&          ' TAPI v2.0
Global Const LINE_AGENTSTATUS                As Long = 22&          ' TAPI v2.0
Global Const LINE_APPNEWCALL                 As Long = 23&          ' TAPI v2.0
Global Const LINE_PROXYREQUEST               As Long = 24&          ' TAPI v2.0
Global Const LINE_REMOVE                     As Long = 25&          ' TAPI v2.0
Global Const PHONE_REMOVE                    As Long = 26&          ' TAPI v2.0



Global Const LINEDEVCAPS_FIXEDSIZE = 236
Global Const LINEDEVSTATUS_FIXEDSIZE = 76

Global Const LINEADDRESSMODE_ADDRESSID = &H1&
Global Const LINEADDRESSMODE_DIALABLEADDR = &H2&

Global Const LINEDEVCAPFLAGS_CROSSADDRCONF = &H1
Global Const LINEDEVCAPFLAGS_HIGHLEVCOMP = &H2
Global Const LINEDEVCAPFLAGS_LOWLEVCOMP = &H4
Global Const LINEDEVCAPFLAGS_MEDIACONTROL = &H8
Global Const LINEDEVCAPFLAGS_MULTIPLEADDR = &H10
Global Const LINEDEVCAPFLAGS_CLOSEDROP = &H20
Global Const LINEDEVCAPFLAGS_DIALBILLING = &H40
Global Const LINEDEVCAPFLAGS_DIALQUIET = &H80
Global Const LINEDEVCAPFLAGS_DIALDIALTONE = &H100

Global Const LINEDIGITMODE_PULSE = &H1&
Global Const LINEDIGITMODE_DTMF = &H2&
Global Const LINEDIGITMODE_DTMFEND = &H4&

Global Const LINEANSWERMODE_NONE = &H1&
Global Const LINEANSWERMODE_DROP = &H2&
Global Const LINEANSWERMODE_HOLD = &H4&




Enum EnumTAPICallStates
    LINECALLSTATE_IDLE = &H1&
    LINECALLSTATE_OFFERING = &H2&
    LINECALLSTATE_ACCEPTED = &H4&
    LINECALLSTATE_DIALTONE = &H8&
    LINECALLSTATE_DIALING = &H10&
    LINECALLSTATE_RINGBACK = &H20&
    LINECALLSTATE_BUSY = &H40&
    LINECALLSTATE_SPECIALINFO = &H80&
    LINECALLSTATE_CONNECTED = &H100&
    LINECALLSTATE_PROCEEDING = &H200&
    LINECALLSTATE_ONHOLD = &H400&
    LINECALLSTATE_CONFERENCED = &H800&
    LINECALLSTATE_ONHOLDPENDCONF = &H1000&
    LINECALLSTATE_ONHOLDPENDTRANSFER = &H2000&
    LINECALLSTATE_DISCONNECTED = &H4000&
    LINECALLSTATE_UNKNOWN = &H8000&

    LINECALLTREATMENT_SILENCE = &H1&            '// TAPI v2.0
    LINECALLTREATMENT_RINGBACK = &H2&           '// TAPI v2.0
    LINECALLTREATMENT_BUSY = &H3&               '// TAPI v2.0
    LINECALLTREATMENT_MUSIC = &H4&              '// TAPI v2.0
End Enum

Enum EnumTAPICallInfoStates
     LINECALLINFOSTATE_OTHER = &H1&
     LINECALLINFOSTATE_DEVSPECIFIC = &H2&
     LINECALLINFOSTATE_BEARERMODE = &H4&
     LINECALLINFOSTATE_RATE = &H8&
     LINECALLINFOSTATE_MEDIAMODE = &H10&
     LINECALLINFOSTATE_APPSPECIFIC = &H20&
     LINECALLINFOSTATE_CALLID = &H40&
     LINECALLINFOSTATE_RELATEDCALLID = &H80&
     LINECALLINFOSTATE_ORIGIN = &H100&
     LINECALLINFOSTATE_REASON = &H200&
     LINECALLINFOSTATE_COMPLETIONID = &H400&
     LINECALLINFOSTATE_NUMOWNERINCR = &H800&
     LINECALLINFOSTATE_NUMOWNERDECR = &H1000&
     LINECALLINFOSTATE_NUMMONITORS = &H2000&
     LINECALLINFOSTATE_TRUNK = &H4000&
     LINECALLINFOSTATE_CALLERID = &H8000&
     LINECALLINFOSTATE_CALLEDID = &H10000
     LINECALLINFOSTATE_CONNECTEDID = &H20000
     LINECALLINFOSTATE_REDIRECTIONID = &H40000
     LINECALLINFOSTATE_REDIRECTINGID = &H80000
     LINECALLINFOSTATE_DISPLAY = &H100000
     LINECALLINFOSTATE_USERUSERINFO = &H200000
     LINECALLINFOSTATE_HIGHLEVELCOMP = &H400000
     LINECALLINFOSTATE_LOWLEVELCOMP = &H800000
     LINECALLINFOSTATE_CHARGINGINFO = &H1000000
     LINECALLINFOSTATE_TERMINAL = &H2000000
     LINECALLINFOSTATE_DIALPARAMS = &H4000000
     LINECALLINFOSTATE_MONITORMODES = &H8000000
End Enum


Enum EnumTAPICallPrivileges
    LINECALLPRIVILEGE_NONE = &H1&
    LINECALLPRIVILEGE_MONITOR = &H2&
    LINECALLPRIVILEGE_OWNER = &H4&
End Enum

Enum EnumTAPIStringFormats
    STRINGFORMAT_ASCII = &H1&
    STRINGFORMAT_DBCS = &H2&
    STRINGFORMAT_UNICODE = &H3&
    STRINGFORMAT_BINARY = &H4&
End Enum

Enum EnumTAPIRequestModes
    LINEREQUESTMODE_MAKECALL = &H1&
    LINEREQUESTMODE_MEDIACALL = &H2&
    LINEREQUESTMODE_DROP = &H4&
End Enum

Enum EnumTAPIBearerModes
    LINEBEARERMODE_VOICE = &H1&
    LINEBEARERMODE_SPEECH = &H2&
    LINEBEARERMODE_MULTIUSE = &H4&
    LINEBEARERMODE_DATA = &H8&
    LINEBEARERMODE_ALTSPEECHDATA = &H10&
    LINEBEARERMODE_NONCALLSIGNALING = &H20&
End Enum

Enum EnumTAPIMediaModes
    LINEMEDIAMODE_UNKNOWN = &H2&
    LINEMEDIAMODE_INTERACTIVEVOICE = &H4&
    LINEMEDIAMODE_AUTOMATEDVOICE = &H8&
    LINEMEDIAMODE_DATAMODEM = &H10&
    LINEMEDIAMODE_G3FAX = &H20&
    LINEMEDIAMODE_TDD = &H40&
    LINEMEDIAMODE_G4FAX = &H80&
    LINEMEDIAMODE_DIGITALDATA = &H100&
    LINEMEDIAMODE_TELETEX = &H200&
    LINEMEDIAMODE_VIDEOTEX = &H400&
    LINEMEDIAMODE_TELEX = &H800&
    LINEMEDIAMODE_MIXED = &H1000&
    LINEMEDIAMODE_ADSI = &H2000&
End Enum

Enum EnumTAPILineToneModes
    LINETONEMODE_CUSTOM = &H1&
    LINETONEMODE_RINGBACK = &H2&
    LINETONEMODE_BUSY = &H4&
    LINETONEMODE_BEEP = &H8&
    LINETONEMODE_BILLING = &H10&
End Enum


Enum EnumCallPartyStatus
    LINECALLPARTYID_BLOCKED = &H1&
    LINECALLPARTYID_OUTOFAREA = &H2&
    LINECALLPARTYID_NAME = &H4&
    LINECALLPARTYID_ADDRESS = &H8&
    LINECALLPARTYID_PARTIAL = &H10&
    LINECALLPARTYID_UNKNOWN = &H20&
    LINECALLPARTYID_UNAVAIL = &H40&
End Enum

Enum EnumDrawIconConstants
    DI_MASK = &H1&
    DI_IMAGE = &H2&
    DI_NORMAL = &H3&
    DI_COMPAT = &H4&
    DI_DEFAULTSIZE = &H8&
End Enum


Public Function GetLineErrString(lparam As Long) As String
'Returns a String description of a TAPI Line Error code
    Dim msg As String
    
    Select Case lparam
        Case LINEERR_ALLOCATED '( = &H80000001)
            msg = "Allocated"
        Case LINEERR_BADDEVICEID '(= &H80000002)
            msg = "Bad Device ID"
        Case LINEERR_BEARERMODEUNAVAIL '(= &H80000003)
            msg = "Bearer Mode Unavail"
        Case LINEERR_CALLUNAVAIL '(= &H80000005)
            msg = "Call UnAvail"
        Case LINEERR_COMPLETIONOVERRUN '(= &H80000006
            msg = "Completion Overrun"
        Case LINEERR_CONFERENCEFULL '(= &H80000007
            msg = "Conference Full"
        Case LINEERR_DIALBILLING '(= &H80000008
            msg = "Dial Billing"
        Case LINEERR_DIALDIALTONE '(= &H80000009
            msg = "Dial Dialtone"
        Case LINEERR_DIALPROMPT '(= &H8000000A
            msg = "Dial Prompt"
        Case LINEERR_DIALQUIET '(= &H8000000B
            msg = "Dial Quiet"
        Case LINEERR_INCOMPATIBLEAPIVERSION '(= &H8000000C
            msg = "Incompatible API Version"
        Case LINEERR_INCOMPATIBLEEXTVERSION '(= &H8000000D
            msg = "Incompatible Ext Version"
        Case LINEERR_INIFILECORRUPT '(= &H8000000E
            msg = "Ini File Corrupt"
        Case LINEERR_INUSE '(= &H8000000F
            msg = "In Use"
        Case LINEERR_INVALADDRESS '(= &H80000010
            msg = "Invalid Address"
        Case LINEERR_INVALADDRESSID '(= &H80000011
            msg = "Invalid Address ID"
        Case LINEERR_INVALADDRESSMODE '(= &H80000012
            msg = "Invalid Address Mode"
        Case LINEERR_INVALADDRESSSTATE '(= &H80000013
            msg = "Invalid Address State"
        Case LINEERR_INVALAPPHANDLE '(= &H80000014
            msg = "Invalid App Handle"
        Case LINEERR_INVALAPPNAME '(= &H80000015
            msg = "Invalid App Name"
        Case LINEERR_INVALBEARERMODE '(= &H80000016
            msg = "Invalid Bearer Mode"
        Case LINEERR_INVALCALLCOMPLMODE '(= &H80000017
            msg = "Invalid Call Completion Mode"
        Case LINEERR_INVALCALLHANDLE '(= &H80000018
            msg = "Invalid Call Handle"
        Case LINEERR_INVALCALLPARAMS '(= &H80000019
            msg = "Invalid Call Params"
        Case LINEERR_INVALCALLPRIVILEGE '(= &H8000001A
            msg = "Invalid Call Privilege"
        Case LINEERR_INVALCALLSELECT '(= &H8000001B
            msg = "Invalid Call Select"
        Case LINEERR_INVALCALLSTATE '(= &H8000001C
            msg = "Invalid Call State"
        Case LINEERR_INVALCALLSTATELIST '(= &H8000001D
            msg = "Invalid Call State List"
        Case LINEERR_INVALCARD '(= &H8000001E
            msg = "Invalid Card"
        Case LINEERR_INVALCOMPLETIONID '(= &H8000001F
            msg = "Invalid Completion ID"
        Case LINEERR_INVALCONFCALLHANDLE '(= &H80000020
            msg = "Invalid Conf Call Handle"
        Case LINEERR_INVALCONSULTCALLHANDLE '(= &H80000021
            msg = "Invalid Consult Call Handle"
        Case LINEERR_INVALCOUNTRYCODE '(= &H80000022
            msg = "Invalid Country Code"
        Case LINEERR_INVALDEVICECLASS '(= &H80000023
            msg = "Invalid Device Class"
        Case LINEERR_INVALDEVICEHANDLE '(= &H80000024
            msg = "Invalid Device Handle"
        Case LINEERR_INVALDIGITLIST '(= &H80000026
            msg = "Invalid Digit List"
        Case LINEERR_INVALDIGITMODE '(= &H80000027
            msg = "Invalid Digit Mode"
        Case LINEERR_INVALDIGITS '(= &H80000028
            msg = "Invalid Digits"
        Case LINEERR_INVALEXTVERSION '(= &H80000029
            msg = "Invalid Ext Version"
        Case LINEERR_INVALGROUPID '(= &H8000002A
            msg = "Invalid Group ID"
        Case LINEERR_INVALLINEHANDLE '(= &H8000002B
            msg = "Invalid Line Handle"
        Case LINEERR_INVALLINESTATE '(= &H8000002C
            msg = "Invalid Line State"
        Case LINEERR_INVALLOCATION '(= &H8000002D
            msg = "Invalid Location"
        Case LINEERR_INVALMEDIALIST '(= &H8000002E
            msg = "Invalid Media List"
        Case LINEERR_INVALMEDIAMODE '(= &H8000002F
            msg = "Invalid Media Mode"
        Case LINEERR_INVALMESSAGEID '(= &H80000030
            msg = "Invalid Message ID"
        Case LINEERR_INVALPARAM '(= &H80000032
            msg = "Invalid Param"
        Case LINEERR_INVALPARKID '(= &H80000033
            msg = "Invalid Park ID"
        Case LINEERR_INVALPARKMODE '(= &H80000034
            msg = "Invalid Park Mode"
        Case LINEERR_INVALPOINTER '(= &H80000035
            msg = "Invalid Pointer"
        Case LINEERR_INVALPRIVSELECT '(= &H80000036
            msg = "Invalid Priv Select"
        Case LINEERR_INVALRATE '(= &H80000037
            msg = "Invalid Rate"
        Case LINEERR_INVALREQUESTMODE '(= &H80000038
            msg = "Invalid Request Mode"
        Case LINEERR_INVALTERMINALID '(= &H80000039
            msg = "Invalid Terminal ID"
        Case LINEERR_INVALTERMINALMODE '(= &H8000003A
            msg = "Invalid Terminal Mode"
        Case LINEERR_INVALTIMEOUT '(= &H8000003B
            msg = "Invalid Time Out"
        Case LINEERR_INVALTONE '(= &H8000003C
            msg = "Invalid Tone"
        Case LINEERR_INVALTONELIST '(= &H8000003D
            msg = "Invalid Tone List"
        Case LINEERR_INVALTONEMODE '(= &H8000003E
            msg = "Invalid Tone Mode"
        Case LINEERR_INVALTRANSFERMODE '(= &H8000003F
            msg = "Invalid Transfer Mode"
        Case LINEERR_LINEMAPPERFAILED '(= &H80000040
            msg = "Line Mapper Failed"
        Case LINEERR_NOCONFERENCE '(= &H80000041
            msg = "No Conference"
        Case LINEERR_NODEVICE '(= &H80000042
            msg = "No Device"
        Case LINEERR_NODRIVER '(= &H80000043
            msg = "No Driver"
        Case LINEERR_NOMEM '(= &H80000044
            msg = "No Memory"
        Case LINEERR_NOREQUEST '(= &H80000045
            msg = "No Request"
        Case LINEERR_NOTOWNER '(= &H80000046
            msg = "Not Owner"
        Case LINEERR_NOTREGISTERED '(= &H80000047
            msg = "Not Registered"
        Case LINEERR_OPERATIONFAILED '(= &H80000048
            msg = "Operation Failed"
        Case LINEERR_OPERATIONUNAVAIL '(= &H80000049
            msg = "Operation Unavailable"
        Case LINEERR_RATEUNAVAIL '(= &H8000004A
            msg = "Rate Unavailable"
        Case LINEERR_RESOURCEUNAVAIL '(= &H8000004B
            msg = "Resource Unavailable"
        Case LINEERR_REQUESTOVERRUN '(= &H8000004C
            msg = "Request Overrun"
        Case LINEERR_STRUCTURETOOSMALL '(= &H8000004D
            msg = "Structure Too Small"
        Case LINEERR_TARGETNOTFOUND '(= &H8000004E
            msg = "Target Not found"
        Case LINEERR_TARGETSELF '(= &H8000004F
            msg = "Target Self"
        Case LINEERR_UNINITIALIZED '(= &H80000050
            msg = "Uninitialized"
        Case LINEERR_USERUSERINFOTOOBIG '(= &H80000051
            msg = "UserUser Info Too Big"
        Case LINEERR_REINIT '(= &H80000052
            msg = "Re-init"
        Case LINEERR_ADDRESSBLOCKED '(= &H80000053
            msg = "Address Blocked"
        Case LINEERR_BILLINGREJECTED '(= &H80000054
            msg = "Billing Rejected"
        Case LINEERR_INVALFEATURE '(= &H80000055
            msg = "Invalid Feature"
        Case LINEERR_NOMULTIPLEINSTANCE '(= &H80000056
            msg = "No Multiple Instance"
        Case Else
            msg = "Unknown Error" ' undefined
    End Select
    
    GetLineErrString = msg
End Function

Public Function GetLineStateString(ByVal state As Long) As String
    Dim msg As String

    Select Case state
        Case LINECALLSTATE_IDLE                       '&H1
            msg = "idle"
        
        Case LINECALLSTATE_OFFERING                   '&H2
            msg = "offering call"
        
        Case LINECALLSTATE_ACCEPTED                   '&H4
            msg = "accepted"
        
        Case LINECALLSTATE_DIALTONE                   '&H8
            msg = "dial-tone detected"
        
        Case LINECALLSTATE_DIALING                    '&H10
            msg = "dialing"
        
        Case LINECALLSTATE_RINGBACK                   '&H20
            msg = "ring-back detected"
        
        Case LINECALLSTATE_BUSY                       '&H40
            msg = "busy detected"
        
        Case LINECALLSTATE_SPECIALINFO                '&H80
            msg = "network error"
        
        Case LINECALLSTATE_CONNECTED                  '&H100
            msg = "connected"
        
        Case LINECALLSTATE_PROCEEDING                 '&H200
            msg = "proceeding"
        
        Case LINECALLSTATE_ONHOLD                     '&H400
            msg = "on hold"
        
        Case LINECALLSTATE_CONFERENCED                '&H800
            msg = "connected to conference"
        
        Case LINECALLSTATE_ONHOLDPENDCONF             '&H1000
            msg = "connecting to conference"
        
        Case LINECALLSTATE_ONHOLDPENDTRANSFER         '&H2000
            msg = "transferring"
        
        Case LINECALLSTATE_DISCONNECTED               '&H4000
            msg = "disconnected"
        
        Case LINECALLSTATE_UNKNOWN                    '&H8000
            msg = "unknown call state"
        
        Case Else
            msg = "unknown LineState"
        
    End Select
    
    GetLineStateString = msg
End Function
