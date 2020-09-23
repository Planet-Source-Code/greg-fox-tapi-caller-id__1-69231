VERSION 5.00
Begin VB.Form frmTAPITest 
   Caption         =   " TAPI Caller ID Test Window"
   ClientHeight    =   5625
   ClientLeft      =   630
   ClientTop       =   1710
   ClientWidth     =   6705
   Icon            =   "frmTAPITest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCID 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "txtCID"
      Top             =   4950
      Width           =   6405
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   4230
      TabIndex        =   9
      Top             =   1860
      Width           =   2310
   End
   Begin VB.PictureBox picIcon 
      Height          =   660
      Left            =   4980
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   6
      Top             =   135
      Width           =   660
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   525
      Left            =   4230
      TabIndex        =   5
      Top             =   1170
      Width           =   2310
   End
   Begin VB.CommandButton cmdDialProps 
      Caption         =   "Dialing Properties..."
      Height          =   525
      Left            =   4230
      TabIndex        =   4
      Top             =   3720
      Width           =   2310
   End
   Begin VB.CommandButton cmdConfigDlg 
      Caption         =   "Line Config Dialog..."
      Height          =   525
      Left            =   4230
      TabIndex        =   3
      Top             =   2970
      Width           =   2310
   End
   Begin VB.ListBox lstLineInfo 
      Height          =   3765
      Left            =   135
      TabIndex        =   2
      Top             =   510
      Width           =   3885
   End
   Begin VB.ComboBox cboLineSel 
      Height          =   315
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label labCID 
      Caption         =   "Caller ID Info:"
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   4710
      Width           =   3885
   End
   Begin VB.Label labAPIVersion 
      Caption         =   "labAPIVersion"
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   4350
      Width           =   3885
   End
   Begin VB.Label lblIcon 
      Caption         =   "Icon:"
      Height          =   195
      Left            =   4440
      TabIndex        =   7
      Top             =   210
      Width           =   510
   End
   Begin VB.Label lblLineSel 
      Caption         =   "Select TAPI Line:"
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   210
      Width           =   1950
   End
End
Attribute VB_Name = "frmTAPITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' TAPI Test Form

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

'****************************************************************
' Thanks to two people, in particular,
' who posted sample code for review...
'       1. Ray Mercer
'       2. Brian Yule
'****************************************************************


Option Explicit



Private WithEvents moTAPICID As CTAPICID
Attribute moTAPICID.VB_VarHelpID = -1
Private moDevList As CTAPIDeviceList


Private Sub Form_Load()

    Dim oDevDesc As CTAPIDeviceDesc
    Dim i As Long
    
    Set moDevList = New CTAPIDeviceList          'Initializes the TAPI DLL
    
    If moDevList.TAPIDeviceCount > 0 Then
        For i = 0 To moDevList.TAPIDeviceCount - 1
            Set oDevDesc = moDevList.TAPIDevice(i)
            cboLineSel.AddItem oDevDesc.LineName
        Next
        
        cboLineSel.ListIndex = 0    '(triggers the click event)
        labAPIVersion.Caption = "TAPI v" & moDevList.TAPIVersion
        
    Else
        lstLineInfo.AddItem moDevList.ErrorString
    End If
    
    Set oDevDesc = Nothing
    Set moTAPICID = New CTAPICID
    
    txtCID.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (moTAPICID Is Nothing) Then
        moTAPICID.CloseCIDLine
        Set moTAPICID = Nothing
    End If
    
    Set moDevList = Nothing
    
End Sub

Private Sub cboLineSel_Click()

    Dim oDevDesc As CTAPIDeviceDesc
    Dim lIndex As Long
    Dim bDialEnable As Boolean

    picIcon.AutoRedraw = True

    lstLineInfo.Clear
    If cboLineSel.List(cboLineSel.ListIndex) <> "" Then
    
        lIndex = cboLineSel.ListIndex
        Set oDevDesc = moDevList.TAPIDevice(lIndex)
        
        If Not (oDevDesc Is Nothing) Then
            'this section just prints out a lot of info about the selected Line
            lstLineInfo.AddItem "TAPI LINE: #" & oDevDesc.DeviceIndex
            lstLineInfo.AddItem "TAPI LINE NAME: " & oDevDesc.LineName
            lstLineInfo.AddItem "TAPI PROVIDER INFO: " & oDevDesc.ProviderInfo
            lstLineInfo.AddItem "TAPI SWITCH INFO: " & oDevDesc.SwitchInfo
            lstLineInfo.AddItem "Permanent Line ID: " & oDevDesc.PermanentLineID
            Select Case oDevDesc.StringFormat
                Case STRINGFORMAT_ASCII
                    lstLineInfo.AddItem "String Format: STRINGFORMAT_ASCII"
                Case STRINGFORMAT_DBCS
                    lstLineInfo.AddItem "String Format: STRINGFORMAT_DBCS"
                Case STRINGFORMAT_UNICODE
                    lstLineInfo.AddItem "String Format: STRINGFORMAT_UNICODE"
                Case STRINGFORMAT_BINARY
                    lstLineInfo.AddItem "String Format: STRINGFORMAT_BINARY"
                Case Else
            End Select
            lstLineInfo.AddItem "Number of addresses associated with this line: " & oDevDesc.NumAddresses
            lstLineInfo.AddItem "Max data rate: " & oDevDesc.MaxDataRate
            lstLineInfo.AddItem "Bearer Modes supported:"
            If LINEBEARERMODE_VOICE And oDevDesc.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_VOICE"
            If LINEBEARERMODE_SPEECH And oDevDesc.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_SPEECH"
            If LINEBEARERMODE_DATA And oDevDesc.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_DATA"
            If LINEBEARERMODE_ALTSPEECHDATA And oDevDesc.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_ALTSPEECHDATA"
            If LINEBEARERMODE_MULTIUSE And oDevDesc.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_MULTIUSE"
            If LINEBEARERMODE_NONCALLSIGNALING And oDevDesc.BearerModes Then lstLineInfo.AddItem vbTab & "LINEBEARERMODE_NONCALLSIGNALING"
            lstLineInfo.AddItem "Address Modes supported:"
            If oDevDesc.AddressModes And LINEADDRESSMODE_ADDRESSID Then lstLineInfo.AddItem vbTab & "LINEADDRESSMODE_ADDRESSID"
            If oDevDesc.AddressModes And LINEADDRESSMODE_DIALABLEADDR Then lstLineInfo.AddItem vbTab & "LINEADDRESSMODE_DIALABLEADDR"
            lstLineInfo.AddItem "Media Modes supported:"
            If LINEMEDIAMODE_ADSI And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_ADSI"
            If LINEMEDIAMODE_AUTOMATEDVOICE And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_AUTOMATEDVOICE"
            If LINEMEDIAMODE_DATAMODEM And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_DATAMODEM"
            If LINEMEDIAMODE_DIGITALDATA And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_DIGITALDATA"
            If LINEMEDIAMODE_G3FAX And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_G3FAX"
            If LINEMEDIAMODE_G4FAX And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_G4FAX"
            If LINEMEDIAMODE_INTERACTIVEVOICE And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_INTERACTIVEVOICE"
            If LINEMEDIAMODE_MIXED And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_MIXED"
            If LINEMEDIAMODE_TDD And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_TDD"
            If LINEMEDIAMODE_TELETEX And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_TELETEX"
            If LINEMEDIAMODE_TELEX And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_TELEX"
            If LINEMEDIAMODE_UNKNOWN And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_UNKNOWN"
            If LINEMEDIAMODE_VIDEOTEX And oDevDesc.MediaModes Then lstLineInfo.AddItem vbTab & "LINEMEDIAMODE_VIDEOTEX"
            lstLineInfo.AddItem "Line Tone Generation supported: " & CBool(oDevDesc.GenerateToneMaxNumFreq)
            If CBool(oDevDesc.GenerateToneMaxNumFreq) Then 'show if tone generation is supported
                If LINETONEMODE_BEEP And oDevDesc.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_BEEP"
                If LINETONEMODE_BILLING And oDevDesc.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_BILLING"
                If LINETONEMODE_BUSY And oDevDesc.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_BUSY"
                If LINETONEMODE_CUSTOM And oDevDesc.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_CUSTOM"
                If LINETONEMODE_RINGBACK And oDevDesc.GenerateToneModes Then lstLineInfo.AddItem vbTab & "LINETONEMODE_RINGBACK"
            End If
            lstLineInfo.AddItem "Number of terminals for this line: " & oDevDesc.NumTerminals
            lstLineInfo.AddItem "Device API Version: " & oDevDesc.DeviceAPIVersion
            
            If oDevDesc.DeviceSupportsVoiceCalls Then bDialEnable = True
                
            picIcon.Picture = LoadPicture()      'for some reason PaintDeviceIcon only works if this is called first
            oDevDesc.PaintDeviceIcon picIcon.hDC, 4, 4
            
        Else
            lstLineInfo.AddItem "<Unable to load selected TAPI Device>"
        End If
    Else
        lstLineInfo.AddItem "<No Valid TAPI Line Selected>"
    End If
    
    Set oDevDesc = Nothing

End Sub


Private Sub cmdListen_Click()

    moTAPICID.OpenCIDLine cboLineSel.ListIndex
    
End Sub

Private Sub cmdClose_Click()

    moTAPICID.CloseCIDLine
    
End Sub

Private Sub cmdConfigDlg_Click()
    
    On Error Resume Next
    moDevList.TAPIDevice(cboLineSel.ListIndex).OpenConfigDialog Me.hWnd
    
End Sub

Private Sub cmdDialProps_Click()

    On Error Resume Next
    moDevList.TAPIDevice(cboLineSel.ListIndex).OpenDialingPropDialog Me.hWnd, ""

End Sub

Private Sub moTAPICID_CallEnded(ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String)

    txtCID.Text = ""
    
End Sub

Private Sub moTAPICID_IncomingCIDMsg(ByVal lTAPIDeviceIndex As Long, _
    ByVal sTAPIDeviceName As String, _
    ByVal sCallerID As String, _
    ByVal sCallerName As String)

    txtCID.Text = "Last Incoming Call From:  " & sCallerName & ", Phone: " & sCallerID
    
End Sub



