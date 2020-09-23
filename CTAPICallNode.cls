VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAPICallNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TAPI Call Node (Caller, Called) (managed by the Call Obj)

'********************************************************
' Code Sample by Gregory Fox, Data Management Associates, Inc.
' Portions borrowed and modified from publically posted samples.
' Provided AS-IS.  Not tested in a production environment.
'********************************************************

Option Explicit


Private msID As String
Private msName As String
Private msStatus As String
Private mbInfoChanged As Boolean        'cleared, then set by UpdateCallInfo() when this info actually changes

Public Property Let ID(ByVal newValue As String)
    msID = newValue
End Property
Public Property Get ID() As String
    ID = msID
End Property

Public Property Let Name(ByVal newValue As String)
    msName = newValue
End Property
Public Property Get Name() As String
    Name = msName
End Property

Public Property Let Status(ByVal newValue As String)
    msStatus = newValue
End Property
Public Property Get Status() As String
    Status = msStatus
End Property

Public Property Let InfoChanged(ByVal newValue As Boolean)
    mbInfoChanged = newValue
End Property
Public Property Get InfoChanged() As Boolean
    InfoChanged = mbInfoChanged
End Property

