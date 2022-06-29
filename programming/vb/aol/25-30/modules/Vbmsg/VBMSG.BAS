Option Explicit

' Declares for VB Messenger v1.0c API
' Copyright © 1993, 1994, 1995 JOSWare, Inc.
' All rights reserved.

Type MODEL
  usVersion         As Integer      'VB version used by control
  fl                As Long         'Bitfield structure
  pctlproc          As Long         'The control proc.
  fsClassStyle      As Integer      'Window class style
  flWndStyle        As Long         'Default window style
  cbCtlExtra        As Integer      '# bytes alloc'd for HCTL structure
  idBmpPalette      As Integer      'BITMAP id for tool palette
  npszDefCtlName    As Integer      'Default control name prefix (near ptr)
  npszClassName     As Integer      'Visual Basic class name (near ptr)
  npszParentClassName As Integer    'Parent window class if subclassed (near ptr)
  npproplist        As Integer      'Property list (near ptr)
  npeventlist       As Integer      'Event list (near ptr)
  nDefProp          As String * 1   'Index of default property
  nDefEvent         As String * 1   'Index of default event
  nValueProp        As String * 1   'Index of control value property
  usCtlVersion      As Integer
End Type

' Get specific address of various variable types
Declare Function ptGetVariableAddress Lib "VBMSG.VBX" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMSG.VBX" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMSG.VBX" Alias "ptGetVariableAddress" (ByVal S As String) As Long
Declare Function ptGetLongAddress Lib "VBMSG.VBX" Alias "ptGetVariableAddress" (L As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMSG.VBX" Alias "ptGetVariableAddress" (I As Integer) As Long

' Get various values of data at specific addresses
Declare Function ptGetIntegerFromAddress Lib "VBMSG.VBX" (ByVal I As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMSG.VBX" (ByVal L As Long) As Long
Declare Function ptGetStringFromAddress Lib "VBMSG.VBX" (ByVal lAddress As Long, ByVal cbBytes As Integer) As String
Declare Sub ptGetTypeFromAddress Lib "VBMSG.VBX" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)

'Set address areas to variable values
Declare Sub ptCopyTypeToAddress Lib "VBMSG.VBX" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)

' Data Type Manipulation Functions
Declare Function ptMakelParam Lib "VBMSG.VBX" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMSG.VBX" (ByVal lParam As Long) As Integer
Declare Function ptHiWord Lib "VBMSG.VBX" (ByVal lParam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMSG.VBX" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMSG.VBX" (ByVal ushortVal As Integer) As Long

' Get the literal text for a message
Declare Function ptMessagetoText Lib "VBMSG.VBX" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String

' Visual Basic API Translator Functions
Declare Function ptRecreateControlHwnd Lib "VBMSG.VBX" (ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMSG.VBX" (ctl As Control, lpm As MODEL) As Long
Declare Sub ptSetControlModel Lib "VBMSG.VBX" (ctl As Control, lpm As MODEL)
Declare Function ptGetControlName Lib "VBMSG.VBX" (ctl As Control) As String

