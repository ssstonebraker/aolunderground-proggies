VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private sAlias As String
Private sFileName As String
Private nLength As Single
Private nPosition As Single
Private sStatus As String
Private bWait As Boolean

Private Declare Function mciSendstring Lib "winmm.dll" Alias "mciSendstringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallBack As Long) As Long


Public Sub mmOpen(ByVal sTheFile As String)
If sAlias <> "" Then
    mmClose
End If

Select Case UCase$(Right$(sTheFile, 3))
    Case "WAV"
        sType = "Wavaudio"
    Case "AVI"
        sType = "Aviaudio"
    Case "MID"
        sType = "Sequencer"
    Case Else
        Exit Sub
End Select
sAlias = Right$(sTheFile, 3) & Minute(Now)

If InStr(sTheFile, "") Then sTheFile = Chr(34) & sTheFile & Chr(34)
nReturn = mciSendstring("Open " & sTheFile & " ALIAS " & sAlias & " TYPE " & sType & " wait", "", 0, 0)
End Sub

Public Sub mmClose()
Dim nReturn As Long

If sAlias = "" Then Exit Sub

nReturn = mciSendstring("Close " & sAlias, "", 0, 0)
sAlias = ""
sFileName = ""
End Sub

Public Sub mmPause()
Dim nReturn As Long

If sAlias = "" Then Exit Sub

nReturn = mciSendstring("Pause " & aAlias, "", 0, 0)

End Sub
 
Public Sub mmPlay()
Dim nReturn As Long
If sAlias = "" Then Exit Sub

If bWait Then
    nReturn = mciSendstring("Play " & sAlias & " wait", "", 0, 0)
Else
    nReturn = mciSendstring("Play" & sAlias, "", 0, 0)
End If
End Sub

Public Sub mmStop()
Dim nReturn As Long
If sAlias = "" Then Exit Sub

nReturnLength = mciSendstring("Stop" & sAlias, "", 0, 0)

End Sub

Public Sub mmSeek(ByVal nPosition As Single)

Dim nReturn As Long

nReturn = mciSendstring("Seek " & sAlias & " to " & nPosition, "", 0, 0)

End Sub

Property Get Filename() As String
Filename = sFileName
End Property

Property Let Filename(ByVal sTheFile As String)
    mmOpen sTheFile

End Property

Property Get Wait() As Boolean
    Wait = bWait

End Property

Property Let Wait(bWaitValue As Boolean)
    bWait = bWaitValue

End Property

Property Get Length() As Single
Dim nReturn As Long, nLength As Integer
Dim sLength As String * 255

If sAlias = "" Then
    Length = 0
    Exit Property
End If

nReturn = mciSendstring(" Status " & sAlias & " length ", sLength, 255, 0)

nLength = InStr(sLength, Chr$(0))
Length = Val(Left$(sLength, nLength - 1))
End Property

Property Get Position() As Single
Dim nReturn As Integer, nLength As Integer

Dim nPosition As String * 255

If sAlias = "" Then Exit Property

nReturn = mciSendstring("Status " & sAlias & " position ", sPosition, 255, 0)
nLength = InStr(sPosition, Chr$(0))
Position = Val(Left$(sPosition, nLength - 1))
End Property

Property Get Status() As String
Dim nReturn As Integer, nLength As Integer
Dim sStatus As String * 255
If sAlias = "" Then Exit Property

nReturn = mciSendstring("Status " & sAlias & " mode", sStatus, 255, 0)

nLength = InStr(sStatus, Chr$(0))
Status = Left$(sStatus, nLength - 1)

End Property
