VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{A22136C1-299F-11D3-BA36-44455354616F}#4.0#0"; "PLAYCD2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "audio²"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   Picture         =   "audio].frx":0000
   ScaleHeight     =   3225
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB5Chat2.Chat Chat2 
      Left            =   2400
      Top             =   3120
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   360
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   960
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin CD_Player_2.CDPlayer CDPlayer1 
      Left            =   2400
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   1800
      Top             =   3120
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Multi As Boolean

Private Sub Form_Load()
ChatSend "<font face=""arial""><font color=#ffffff>• audio² • mage • " & Date & ""
Pause 0.2
ChatSend "<font face=""arial""><font color=#ffffff>• loaded by: " + LCase(GetUser) + ""
Chat1.ScanOn
List1.AddItem "               Mp3 Player Options       "
List1.AddItem ".dir  -  sets directory"
List1.AddItem ".exit  -  exits audio"
List1.AddItem ".play  -  plays mp3"
List1.AddItem ".stop  -  stops mp3 playing"
List1.AddItem ".total  -  scrolls total mp3s"
List1.AddItem ".ran  -  plays a radom mp3"
List1.AddItem ".mute  -  mutes mp3 player"
List1.AddItem ".unmute  -  unmutes mp3 player"
List1.AddItem ".pause  -  pauses mp3 player"
List1.AddItem ".unpause  -  unpauses mp3"
List1.AddItem ".vol  -  sets the volume"
List1.AddItem ".total - couts total mp3s"
List1.AddItem ".cd  -  enables cd player"
List1.AddItem ".mp3  -  enables mp3 player"
List1.AddItem "               Cd Player Options       "
List1.AddItem ".playcd  -  plays cd player"
List1.AddItem ".stopcd  -  stops cd player"
List1.AddItem ".cdopen  -  opens cd door"
List1.AddItem ".cdclose  -  closes cd door"
List1.AddItem ".prev  -  plays prev song"
List1.AddItem ".next  -  plays next song"
On Error Resume Next
File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.ini")
End Sub

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If Screen_Name$ = GetUser$ And InStr(What_Said, ".") = 1& Then
      lngSpace& = InStr(What_Said$, " ")
      If lngSpace& = 0& Then
         strCommand$ = What_Said$
      Else
         strCommand$ = Left(What_Said$, lngSpace& - 1&)
      End If
      Select Case strCommand$
      Case ".pause"
      MediaPlayer1.Pause
      ChatSend "<font face=""arial""><font color=#ffffff>• audio² - Mp3 Player Paused"
      
       Case ".unpause"
      MediaPlayer1.Play
      ChatSend "<font face=""arial""><font color=#ffffff>• audio² - Mp3 Player unPaused"

Case ".cancel"
If Multi = True Then
    List2.Clear
    Multi = False
    ChatSend "<font face=""arial""><font color=#ffffff>• audio²  [multiple strings cancelled]"
End If

Case ".exit"

ChatSend "<font face=""arial""><font color=#ffffff>• audio² • mage • " & Date & ""
Pause 0.2
ChatSend "<font face=""arial""><font color=#ffffff>• unloaded by: " + LCase(GetUser) + ""
End


Case ".vol"
If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
                Pause 0.5
            
            MediaPlayer1.Volume = (strArgument1$ * 25) - 2500
            ChatSend "<font face=""arial""><font color=#ffffff>• audio² - Mp3 Volume at " + strArgument1 + ""
            End If

Case ".ran"
On Error Resume Next
            File1.ListIndex = RandomNumber1(File1.ListCount)
            strArgument1$ = File1.Path & "\" & File1.filename
            On Error Resume Next
            MediaPlayer1.Open strArgument1$
            MediaPlayer1.Play
            ChatSend "<font face=""arial""><font color=#ffffff>• audio² - Playing [" & ReplaceString(ReplaceString(strArgument1$, File1.Path, ""), "\", "") & "]"

Case ".play"
If Multi <> True Then
    Call PlayIt(Right(What_Said, Len(What_Said) - 6))
    Exit Sub
End If
If Multi = True Then
    On Error GoTo kdmerr
    intt% = Right(What_Said, Len(What_Said) - 6)
    MediaPlayer1.Open File1.Path & "\" & List2.List(intt%)
    ChatSend "<font face=""arial""><font color=#ffffff>• audio²  now playing [" & ReplaceString(List2.List(intt%), ".mp3", "") & "]"
    Multi = False
    List2.Clear
    Exit Sub
kdmerr:
    If What_Said <> "<font face=""arial""><font color=#ffffff>• audio² [type ''.play'' and 0 - " & List2.ListCount - 1 & "or ''.cancel'']" Then ChatSend "<font color=#ffffff><font face=""arial"">×­› audio¹  [invaled number]"
End If


Case ".stop"
ChatSend "<font face=""arial""><font color=#ffffff>• audio² Mp3 Now Stoped"
MediaPlayer1.Stop


Case ".mute"
ChatSend "<font face=""arial""><font color=#ffffff>• audio² Mp3 Now Mute"
MediaPlayer1.Mute = True


Case ".unmute"
ChatSend "<font face=""arial""><font color=#ffffff>• audio² mp3 is now unmuted"
MediaPlayer1.Mute = False

Case ".dir"
strArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument1$
     Call WriteToINI("mp3dir", "dir", strArgument1$, App.Path & "\mp3.ini")
ChatSend "<font face=""arial""><font color=#ffffff>• audio² dir is " + strArgument1 + ""

Case ".total"
win$ = File1.ListCount
ChatSend "<font face=""arial""><font color=#ffffff>• audio² • total mp3s [" & win & "]"

Case ".cd"
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · cd player enabled "
Chat2.ScanOn
Chat1.ScanOff




Case ".mp3"
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · mp3 player enabled "
Chat1.ScanOn
Chat2.ScanOff
'
'
' CD CODING STARTS HERE
'
'
'






End Select
End If
End Sub

Private Sub Chat2_ChatMsg(Screen_Name As String, What_Said As String)
Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If Screen_Name$ = GetUser$ And InStr(What_Said, ".") = 1& Then
      lngSpace& = InStr(What_Said$, " ")
      If lngSpace& = 0& Then
         strCommand$ = What_Said$
      Else
         strCommand$ = Left(What_Said$, lngSpace& - 1&)
      End If
      Select Case strCommand$
Case ".playcd"
CDPlayer1.PlayCD
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · now playing track " & num & ""


Case ".cdopen"
CDPlayer1.OpenCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · now opening cd door"


Case ".cdclose"
CDPlayer1.CloseCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · now closing cd door"


Case ".stopcd"
Call CDPlayer1.StopCD
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · stoped playing track " & num & ""


Case ".next"
CDPlayer1.NextTrack
Pause 0.5
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · track " & num & ""


Case ".prev"
CDPlayer1.PreviousTrack
Pause 0.5
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · new track " & num & ""

Case ".playcd"

CDPlayer1.PlayCD
num$ = CDPlayer1.GetTrack
ChatSend "<font color=#ffffff><font face=""arial"">×­› audio¹ · now playing track " & num & ""


Case ".cdopen"
CDPlayer1.OpenCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · now opening cd door"


Case ".cdclose"
CDPlayer1.CloseCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · now closing cd door"


Case ".stopcd"
Call CDPlayer1.StopCD
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · stoped playing track " & num & ""


Case ".mp3"
ChatSend "<font face=""arial""><font color=#ffffff>• audio² · mp3 player enabled "
Chat1.ScanOn
Chat2.ScanOff

Case ".exit"

ChatSend "<font face=""arial""><font color=#ffffff>• audio² • mage • " & Date & ""
Pause 0.2
ChatSend "<font face=""arial""><font color=#ffffff>• unloaded by: " + LCase(GetUser) + ""
End

End Select
End If
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub




Public Sub PlayIt(sString As String)
    For i = 0 To File1.ListCount - 1
        If InStr(LCase(TrimSpaces(File1.List(i))), LCase(TrimSpaces(sString$))) Then
            List2.AddItem (File1.List(i))
        End If
    Next i
    If List2.ListCount = 1 Then
        MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        ChatSend "<font face=""arial""><font color=#ffffff>• audio² playing -[" & ReplaceString(List2.List(0), ".mp3", "") & "]"
        List2.Clear
        Exit Sub
    End If
    If List2.ListCount = 0 Then
        ChatSend "<font face=""arial""><font color=#ffffff>• audio² [no strings found]"
    End If
    If List2.ListCount > 1 Then
        ChatSend "<font face=""arial""><font color=#ffffff>• audio²  [multiple strings found]"
        For i = 0 To List2.ListCount - 1
            ChatSend "<font face=""arial""><font color=#ffffff>•  [" & i & ") " & List2.List(i) & "]"
            Pause 1
        Next i
        ChatSend "<font face=""arial""><font color=#ffffff>• audio²  [type ''.play'' and 0 - " & List2.ListCount - 1 & " or ''.cancel'']"
        Multi = True
    End If
End Sub



