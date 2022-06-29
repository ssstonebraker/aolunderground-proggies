VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{A22136C1-299F-11D3-BA36-44455354616F}#4.0#0"; "PLAYCD2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "audio].frx":0000
   ScaleHeight     =   1110
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "dir:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      TabIndex        =   13
      Top             =   3000
      Width           =   3375
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   2400
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4680
      Width           =   2055
   End
   Begin VB5Chat2.Chat Chat2 
      Left            =   360
      Top             =   5640
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3360
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   900
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin CD_Player_2.CDPlayer CDPlayer1 
      Left            =   840
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   480
      Top             =   5640
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "  ^"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "  v"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "  x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "  ^"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "     v"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Label3"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "  -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
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
      DisplayForeColor=   0
      DisplayMode     =   0
      DisplaySize     =   0
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
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Multi As Boolean

Private Sub Command1_Click()
strArgument1$ = Text2.text
On Error Resume Next
File1.Path = strArgument1$
     Call WriteToINI("mp3dir", "dir", strArgument1$, App.Path & "\mp3.ini")
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio dir is " + strArgument1 + ""
End Sub

Private Sub File1_Click()
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Stop
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio Mp3 Now Stopped"
Pause 1
MediaPlayer1.Open File1.Path & "\" & File1.FileName
        
    ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio playing -[" & ReplaceString(File1.FileName, ".mp3", "") & "]"
Exit Sub
End If
MediaPlayer1.Open File1.Path & "\" & File1.FileName
        ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio playing -[" & ReplaceString(File1.FileName, ".mp3", "") & "]"
End Sub

Private Sub Form_Load()
FormOnTop Me
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • cash • " & Date & ""
Pause 0.8
ChatSend "<font face=""arial""><font color=#ffffff>• so easy to use, no wonder its # 1 •"
Pause 0.8
ChatSend "<font face=""arial""><font color=#ffffff>• loaded by: " + LCase(GetUser) + ""
Chat1.ScanOn
List1.AddItem "                    ez - audio commands       "
List1.AddItem "-------------------------------------------------------------------------------"
List1.AddItem ".adv  -  advertise"
List1.AddItem ".coms  -  views commands"
List1.AddItem ".dir  -  sets directory"
List1.AddItem ".exit  -  exits audio"
List1.AddItem ".hide  -  hides mp3 player"
List1.AddItem ".length  -  scrolls song length"
List1.AddItem ".loop  -  loops mp3"
List1.AddItem ".mute  -  mutes mp3 player"
List1.AddItem ".pause  -  pauses mp3 player"
List1.AddItem ".play  -  plays mp3"
List1.AddItem ".pr  -  enters private room"
List1.AddItem ".ran  -  plays a random mp3"
List1.AddItem ".refdir  -  refreshes mp3's"
List1.AddItem ".stop  -  stops mp3 playing"
List1.AddItem ".total  -  scrolls total mp3s"
List1.AddItem ".viewmp3s  -  views mp3's"
List1.AddItem ".vol  -  sets the volume"
List1.AddItem ".xcoms  -  hides commands"
List1.AddItem ".xhide  -  unhides mp3 player"
List1.AddItem ".xloop  -  unloops mp3"
List1.AddItem ".xmute  -  unmutes mp3 player"
List1.AddItem ".xpause  -  unpauses mp3"
List1.AddItem ".xviewmp3s  -  hides mp3's"
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
      ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Player Paused"
      
       Case ".xpause"
      MediaPlayer1.Play
      ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Player unPaused"

Case ".loop"
MediaPlayer1.PlayCount = 0
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Player · looping now enabled"

Case ".adv"
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • cash •"
Pause 0.8
ChatSend "<font face=""arial""><font color=#ffffff>• so easy to use, no wonder its # 1 •"

Case ".pr"
PrivateRoom (Right(What_Said, Len(What_Said) - 4))

Case ".hide"
Form1.Visible = False

Case ".xhide"
Form1.Visible = True

Case ".viewmp3s"
Form1.Height = "3045"
Form1.Visible = True
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • viewing mp3's"
Case ".xviewmp3s"
Form1.Visible = False
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • hiding mp3's"

Case ".xcoms"
Form1.Visible = False
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • hiding commands"

Case ".xloop"
MediaPlayer1.PlayCount = 1
ChatSend "</b><font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Player · looping now disabled"
Case ".refdir"
File1.Refresh
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Player · mp3 list refreshed"
Case ".length"
Dim lngSeconds As Long
Dim lngMinutes As Long
Dim strSeconds As String
Dim strMinutes As String
Dim lngPos As Long
    strSeconds$ = MediaPlayer1.Duration
    If InStr(strSeconds$, ".") <> 0& Then
        strSeconds$ = Right(strSeconds$, Len(strSeconds$) - InStr(strSeconds$, "*"))
    End If
    lngSeconds& = strSeconds$
    Do Until lngSeconds& <= 59
        DoEvents
        lngSeconds& = lngSeconds& - 60
        lngMinutes& = lngMinutes& + 1
    Loop
    strSeconds$ = Format$(lngSeconds&, "00")
    strMinutes$ = Format$(lngMinutes&, "00")
    SongLength$ = strMinutes$ & ":" & strSeconds$
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Player Song length : " + SongLength + ""

Case ".exit"
If MediaPlayer1.PlayState = mpPlaying Then
MsgBox "Stop the current mp3", vbExclamation, "ez - audio"
Exit Sub
End If
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • cash • " & Date & ""
Pause 0.8
ChatSend "<font face=""arial""><font color=#ffffff>• unloaded by: " + LCase(GetUser) + ""
End

Case ".coms"
If Form1.Visible = True Then
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • viewing commands"
Exit Sub
End If
Form1.Visible = True
Form1.Height = "1200"
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • viewing commands"

Case ".vol"
If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
                Pause 0.5
            If strArgument1 > 100 Then
            ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Volume too high"
            Else
            Pause 0.5
            MediaPlayer1.Volume = (strArgument1$ * 25) - 2500
            ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio - Mp3 Volume at " + strArgument1 + ""
            End If
            End If

Case ".ran"
On Error Resume Next
            File1.ListIndex = RandomNumber1(File1.ListCount)
            strArgument1$ = File1.Path & "\" & File1.FileName
            On Error Resume Next
            MediaPlayer1.Open strArgument1$
            MediaPlayer1.Play
            

Case ".play"
On Error Resume Next
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Stop
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio Mp3 Now Stopped"
Pause 1
End If
If Multi <> True Then
    Call PlayIt(Right(What_Said, Len(What_Said) - 6))
    Exit Sub
End If
If Multi = True Then
    On Error GoTo kdmerr
    intt% = Right(What_Said, Len(What_Said) - 6)
    MediaPlayer1.Open File1.Path & "\" & List2.List(intt%)
    Text1.text = "• ez - audio  now playing [" & ReplaceString(List2.List(intt%), ".mp3", "") & "]"
    If Text1.text = "• ez - audio  now playing []" Then
    ChatSend "<font color=#ffffff><font face=""arial"">• ez - audio [invaled number]"
    Exit Sub
    End If
    ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio  now playing [" & ReplaceString(List2.List(intt%), ".mp3", "") & "]"
    Multi = False
    List2.Clear
    Exit Sub
kdmerr:
    If What_Said <> "<font face=""arial""><font color=#ffffff>• ez - audio [type ''.play'' and 0 - " & List2.ListCount - 1 & "or ''.cancel'']" Then ChatSend "<font color=#ffffff><font face=""arial"">×­› audio¹  [invaled number]"
End If





Case ".stop"
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • Mp3 Now Stopped"
MediaPlayer1.Stop


Case ".mute"
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio Mp3 Now Mute"
MediaPlayer1.Mute = True


Case ".xmute"
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio mp3 is now unmuted"
MediaPlayer1.Mute = False

Case ".dir"
strArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument1$
     Call WriteToINI("mp3dir", "dir", strArgument1$, App.Path & "\mp3.ini")
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio dir is " + strArgument1 + ""

Case ".total"
win$ = File1.ListCount
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • total mp3s [" & win & "]"

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
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · now playing track " & num & ""


Case ".cdopen"
CDPlayer1.OpenCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · now opening cd door"


Case ".cdclose"
CDPlayer1.CloseCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · now closing cd door"


Case ".stopcd"
Call CDPlayer1.StopCD
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · stoped playing track " & num & ""


Case ".next"
CDPlayer1.NextTrack
Pause 0.5
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · track " & num & ""


Case ".prev"
CDPlayer1.PreviousTrack
Pause 0.5
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · new track " & num & ""

Case ".playcd"

CDPlayer1.PlayCD
num$ = CDPlayer1.GetTrack
ChatSend "<font color=#ffffff><font face=""arial"">×­› audio¹ · now playing track " & num & ""


Case ".cdopen"
CDPlayer1.OpenCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · now opening cd door"


Case ".cdclose"
CDPlayer1.CloseCDDoor
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · now closing cd door"


Case ".stopcd"
Call CDPlayer1.StopCD
num$ = CDPlayer1.GetTrack
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · stoped playing track " & num & ""


Case ".mp3"
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio · mp3 player enabled "
Chat1.ScanOn
Chat2.ScanOff

Case ".exit"

ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • cash • " & Date & ""
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
        ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio playing -[" & ReplaceString(List2.List(0), ".mp3", "") & "]"
        List2.Clear
        Exit Sub
    End If
    If List2.ListCount = 0 Then
        ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio [no strings found]"
    End If
    If List2.ListCount > 1 Then
        ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio  [multiple strings found]"
        For i = 0 To List2.ListCount - 1
            ChatSend "<font face=""arial""><font color=#ffffff>•  [" & i & ") " & List2.List(i) & "]"
            Pause 1
        Next i
        ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio  [type ''.play'' and 0 - " & List2.ListCount - 1 & " or ''.cancel'']"
        Multi = True
    End If
End Sub

Public Sub PlayIt2(sString As String)
        MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio playing -[" & ReplaceString(List2.List(0), ".mp3", "") & "]"
        List2.Clear
        
End Sub



Private Sub Label1_Click()
Form1.WindowState = "1"
End Sub

Private Sub Label5_Click()
Form1.Height = "3045"
End Sub

Private Sub Label6_Click()
Form1.Height = "1200"
Label5.Visible = True
End Sub

Private Sub Label7_Click()
If MediaPlayer1.PlayState = mpPlaying Then
MsgBox "Stop the current mp3", vbExclamation, "ez - audio"
Exit Sub
End If
ChatSend "<font face=""arial""><font color=#ffffff>• ez - audio • cash • " & Date & ""
Pause 0.8
ChatSend "<font face=""arial""><font color=#ffffff>• unloaded by: " + LCase(GetUser) + ""
End
End Sub

Private Sub Label8_Click()
Form1.Height = "3435"
End Sub

Private Sub Label9_Click()
Form1.Height = "3045"
End Sub

