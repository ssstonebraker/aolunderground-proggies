VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{DE8D4E3E-DD62-11D2-821F-444553540001}#1.0#0"; "CHATSCAN³.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   9210
   ClientTop       =   3015
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "mp3player.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   1680
      Top             =   2760
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin chatscan³.Chat Chat1 
      Left            =   3600
      Top             =   4080
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   120
      Pattern         =   "*.mp3*"
      TabIndex        =   1
      Top             =   2280
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "intrex³ coded by [ • ip³ • ]"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[ • ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label lblTotalTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[ • ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label lblElapsedTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   4575
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   4575
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   6015
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Beta"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Multi As Boolean
Private Sub Form_Load()
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • coded by [ • ip³ • ]"
Pause 0.2
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • status [ • loaded • ] • " & Time
Chat1.ScanOn
List1.AddItem "           mp3 commands"
List1.AddItem ".dir  -  sets directory"
List1.AddItem ".play  -  plays mp3"
List1.AddItem ".stop  -  stops mp3 playing"
List1.AddItem ".next  -  plays next mp3"
List1.AddItem ".ran  -  plays mp3s randomly"
List1.AddItem ".mute  -  mutes mp3 player"
List1.AddItem ".unmute  -  unmutes mp3 player"
List1.AddItem ".pause  -  pauses mp3 player"
List1.AddItem ".unpause  -  unpauses mp3"
List1.AddItem ".vol  -  sets the volume"
List1.AddItem ".total  -  scrolls total mp3s"
List1.AddItem "           other commands"
List1.AddItem ".pr  -  enter's a private room"
List1.AddItem ".x  -  ignores fellow chat member"
List1.AddItem ".ion  -  turns ims on"
List1.AddItem ".ioff -  turns ims off"
List1.AddItem ".get  -  get member profile"
List1.AddItem ".rt  -  scrolls total people in room"
List1.AddItem "kw  -  keyword"
List1.AddItem ".exit  -  exits intrex³"
On Error Resume Next
File1.Path = GetFromINI("intrex³mp3dir", "dir", App.Path & "\mp3.ini")
End Sub

Private Sub Form_Resize()
FormOnTop Me
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
      ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 paused"
      
       Case ".unpause"
      MediaPlayer1.Play
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 unpaused"

Case ".cancel"
If Multi = True Then
    List2.Clear
    Multi = False
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • cancelled"
End If

Case ".vol"
If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
                Pause 0.5
            
            MediaPlayer1.Volume = (strArgument1$ * 25) - 2500
            ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 volume[" + strArgument1 + "]"
            End If

Case ".ran"
    On Error Resume Next
    File1.ListIndex = RandomNumber1(File1.ListCount)
    strArgument1$ = File1.Path & "\" & File1.FileName
    On Error Resume Next
    MediaPlayer1.Open strArgument1$
    MediaPlayer1.Play
    Option1.Value = True
    ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • random [ " & File1.ListIndex & " ] • [ " & ReplaceString(ReplaceString(strArgument1$, File1.Path, ""), "\", "") & " ]"

    
Case ".play"
If Multi <> True Then
    Call PlayIt(Right(What_Said, Len(What_Said) - 6))
    Exit Sub
End If
If Multi = True Then
    On Error GoTo kdmerr
    intt% = Right(What_Said, Len(What_Said) - 6)
    MediaPlayer1.Open File1.Path & "\" & List2.List(intt%)
    ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • playing [ " & ReplaceString(List2.List(intt%), ".mp3", "") & " ]"
    Multi = False
    List2.Clear
    Exit Sub
kdmerr:
    If What_Said <> "<font face=wingdings 2><font color=#ffffff>b<font face=""arial""> <b>zip³</b> •[type *.play* and 0 - " & List2.ListCount - 1 & " or *.cancel*]" Then ChatSend "<font face=wingdings 2><font color=#ffffff>b<font face=""arial""> <b>zip³</b> •[invaled number]"
End If


Case ".stop"
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 stopped"
If Option1.Value = True Then
    Option1.Value = False
End If
MediaPlayer1.Stop

Case ".next"
    MediaPlayer1.Stop


Case ".mute"
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 muted"
MediaPlayer1.Mute = True


Case ".unmute"
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 unmuted"
MediaPlayer1.Mute = False

Case ".dir"
strArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument1$
     Call WriteToINI("intrex³mp3dir", "dir", strArgument1$, App.Path & "\mp3.ini")
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • dir[ " + strArgument1 + " ]"

Case ".total"
win$ = File1.ListCount
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • mp3 total [ " & win & " ]"

Case ".exit"
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • coded by [ • ip³ • ]"
Pause 0.2
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • status [ • unloaded • ] • " & Time
End

Case ".time"
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • " & Time

Case ".pr"
strArgument1$ = Mid(What_Said, 5)
PrivateRoom (strArgument1$)

Case ".x"
strArgument1$ = Mid(What_Said, 4)
ChatIgnoreByName (strArgument1$)
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • " & strArgument1$ & " ignored"

Case ".ion"
IMsOn
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • ims on"

Case ".ioff"
IMsOff
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • ims off"

Case ".get"
strArgument1$ = Mid(What_Said, 6)
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • task complete"
ProfileGet (strArgument1$)

Case ".rt"
strArgument1$ = Mid(What_Said, 5)
ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • people in room [ " & RoomCount & " ]"

Case ".kw"
strArgument1$ = Mid(What_Said, 5)
Keyword (strArgument1$)

End Select
End If
End Sub

Public Sub PlayIt(sString As String)
Dim Title As String
    lblName.Caption = File1.FileName
    For i = 0 To File1.ListCount - 1
        If InStr(LCase(TrimSpaces(File1.List(i))), LCase(TrimSpaces(sString$))) Then
            List2.AddItem (File1.List(i))
        Title = File1.FileName
        End If
    Next i
    If List2.ListCount = 1 Then
        MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • playing [ " & ReplaceString(List2.List(0), ".mp3", "") & " ]"
        lblName.Caption = MediaPlayer1.FileName
        List2.Clear
        Exit Sub
    End If
    If List2.ListCount = 0 Then
        ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • no strings found"
    End If
    If List2.ListCount > 1 Then
        ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • multiple strings found"
        For i = 0 To List2.ListCount - 1
            ChatSend "<font face=wingdings 2><font color=#ffffff>b<font face=""arial""> • [" & i & "•" & List2.List(i) & "]"
            Pause 1
        Next i
        ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • type *.play* and 0 - " & List2.ListCount - 1 & " or *.cancel*"
        Multi = True
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub MediaPlayer1_OpenStateChange(ByVal OldState As Long, ByVal NewState As Long)
Dim X As String
Min = MediaPlayer1.Duration \ 60
Sec = MediaPlayer1.Duration - (Min * 60)
lblTotalTime.Caption = "Total Time: " & Format(Min, "0#") _
    & ":" & Format(Sec, "0#") 'format time to 00:00
lblName.Caption = File1.FileName
FileOpen = CBool(NewState)
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
If Option1.Value = True Then
    If MediaPlayer1.PlayState = mpStopped Then
        Label1.Caption = True
        On Error Resume Next
            File1.ListIndex = RandomNumber1(File1.ListCount)
            strArgument1$ = File1.Path & "\" & File1.FileName
            On Error Resume Next
            MediaPlayer1.Open strArgument1$
            MediaPlayer1.Play
            ChatSend "<font face=""wingdings 2""><font color=#0000CC>c<font face=""haettenschweiler""> • intrex³ • random [ " & File1.ListIndex & " ] • [ " & ReplaceString(ReplaceString(strArgument1$, File1.Path, ""), "\", "") & " ]"
            lblName.Caption = File1.FileName
    End If
End If
End Sub

Private Sub Timer1_Timer()
Min = MediaPlayer1.CurrentPosition \ 60
Sec = MediaPlayer1.CurrentPosition - (Min * 60)
If Min > 0 Or Sec > 0 Then
    lblElapsedTime.Caption = Format(Min, "0#") _
        & ":" & Format(Sec, "0#")
Else
    lblElapsedTime.Caption = "00:00"
End If
End Sub






