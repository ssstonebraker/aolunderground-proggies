VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "[V]atriX: REBIRTH by Freeza & FreAka"
   ClientHeight    =   4170
   ClientLeft      =   4650
   ClientTop       =   4740
   ClientWidth     =   4125
   ForeColor       =   &H00000000&
   Height          =   4575
   Left            =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CCom.frx":0000
   ScaleHeight     =   4170
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Top             =   4395
   Width           =   4245
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[V]atrix: REBIRTH by Freeza and FreAka"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4095
   End
   Begin CD_Player_2.CDPlayer CDPlayer1 
      Left            =   9105
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label3 
      Caption         =   "C:\Program Files\Napster\Music"
      Height          =   135
      Left            =   4080
      TabIndex        =   3
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C:\Program Files\Napster\Music"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mp3 Directory:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Abadi MT Condensed Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   3375
      Left            =   4080
      TabIndex        =   0
      Top             =   2040
      Width           =   4290
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
      DisplayBackColor=   
      DisplayForeColor=   
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
      VideoBorderColor=   
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin chatscan³.Chat Chat1 
      Left            =   480
      Top             =   2280
      _ExtentX        =   4022
      _ExtentY        =   2275
   End
   Begin ActiveDragControl.ActiveDrag ActiveDrag1 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub ActiveDrag1_MoveX(LeftX As Long)
    Left = LeftX
End Sub

Private Sub ActiveDrag1_MoveY(TopY As Long)
    Top = TopY
End Sub


Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If (Screen_Name) = GetUser And LCase(What_Said) = ".ad" Then
    ChatSend "-»[( [V]atriX: REBIRTH by Freeza & FreAka"
    ChatSend "-»[( Loaded by: " & GetUser
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".ownz *") Then
    What_Said = Mid(What_Said, 7)
    ChatSend "-»[( " & GetUser() & " owns " & What_Said
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".play *") Then
    What_Said = Mid(What_Said, 7)
    On Error Resume Next
    If FileExists("" + Label2.Caption + "\" + What_Said + ".mp3" = True) Then
    MediaPlayer1.filename = Label2.Caption & "\" & What_Said & ".mp3"
    MediaPlayer1.Play
    ChatSend "-»[( Now playing: " & What_Said
Else
ChatSend "-»[( " & What_Said & " not found"
End If
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".stop") Then
    On Error Resume Next
    MediaPlayer1.Stop
    ChatSend "-»[( Now Stoping"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".x *") Then
    What_Said = Mid(What_Said, 4)
    IMIgnore "" & What_Said
    ChatSend "-»[( " & What_Said & " Has Been [x]'d so STFU!"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".unx *") Then
    What_Said = Mid(What_Said, 6)
    IMUnIgnore "" & What_Said
    ChatSend "-»[( " & What_Said & ", Use ur speach wisely!"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".lol") Then
    ChatSend "-»[( That was funny, LoLz"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".lmao") Then
    ChatSend "-»[( I...Cant...STOP"
    Pause 0.5
    ChatSend "-»[( LAUGHING, LoLz!!"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".rofl") Then
    ChatSend "-»[( OMG!!!"
    Pause 0.5
    ChatSend "-»[( That was some..."
    Pause 0.5
    ChatSend "-»[( DEFINATE LOL MATERIAL!!"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".dir *") Then
    What_Said = Mid(What_Said, 6)
    Label2.Caption = What_Said
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".linkClr *") Then
    What_Said = Mid(What_Said, 10)
    ChatSend "<body link=#" & What_Said & ">-»[( Links are now RGB: " & What_Said
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".Kick *") Then
    What_Said = Mid(What_Said, 7)
    ChatSend "-»[( " + GetUser + " kicked: " + What_Said + " in the nuts!"
    Pause 0.5
    ChatSend "-»[( " & What_Said & "Ouch"
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".PlayCD") Then
    CDPlayer1.PlayCD
End If
    
 If (What_Said) Like ".im *" And Screen_Name = GetUser Then
  imee$ = Mid(What_Said, 4, InStr(1, What_Said, ",", 1) - 4)
  If imee$ <> "" Then
  message1$ = Trim(Mid(What_Said, InStr(1, What_Said, ",") + 1))
  If imee$ <> "" And message1$ <> "" Then
   Call InstantMessage(imee$, message1$)
   End If
        End If
        End If
If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".stopcd") Then
    CDPlayer1.StopCD
End If

If (Screen_Name) = GetUser And LCase(What_Said) Like LCase(".NextTrack") Then
    CDPlayer1.NextTrack
End If

If (ScreenName) = GetUser And LCase(What_Said) Like LCase(".PrevTrack") Then
    CDPlayer1.PreviousTrack
End If

If (ScreenName) = GetUser And LCase(What_Said) Like LCase(".pausecd") Then
    CDPlayer1.PauseCD
End If

If (ScreenName) = GetUser And LCase(What_Said) Like LCase(".unpausecd") Then
    CDPlayer1.UnPauseCD
End If


End Sub
Private Sub Chat2_ChatMessage(AolName As String, AolText As String)

End Sub

Private Sub Chat1_ChatMessage(AolName As String, AolText As String)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Chat1.ScanOn
    ChatSend "-»[( [V]atriX: REBIRTH by Freeza & FreAka"
    ChatSend "-»[( Loaded by: " & GetUser
    FormOnTop Me
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveDrag1.SnapRange = 195
    ActiveDrag1.DragNow Width, Height, Left, Top, Button, X, Y
End Sub


Private Sub NSPlay1_KeyDown(KeyCode As Integer, ShiftState As Integer)

End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveDrag1.SnapRange = 195
    ActiveDrag1.DragNow Width, Height, Left, Top, Button, X, Y
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveDrag1.SnapRange = 195
    ActiveDrag1.DragNow Width, Height, Left, Top, Button, X, Y

End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveDrag1.SnapRange = 195
    ActiveDrag1.DragNow Width, Height, Left, Top, Button, X, Y
End Sub


