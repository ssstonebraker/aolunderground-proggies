VERSION 5.00
Object = "{BC326F64-5766-11D5-9845-001E5AC10000}#3.0#0"; "CHATSCAN20.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   30
   ClientLeft      =   -18060
   ClientTop       =   9075
   ClientWidth     =   30
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mp3.frx":0000
   ScaleHeight     =   30
   ScaleWidth      =   30
   ShowInTaskbar   =   0   'False
   Begin TBChatScan20.TBScan chat1 
      Left            =   360
      Top             =   4800
      _ExtentX        =   2196
      _ExtentY        =   953
   End
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

Private Sub File1_Click()
Dim Asci As String
Dim Esci As String
Asci = "</b></i></u><font face=""arial Narrow""><font color=''#000000''><^"
Esci = "^>"
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Stop
Pause 1
MediaPlayer1.Open File1.Path & "\" & File1.FileName
        
    ChatSend (Asci) & "þLåýïñg [<u>" & ReplaceString(File1.FileName, ".mp3", "") & "</u>]" & (Esci)
Exit Sub
End If
MediaPlayer1.Open File1.Path & "\" & File1.FileName
        ChatSend (Asci) & "þLåýïñg [<u>" & ReplaceString(File1.FileName, ".mp3", "") & "</u>]" & (Esci)
End Sub

Private Sub Form_Load()
chat1.Scan_On
main.com.AddItem ".dir  -  sets directory"
main.com.AddItem ".mute  -  mutes mp3 player"
main.com.AddItem ".pause  -  pauses mp3 player"
main.com.AddItem ".play  -  plays mp3"
main.com.AddItem ".rand  -  plays a random mp3"
main.com.AddItem ".ref  -  refreshes mp3's"
main.com.AddItem ".stop  -  stops mp3 playing"
main.com.AddItem ".vol  -  sets the volume"
main.com.AddItem ".xmute  -  unmutes mp3 player"
main.com.AddItem ".xpause  -  unpauses mp3"
main.com.AddItem ".mp3s - Views mp3s"
On Error Resume Next
File1.Path = GetFromINI("Settings", "dir", "c:\windows\system\cap.set")
End Sub

Private Sub Chat1_scan(Screen_Name As String, What_Said As String)
Dim Asci As String
Dim Esci As String
Asci = "</b></i></u><font face=""arial Narrow""><font color=''#000000''><^"
Esci = "^>"
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
      ChatSend (Asci) & "Mp3 Paused" & (Esci)
      
       Case ".xpause"
      MediaPlayer1.Play
      ChatSend (Asci) & "mp3 unPaused" & (Esci)

Case ".repon"
MediaPlayer1.PlayCount = 0
ChatSend (Asci) & "looping now enabled" & (Esci)


Case ".mp3s"
ChatSend (Asci) & "viewing mp3's" & (Esci)
mp3s.Show

Case ".repoff"
MediaPlayer1.PlayCount = 1
ChatSend (Asci) & "looping now disabled" & (Esci)

Case ".ref"
File1.Refresh
ChatSend (Asci) & "mp3s RéfréshéÐ " & File1.ListCount & " †ø†ål" & (Esci)

Case ".vol"
If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
            If strArgument1 > 100 Then
            ChatSend (Asci) & "Mp3 Volume too high" & (Esci)
            Else
            MediaPlayer1.Volume = (strArgument1$ * 25) - 2500
            ChatSend (Asci) & "Mp3 Volume at <u>" + strArgument1 + "</u>" & (Esci)
            End If
            End If

Case ".rand"
Dim DD As String
DD = GetFromINI("settings", "dir", "c:\windows\system\cap.set")
If DD = "" Then
ChatSend (Asci) & "Dir Needed" & (Esci)
Else
On Error Resume Next
            ChatSend (Asci) & "Random mp3" & (Esci)
            File1.ListIndex = RandomNumber1(File1.ListCount)
            strArgument1$ = File1.Path & "\" & File1.FileName
            On Error Resume Next
            MediaPlayer1.Open strArgument1$
            MediaPlayer1.Play
            
End If
Case ".play"
DD = GetFromINI("settings", "dir", "c:\windows\system\cap.set")
If DD = "" Then
ChatSend (Asci) & "Dir Needed" & (Esci)
Else
On Error Resume Next
If MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Stop
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
    Text1.text = "now playing [" & ReplaceString(List2.List(intt%), ".mp3", "") & "]"
    If Text1.text = "now playing []" Then
    ChatSend (Asci) & "invaled number" & (Esci)
    Exit Sub
    End If
    ChatSend (Asci) & "þLåýïñg [<u>" & ReplaceString(List2.List(intt%), ".mp3", "") & "</u>]" & (Esci)
    Multi = False
    List2.Clear
    Exit Sub
kdmerr:
    If What_Said <> (Asci) & "[type ''.play'' and 0 - " & List2.ListCount - 1 & "or ''.cancel'']" & (Esci) Then ChatSend (Asci) & "invaled number" & (Esci)
End If
End If

Case ".cancel"
ChatSend (Asci) & "Mult canceled" & (Esci)
List2.Clear
Multi = False


Case ".stop"
ChatSend (Asci) & "Mp3 Now Stopped" & (Esci)
MediaPlayer1.Stop



Case ".mute"
ChatSend (Asci) & "Mp3 Now Mute" & (Esci)
MediaPlayer1.Mute = True


Case ".xmute"
ChatSend (Asci) & "Mp3 is now unmuted" & (Esci)
MediaPlayer1.Mute = False

Case ".dir"
strArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument1$
      If File1.ListCount = 0 Then
       ChatSend (Asci) & "No Mp3s In This Dir" & (Esci)
       Call WriteToINI("Settings", "dir", "", "c:\windows\system\cap.set")
      Else
     Call WriteToINI("Settings", "dir", strArgument1$, "c:\windows\system\cap.set")
ChatSend "<font face=""arial Narrow""><font color=#000000>" & (Asci) & "dir is Set To <u>" + strArgument1 + "" & (Esci)
ChatSend (Asci) & File1.ListCount & " †ø†ål Mp3s" & (Esci)
End If
End Select
End If
End Sub
Public Sub PlayIt(sString As String)
Asci = "</b></i></u><font face=""arial Narrow""><font color=''#000000''><^"
Esci = "^>"
For i = 0 To File1.ListCount - 1
        If InStr(LCase(TrimSpaces(File1.List(i))), LCase(TrimSpaces(sString$))) Then
            List2.AddItem (File1.List(i))
        End If
    Next i
    If List2.ListCount = 1 Then
        MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        ChatSend (Asci) & "þLåýïñg [<u>" & ReplaceString(List2.List(0), ".mp3", "") & "</u>]" & (Esci)
        List2.Clear
        Exit Sub
    End If
    If List2.ListCount = 0 Then
        ChatSend "<font face=""arial Narrow""><font color=''#000000''>Lèåñ Høw Spèll Yøù Mø®øñ"
    End If
    If List2.ListCount > 1 Then
        ChatSend (Asci) & "Šç®ølling <u>Mp3s</u>" & (Esci)
        For i = 0 To List2.ListCount - 1
            ChatSend (Asci) & "[<u>" & i & ") " & List2.List(i) & "</u>]" & (Esci)
            Pause 1
        Next i
        ChatSend (Asci) & "[type .play and 0 - " & List2.ListCount - 1 & " or .cancel ]" & (Esci)
        Multi = True
    End If
End Sub

Public Sub PlayIt2(sString As String)
Asci = "</b></i></u><font face=""arial Narrow""><font color=''#000000''><^"
Esci = "^>"
MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        ChatSend (Asci) & "þLåýïñg [<u>" & ReplaceString(List2.List(0), ".mp3", "") & "</u>]" & (Esci)
        List2.Clear
        End Sub
