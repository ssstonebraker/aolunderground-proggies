VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Mp"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   Icon            =   "mp3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer LoopFile 
      Interval        =   1000
      Left            =   1320
      Top             =   3480
   End
   Begin VB.Timer LoopList 
      Interval        =   1000
      Left            =   1320
      Top             =   3480
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   2400
      Top             =   1440
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   975
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
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
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "__"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   0
      X2              =   3360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mp³ PlayerPlus"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   -840
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
      Case ".dir"
      strArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument1$
      ChatSend "‹¦­Mp³ PlayerPlus­¦› • dir set [" & strArgument1$ & "]"
      Call WriteToINI("mp3dir", "dir", strArgument1$, App.Path & "\mp3.dat")
      Case ".play"
      On Error Resume Next
       File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.dat")
      Dim possibles%, Others$
             strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace)
             thepart$ = strArgument1$
      thepart$ = LCase(ReplaceString(thepart$, " ", ""))
    File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.dat")
            
    thepart = Mid(What_Said, 7)
    thepart = LCase(ReplaceString(thepart, " ", ""))
For z = 0 To File1.ListCount - 1
Chcky1$ = File1.List(z)
d = InStr(1, Chcky1$, thepart, vbTextCompare)
If d > 0 Then
If Others = "" Then Others = File1.List(z): possibles = (possibles + 1) Else possibles = (possibles + 1)
End If
Next z
If possibles > 1 Then ChatSend "‹¦­Mp³ PlayerPlus­¦› [" & possibles% & "] Possibilitys": Exit Sub
If possibles = 0 Then ChatSend "‹¦­Mp³ PlayerPlus­¦› [" & thepart & "] not found": Exit Sub
fiiles$ = "" & File1.Path & "\" & Others & ""
MediaPlayer1.Open fiiles$
ChatSend "‹¦­Mp³ PlayerPlus­¦› • now playing [" & ReplaceString(Others, ".mp3", "") & "]"
             Case ".ran"
        On Error Resume Next
            File1.ListIndex = RandomNumber1(File1.ListCount)
            strArgument1$ = File1.Path & "\" & File1.filename
            On Error Resume Next
            MediaPlayer1.Open strArgument1$
            MediaPlayer1.Play
            ChatSend "‹¦­Mp³ PlayerPlus­¦› - Playing [" & ReplaceString(ReplaceString(strArgument1$, File1.Path, ""), "\", "") & "]"
        Case ".pause"
             MediaPlayer1.Pause
             ChatSend "‹¦­Mp³ PlayerPlus­¦› - Paused"
        Case ".stop"
            LoopList.Enabled = False
            LoopFile.Enabled = False
             MediaPlayer1.Stop
             ChatSend "‹¦­Mp³ PlayerPlus­¦› • stopped"
        Case ".unpause"
             MediaPlayer1.Pause
             ChatSend "‹¦­Mp³ PlayerPlus­¦› • UnPaused"
        Case ".loop"
            strArgument1$ = Mid(What_Said, 7)
            If LCase(strArgument1$) = LCase("list") Then LoopList.Enabled = True: LoopFile.Enabled = False: ChatSend "" & Text1 & "mp³ player • now looping list"
            If LCase(strArgument1$) = LCase("file") Then LoopFile.Enabled = True: LoopList.Enabled = False: ChatSend "" & Text1 & "mp³ player • now looping file"
       
       Case ".end"
       ChatSend "‹¦­Mp³ PlayerPlus­¦›‹¦­mage­¦›"
       End
       End Select
End If
End Sub

Private Sub Form_Load()
FormOnTop Me
ChatSend ("‹¦­Mp³ PlayerPlus­¦›‹¦­mage­¦›")
Chat1.ScanOn
On Error Resume Next
File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.dat")
End Sub

Private Sub Label2_Click()
WindowState = 1
End Sub

Private Sub Label3_Click()
ChatSend ("‹¦­Mp³ PlayerPlus­¦›‹¦­mage­¦›")
End
End Sub
