VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "mp3 stylist by maqic"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   Icon            =   "m3styl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1920
      Top             =   960
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   960
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   600
      ScaleHeight     =   255
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "mp3 stylist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   360
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   360
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   600
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   600
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1080
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   1080
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   1320
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ""
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   960
      Width           =   135
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   1200
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   840
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "II"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "u "
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   120
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label mp3lab 
      BackStyle       =   0  'Transparent
      Caption         =   "status : idle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "m3styl.frx":22A2
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   255
      Left            =   600
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Me
MsgBox "No file paused to play.", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Play
End Sub

Private Sub Command2_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Form1
MsgBox "No file playing to pause.", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Pause
End Sub

Private Sub Command3_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Form1
MsgBox "No file playing to stop.", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Stop
End Sub

Private Sub Form_Load()
dos32.FormOnTop Me
Call click32.FadeBy2(Me.Picture1, vbBabyBlue, vbNavyBlue)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.Visible = False
Shape8.Visible = False
Label10.Visible = False
Shape7.Visible = False
Label6.Visible = False
Shape6.Visible = False
Shape5.Visible = False
Label9.Visible = False
End Sub

Private Sub Image1_Click()
Me.PopupMenu Form2.file, 0, 4700, 20
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dos32.FormDrag Me
dos32.FormOnTop Me
End Sub

Private Sub Label2_Click()
dos32.FormNotOnTop Me
If MsgBox("Do you realy want to exit?", vbQuestion + vbYesNo, "mp3 stylist : exiting") = vbYes Then
End
Else:
Cancel = True
End If

End Sub

Private Sub Label3_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label4_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Form1
MsgBox "No file paused to play", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Play
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.Visible = True
Shape8.Visible = True
End Sub

Private Sub Label5_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Form1
MsgBox "No file to pause", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Pause
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.Visible = True
Shape7.Visible = True
End Sub

Private Sub Label6_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Form1
MsgBox "No file playing to stop", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Stop
End Sub

Private Sub Label7_Click()
If List1.Text = "" Then
dos32.FormNotOnTop Form1
MsgBox "No file playing to stop.", vbExclamation, "mp3 stylist : error"
Exit Sub
End If
MediaPlayer1.Stop
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.Visible = False
Label9.Visible = False
Label6.Visible = True
Shape6.Visible = True

End Sub

Private Sub Label8_Click()
Form1.CommonDialog1.Filter = ("*.mp3")
Form1.CommonDialog1.ShowOpen
If Form1.CommonDialog1.FileName = "" Then
Exit Sub
End If
mp32play = Form1.CommonDialog1.FileName
Form1.List1.AddItem mp32play
Form1.mp3lab.Caption = "status : mp3(s) loaded"
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.Visible = True
Label9.Visible = True

End Sub

Private Sub List1_Click()
Dim List1
MediaPlayer1.Open Form1.List1.Text
mp3lab.Caption = "status : playing mp3"

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dos32.FormDrag Me
dos32.FormOnTop Me
End Sub

Private Sub Timer1_Timer()
If List1.Text = "" Then
MediaPlayer1.Stop
End If
End Sub
