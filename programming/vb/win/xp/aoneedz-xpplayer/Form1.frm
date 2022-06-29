VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{7314ED99-8643-4E82-A4F8-5E9F4DEC14BE}#1.0#0"; "VolumeControl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6975
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3255
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   240
      TabIndex        =   35
      Top             =   7200
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   600
      MousePointer    =   3  'I-Beam
      TabIndex        =   27
      Top             =   720
      Width           =   855
   End
   Begin VB.ListBox List6 
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   25
      Top             =   6240
      Width           =   735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   5520
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      MousePointer    =   1  'Arrow
      Pattern         =   "*.mp3"
      TabIndex        =   16
      Top             =   1680
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   135
      Left            =   120
      Max             =   100
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   840
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   600
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   3255
      TabIndex        =   4
      Top             =   10
      Width           =   3285
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3000
         TabIndex        =   33
         Top             =   -75
         Width           =   60
      End
      Begin VB.Label lbclose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   3120
         TabIndex        =   32
         Top             =   -45
         Width           =   75
      End
      Begin VB.Label lbmin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "•"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2880
         TabIndex        =   31
         Top             =   -40
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xp amp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Double click to roll up"
         Top             =   -38
         Width           =   510
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   495
      Left            =   960
      TabIndex        =   36
      Top             =   4440
      Width           =   855
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
   Begin VolControl.VolumeControl VolumeControl1 
      Left            =   720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   100
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ref."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2950
      TabIndex        =   30
      ToolTipText     =   "Refresh"
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   2280
      TabIndex        =   29
      Top             =   2550
      Width           =   45
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "skip :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cont."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1920
      TabIndex        =   23
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "repeat"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1440
      TabIndex        =   22
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "arand"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   960
      TabIndex        =   21
      Top             =   1320
      Width           =   390
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "files: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   2550
      Width           =   435
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1635
      TabIndex        =   19
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   2040
      TabIndex        =   18
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lbtimenum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2160
      TabIndex        =   17
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "view mp3's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "View Mp3's"
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "random"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2000
      TabIndex        =   13
      ToolTipText     =   "Random"
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1600
      TabIndex        =   12
      ToolTipText     =   "Mute"
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pause"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2580
      TabIndex        =   11
      ToolTipText     =   "Pause"
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1800
      TabIndex        =   10
      ToolTipText     =   "Stop"
      Top             =   480
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "idlein..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   720
      TabIndex        =   9
      Top             =   1140
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   480
      TabIndex        =   8
      Top             =   1335
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vol:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   7
      Top             =   1335
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "status:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "random :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "play :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/XP Amp (Windows xp)
'/Fully coded by: Paul Whitehead aka killz
'/Programming Language: VB6ee
'/Aim: a6c
'/Aol sn: Aoneedz

Private Sub File1_DblClick()
If Me.MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Stop
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
Pause 0.5
Form1.MediaPlayer1.Open File1.Path & "\" & File1.FileName
HScroll1.Value = VolumeControl1.Volume
vl$ = HScroll1.Value
Label5.Caption = "" & vl$ & "%"
Me.Timer1.Enabled = True
Me.HScroll1.Enabled = True
Label10.Enabled = True
Me.Label1.Caption = "playing:"
Me.Label6 = "" & File1.FileName
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = "1560"
Me.Label11.Caption = "view mp3's"
Else
Form1.MediaPlayer1.Open File1.Path & "\" & File1.FileName
HScroll1.Value = VolumeControl1.Volume
vl$ = HScroll1.Value
Label5.Caption = "" & vl$ & "%"
Me.Timer1.Enabled = True
Me.HScroll1.Enabled = True
Label10.Enabled = True
Me.Label1.Caption = "playing:"
Me.Label6 = "" & File1.FileName
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = "1560"
Me.Label11.Caption = "view mp3's"
End If
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
End Sub
Private Sub Form_Load()
Me.Height = 1560
Me.Width = 3285
StayOnTop Me
Call FadeBy2(Picture1, vbBlack, vbWhite)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormNoCaption Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RespondToTray(X) <> 0 Then Call ShowFormAgain(Me)
End Sub

Private Sub HScroll1_Change()
vl$ = Me.HScroll1.Value
Label5.Caption = "" & vl$ & "%"
VolumeControl1.Volume = HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
vl$ = Me.HScroll1.Value
Label5.Caption = "" & vl$ & "%"
VolumeControl1.Volume = HScroll1.Value
End Sub
Private Sub Label10_Click()
If MediaPlayer1.PlayState = mpPlaying Then
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
MediaPlayer1.Stop
Pause 0.5
Form1.File1.ListIndex = RandomNumber(Form1.File1.ListCount)
Form1.MediaPlayer1.Open Form1.File1.Path & "\" & Form1.File1.FileName
Me.Timer1.Enabled = True
Me.HScroll1.Enabled = True
HScroll1.Value = VolumeControl1.Volume
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
Else
Form1.File1.ListIndex = RandomNumber(Form1.File1.ListCount)
Form1.MediaPlayer1.Open Form1.File1.Path & "\" & Form1.File1.FileName
Me.Timer1.Enabled = True
    Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Me.HScroll1.Enabled = True
HScroll1.Value = VolumeControl1.Volume
vl$ = HScroll1.Value
Label5.Caption = "" & vl$ & "%"
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
End If
End Sub
Private Sub Label11_Click()
If Me.Label11.Caption = "view mp3's" Then
Do: DoEvents
Me.Height = Me.Height + 1
Loop Until Me.Height = 2775
Me.Label11.Caption = "hide mp3's"
Me.Label11.ToolTipText = "Hide Mp3's"
cnt$ = Me.File1.ListCount
Me.Label15.Caption = "files: " & cnt$ & ""
Label10.Enabled = False
Else
If Me.Label11.Caption = "hide mp3's" Then
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = "1560"
Me.Label11.Caption = "view mp3's"
Me.Label11.ToolTipText = "View Mp3's"
Label10.Enabled = True
End If
End If
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Label12.ForeColor = &HFF&
        Label22.ForeColor = &HFF&
        lbclose.ForeColor = &HFF&
MoveFormNoCaption Me
        Label12.ForeColor = &HFFFFFF
        Label22.ForeColor = &H0&
        lbclose.ForeColor = &H0&
        Else
    End If
End Sub
Private Sub Label16_Click()
If Me.Label6.Caption = "idlein..." Then
Exit Sub
Else
If Me.Label16.ForeColor = &H80000012 Then
Me.Label16.ForeColor = &H808080
Me.Label17.ForeColor = &H80000012
Me.Label18.ForeColor = &H80000012
Else
Me.Label16.ForeColor = &H80000012
End If
End If
End Sub
Private Sub Label17_Click()
If Me.Label6.Caption = "idlein..." Then
Exit Sub
Else
If Me.Label17.ForeColor = &H80000012 Then
Me.Label17.ForeColor = &H808080
Me.Label16.ForeColor = &H80000012
Me.Label18.ForeColor = &H80000012
Else
Me.Label17.ForeColor = &H80000012
End If
End If
End Sub

Private Sub Label18_Click()
If Me.Label6.Caption = "idlein..." Then
Exit Sub
Else
If Me.Label18.ForeColor = &H80000012 Then
Me.Label18.ForeColor = &H808080
Me.Label16.ForeColor = &H80000012
Me.Label17.ForeColor = &H80000012
Else
Me.Label18.ForeColor = &H80000012
End If
End If
End Sub

Private Sub Label20_Click()
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = "1560"
List1.Visible = False
List1.Clear
File1.Visible = True
Label10.Enabled = True
Label11.Enabled = True
Label20.Caption = ""
End Sub

Private Sub Label21_Click()
fresh$ = Form1.File1.ListCount
If fresh$ = "" Then
Exit Sub
Else
a$ = Me.Label1.Caption
b$ = Me.Label6.Caption
Me.Label1.Caption = "status:"
Me.Label6.Caption = "refreshing playlist"
Form1.File1.Refresh
Pause 0.5
fresh2$ = Form1.File1.ListCount
¹% = fresh2$
²% = fresh$
³% = ¹% - ²%
Me.Label6.Left = 840
Me.Label1.Caption = "refreshed:"
Me.Label6.Caption = "" & ³% & " out of " & fresh$ & " mp3's"
Pause 1
Me.Label6.Left = 720
Me.Label1.Caption = "" & a$ & ""
Me.Label6.Caption = "" & b$ & ""
End If
End Sub

Private Sub Label23_Click()
Me.WindowState = 1
Form2.Show
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Label22.ForeColor = vbRed
End Sub

Private Sub Label22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Label22.ForeColor = vbBlack
Pause 0.5
Call AddToTray(Form1.Icon, Form1.Caption, Form1)
End Sub

Private Sub Label7_Click()
If Me.Label6.Caption = "idlein..." Then
Exit Sub
Else
Me.Label1.Caption = "stoped:"
Me.MediaPlayer1.Stop
Pause 0.5
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
End If
If Label8.Caption = "xpause" Then
MediaPlayer1.Stop
Pause 0.5
Label1.Caption = "status:"
Label6.Caption = "idlein..."
Label8.Caption = "pause"
Label10.Enabled = True
Label11.Enabled = True
Label9.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Label5.Caption = "0%"
Me.HScroll1.Enabled = False
End If
If Label9.Caption = "xmute" Then
MediaPlayer1.Stop
Pause 0.5
Label1.Caption = "status:"
Label6.Caption = "idlein..."
Me.Label9.Caption = "mute"
Me.Label9.Left = 1600
Label10.Enabled = True
Label11.Enabled = True
Label9.Enabled = True
Label8.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Label5.Caption = "0%"
Me.HScroll1.Enabled = False
End If
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Label5.Caption = "0%"
Me.HScroll1.Enabled = False
End Sub

Private Sub Label8_Click()
If Me.Label6.Caption = "idlein..." Then
Exit Sub
Else
If Me.Label8.Caption = "pause" Then
Me.Label8.Caption = "xpause"
Me.Label8.ToolTipText = "UnPause"
Me.MediaPlayer1.Pause
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Label9.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text4.Enabled = False
Me.Label1.Caption = "paused:"

Else
If Me.Label8.Caption = "xpause" Then
Me.Label8.Caption = "pause"
Me.Label8.ToolTipText = "Pause"
MediaPlayer1.Play
Me.Label1.Caption = "playing:"
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Label9.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
End If
End If
End If
End Sub

Private Sub Label9_Click()
If Me.Label6.Caption = "idlein..." Then
Exit Sub
Else
If Me.Label9.Caption = "mute" Then
Me.Label9.Left = 1520
Me.Label9.Caption = "xmute"
Me.Label9.ToolTipText = "UnMute"
Me.MediaPlayer1.Mute = True
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Label8.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text4.Enabled = False
Label16.Enabled = False
Label17.Enabled = False
Label18.Enabled = False
Me.Label1.Caption = "mutted:"

Else
If Me.Label9.Caption = "xmute" Then
Me.Label9.Left = 1600
Me.Label9.Caption = "mute"
Me.Label9.ToolTipText = "Mute"
MediaPlayer1.Mute = False
Me.Label1.Caption = "playing:"
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Label8.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
End If
End If
End If
End Sub

Private Sub lbclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lbclose.ForeColor = vbRed
End Sub

Private Sub lbclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lbclose.ForeColor = vbBlack
Pause 0.5
End
End Sub

Private Sub lbmin_Click()
If Me.Height = 1560 Then
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = 210
Me.lbmin.ForeColor = vbRed
Else
If Me.Height = 210 Then
Do: DoEvents
Me.Height = Me.Height + 1
Loop Until Me.Height = 1560
Me.lbmin.ForeColor = vbBlack
End If
End If
End Sub

Private Sub List1_DblClick()
If Me.MediaPlayer1.PlayState = mpPlaying Then
MediaPlayer1.Stop
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
Pause 0.5
Form1.MediaPlayer1.Open File1.Path & "\" & List1
HScroll1.Value = VolumeControl1.Volume
vl$ = HScroll1.Value
Label5.Caption = "" & vl$ & "%"
Label20.Caption = ""
Me.Timer1.Enabled = True
Label11.Enabled = True
Label10.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Caption = "playing:"
Me.Label6 = "" & List1
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = "1560"
Me.Label11.Caption = "view mp3's"
Me.List1.Clear
Me.List1.Visible = False
Me.File1.Visible = True
Else
Form1.MediaPlayer1.Open File1.Path & "\" & List1
HScroll1.Value = VolumeControl1.Volume
vl$ = HScroll1.Value
Label5.Caption = "" & vl$ & "%"
Label20.Caption = ""
Me.Timer1.Enabled = True
Label11.Enabled = True
Label10.Enabled = True
Me.HScroll1.Enabled = True
Me.Label1.Caption = "playing:"
Me.Label6 = "" & List1
Do: DoEvents
Me.Height = Me.Height - 1
Loop Until Me.Height = "1560"
Me.Label11.Caption = "view mp3's"
Me.List1.Clear
Me.List1.Visible = False
Me.File1.Visible = True
End If
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal result As Long)
If Me.Label16.ForeColor = "&H808080" Then
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
MediaPlayer1.Stop
Pause 0.5
Form1.File1.ListIndex = RandomNumber(Form1.File1.ListCount)
Form1.MediaPlayer1.Open Form1.File1.Path & "\" & Form1.File1.FileName
Me.Timer1.Enabled = True
Me.HScroll1.Enabled = True
HScroll1.Value = VolumeControl1.Volume
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
Else
If Me.Label17.ForeColor = "&H808080" Then
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
MediaPlayer1.Stop
Pause 0.5
Form1.MediaPlayer1.Open File1.Path & "\" & File1.FileName
MediaPlayer1.ClickToPlay = True
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
Else
If Me.Label18.ForeColor = "&H808080" Then
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
MediaPlayer1.Stop
Pause 0.5
File1.ListIndex = File1.ListIndex + 1
Form1.MediaPlayer1.Open File1.Path & "\" & File1.FileName
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
Else
If MediaPlayer1.PlayState = mpStopped Then
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
Me.HScroll1.Enabled = False
Label5.Caption = "0%"
Timer1.Enabled = False
Label14.Caption = "00:00"
lbtimenum.Caption = "00:00"
End If
End If
End If
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Label12.ForeColor = &HFF&
        Label22.ForeColor = &HFF&
        lbclose.ForeColor = &HFF&
MoveFormNoCaption Me
        Label12.ForeColor = &HFFFFFF
        Label22.ForeColor = &H0&
        lbclose.ForeColor = &H0&
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
For a = 0 To File1.ListCount 'Loops through the list to see if the string is in each song
If InStr(LCase(File1.List(a)), LCase(Text1.Text)) Then 'The string is in this song!  add it to the list!
List1.AddItem File1.List(a) 'ADDS
End If
Next a 'Loops
cnt$ = Me.List1.ListCount
If cnt$ = 0 Then
Strng$ = Label6.Caption
Me.Label1.Caption = "error:"
Me.Label6.Caption = "string not found"
Text1 = ""
Pause 0.5
Me.Label1.Caption = "status:"
Me.Label6.Caption = "" & Strng$
Exit Sub
Else
List1.Visible = True
File1.Visible = False
Label10.Enabled = False
Me.Height = 2775
Label20.Caption = "cancel?"
Label11.Enabled = False
Me.Label15.Caption = "found " & cnt$ & " strings for " & Text1.Text & ""
Text1.Text = ""
If List1.ListCount = 1 Then  'There was only one string found,   play it
Me.Label1.Caption = "status:"
Me.Label6.Caption = "idlein..."
Me.Height = 1560
Label11.Enabled = True
Me.List1.Visible = False
Me.File1.Visible = True
Pause 0.5
Form1.MediaPlayer1.Open Form1.File1.Path & "\" & List1.List(0) 'play
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & List1.List(0)
HScroll1.Value = VolumeControl1.Volume
HScroll1.Enabled = True
Timer1.Enabled = True
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Label10.Enabled = True
Text1 = ""
List1.Clear 'Make sure you clear this or you'll get errors
Exit Sub
End If
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
For J = 0 To File1.ListCount 'Loops through the list to see if the string is in each song
If InStr(LCase(File1.List(J)), LCase(Text2.Text)) Then 'The string is in this song!  add it to the list!
List1.AddItem File1.List(J) 'ADDS
End If
Next J 'Loops
If List1.ListCount = 1 Then  'There was only one string found,   play it
Form1.MediaPlayer1.Open Form1.File1.Path & "\" & List1.List(0) 'play
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & List1.List(0)
HScroll1.Value = VolumeControl1.Volume
HScroll1.Enabled = True
Timer1.Enabled = True
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Text2 = ""
List1.Clear 'Make sure you clear this or you'll get errors
Exit Sub
End If
If List1.ListCount = 0 Then 'There were no songs with thoes strings found!
Strng$ = Label6.Caption
Me.Label1.Caption = "error:"
Me.Label6.Caption = "string not found"
Text2 = ""
Pause 0.5
Me.Label1.Caption = "status:"
Me.Label6.Caption = "" & Strng$
List1.Clear
Exit Sub
End If
    Randomize Timer
 MyValue = Int((List1.ListCount * Rnd))
 List1.ListIndex = MyValue
     Text3.Text = List1
     List6 = List1
     Dim err As String
     err = List1
     err = ReplaceString(err, ".mp3", "")
     Text3 = err
     Form1.Label6 = Text3
    If Text3.Text <> "" Then
Form1.MediaPlayer1.Open File1.Path & "\" & List1
Label16.Enabled = True
Label17.Enabled = True
Label18.Enabled = True
Form1.Label6 = "" & List1
Timer1.Enabled = True
Me.HScroll1.Enabled = True
HScroll1.Value = VolumeControl1.Volume
Me.Text2 = ""
List1.Clear
End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Me.Label6.Caption = "idlein..." Then
Text4.Text = ""
Exit Sub
Else
sec$ = Text4.Text
If sec$ >= MediaPlayer1.Duration Then
seconds$ = MediaPlayer1.Duration
Me.Label1.Caption = "error:"
Me.Label6.Caption = "pick a lower number then " & seconds$ & ""
Text4.Text = ""
Pause 1
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
Else
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + sec$
Me.Label1.Caption = "skipped:"
Me.Label6.Caption = "skipped: " & sec$ & " seconds"
Text4.Text = ""
Pause 1
Me.Label1.Caption = "playing:"
Me.Label6.Caption = "" & File1.FileName
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
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
    Songlength$ = strMinutes$ & ":" & strSeconds$
    Me.Label14.Caption = Songlength$
On Error Resume Next
Dim format1
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
    Songlength$ = strMinutes$ & ":" & strSeconds$
     str1Seconds$ = MediaPlayer1.CurrentPosition
    If InStr(str1Seconds$, ".") <> 0& Then
        str1Seconds$ = Right(str1Seconds$, Len(str1Seconds$) - InStr(str1Seconds$, "*"))
    End If
    lng1Seconds& = str1Seconds$
    Do Until lng1Seconds& <= 59
        DoEvents
        lng1Seconds& = lng1Seconds& - 60
        lng1Minutes& = lng1Minutes& + 1
    Loop
    str1Seconds$ = Format$(lng1Seconds&, "00")
    str1Minutes$ = Format$(lng1Minutes&, "00")
    songlength1$ = str1Minutes$ & ":" & str1Seconds$
    Form1.lbtimenum = songlength1$
End Sub
