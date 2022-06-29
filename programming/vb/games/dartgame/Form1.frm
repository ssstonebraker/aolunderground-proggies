VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DART 1.0        Made by    Jason Fleury"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCriquet 
      Caption         =   "CRIQUET"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "301"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "501"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin MediaPlayerCtl.MediaPlayer wav 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
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
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.PictureBox BitBltpic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   670
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.PictureBox BitBltpic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   670
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":0262
      ScaleHeight     =   615
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4440
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.PictureBox bitblttarget 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   600
      Picture         =   "Form1.frx":04C4
      ScaleHeight     =   4575
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label lblWin 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Caption         =   "START A NEW GAME"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      Caption         =   "PLAYER 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      Caption         =   "301"
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      Caption         =   "301"
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTurn 
      Alignment       =   2  'Center
      Caption         =   "Turn 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "PLAYER 2"
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PLAYER 1"
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Bar2 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   600
      Top             =   5640
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   480
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Shape Bar1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   5400
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   5280
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ValX As Integer
Public ValY As Integer
Public down As Boolean
Public vtop As Integer
Public vleft As Integer
Public Premier As Boolean
Public Player As Byte
Public Deuxieme As Boolean
Public Coup As Integer
Public Turn As Integer
Private Sub Command1_Click()
Dim tmp3 As Variant
If Deuxieme = True Then
    Deuxieme = False
    wav.Stop
    wav.filename = App.Path & "\dingding.wav"
    wav.Play
    If Bar2.Left > 2760 And Bar1.Top < 2760 Then
        Call Pts1
    End If
    If Bar2.Left > 2760 And Bar1.Top >= 2760 Then
        Call Pts2
    End If
    If Bar2.Left <= 2760 And Bar1.Top >= 2760 Then
        Call Pts3
    End If
    If Bar2.Left <= 2760 And Bar1.Top < 2760 Then
        Call Pts4
    End If
    If Coup Mod 3 = 1 Then Form1.bitblttarget.Refresh
    
    Call BitBltTarget_MouseMove(0, 0, (Bar2.Left - 680) / Screen.TwipsPerPixelY, (Bar1.Top - 680) / Screen.TwipsPerPixelX)
    Coup = Coup + 1
    If Coup Mod 3 = 1 Then
        If Fleche = 0 Then
            Fleche = 1
        Else
            Fleche = 0
            Turn = Turn + 1
            lblTurn.Caption = "Turn " & Turn
        End If
        If Fleche = 0 Then
            Player = 1
            lblPlayer.Caption = "PLAYER 1"
        Else
            Player = 2
            lblPlayer.Caption = "PLAYER 2"
        End If
    End If
    Exit Sub
End If
If Premier = True And Deuxieme = False Then
    Premier = False
    Deuxieme = True
End If
If Premier = False And Deuxieme = False Then Premier = True
End Sub

Private Sub Command2_Click()
Score(1) = 501
Score(2) = 501
Command3.Visible = False
Command2.Visible = False
cmdCriquet.Visible = False
Label1.Visible = False
Command1.Enabled = True
Label2.Visible = True
Label3.Visible = True
Score(2).Visible = True
Score(1).Visible = True
End Sub

Private Sub Command3_Click()
Score(1) = 301
Score(2) = 301
Command3.Visible = False
Command2.Visible = False
cmdCriquet.Visible = False
Label1.Visible = False
Command1.Enabled = True
Label2.Visible = True
Label3.Visible = True
Score(2).Visible = True
Score(1).Visible = True
End Sub


Private Sub Form_Load()
Coup = 1
Fleche = 0
Premier = False
vtop = 10
vleft = 10
Turn = 1
Deuxieme = False
Player = 1
End Sub

Private Sub Label2_Click()
If Label2.ForeColor = &H80000012 Then
    Label2.ForeColor = &HFF&
Else: Label2.ForeColor = &H80000012
End If

End Sub

Private Sub Label3_Click()
If Label3.ForeColor = &H80000012 Then
    Label3.ForeColor = &HFF&
Else: Label3.ForeColor = &H80000012
End If

End Sub

Private Sub Label4_Click()
If Label4.ForeColor = &H80000012 Then
    Label4.ForeColor = &HFF&
Else: Label4.ForeColor = &H80000012
End If

End Sub

Private Sub Label5_Click()
If Label5.ForeColor = &H80000012 Then
    Label5.ForeColor = &HFF&
Else: Label5.ForeColor = &H80000012
End If

End Sub

Private Sub Timer1_Timer()
If Premier = True Then
    If Bar1.Top >= Shape1.Height - 100 Then
        down = True
    Else
        If Bar1.Top <= Shape1.Top + 250 Then
            down = False
        End If
    End If
    If down = True Then
        Bar1.Top = Bar1.Top - 150
    Else: Bar1.Top = Bar1.Top + 150
    End If
Else
    If Deuxieme = True Then
        If Bar2.Left <= Shape2.Left + 300 Then
            down = True
        Else
            If Bar2.Left >= Shape2.Width Then
                down = False
            End If
        End If
        If down = True Then
            Bar2.Left = Bar2.Left + 150
        Else: Bar2.Left = Bar2.Left - 150
        End If
    End If
End If

End Sub

Public Sub Pts1()
        tmp = (Bar2.Left) - 2760
        tmp2 = 2760 - ((Bar1.Top))
    tmp3 = Sqr((tmp ^ 2) + (tmp2 ^ 2))
        If tmp < tmp2 Then
            If (ASin(tmp / tmp3)) <= 9 Then TmpPts = 20
            If (ASin(tmp / tmp3)) > 9 And (ASin(tmp / tmp3)) <= 27 Then TmpPts = 1
            If (ASin(tmp / tmp3)) > 27 And (ASin(tmp / tmp3)) <= 45 Then TmpPts = 18
            If (ASin(tmp / tmp3)) > 45 And (ASin(tmp / tmp3)) <= 63 Then TmpPts = 4
            If (ASin(tmp / tmp3)) > 63 And (ASin(tmp / tmp3)) <= 81 Then TmpPts = 13
            If (ASin(tmp / tmp3)) > 81 And (ASin(tmp / tmp3)) <= 99 Then TmpPts = 6
            'MsgBox tmp & " " & tmp2 & " " & tmp3 & " -  " & (ASin(tmp / tmp3))
        Else
            'MsgBox tmp & " " & tmp2 & " " & tmp3 & " -  " & 90 - (ASin(tmp2 / tmp3))
            If 90 - (ASin(tmp2 / tmp3)) > 0 And 90 - (ASin(tmp2 / tmp3)) <= 9 Then TmpPts = 20
            If 90 - (ASin(tmp2 / tmp3)) > 9 And 90 - (ASin(tmp2 / tmp3)) <= 27 Then TmpPts = 1
            If 90 - (ASin(tmp2 / tmp3)) > 27 And 90 - (ASin(tmp2 / tmp3)) <= 45 Then TmpPts = 18
            If 90 - (ASin(tmp2 / tmp3)) > 45 And 90 - (ASin(tmp2 / tmp3)) <= 63 Then TmpPts = 4
            If 90 - (ASin(tmp2 / tmp3)) > 63 And 90 - (ASin(tmp2 / tmp3)) <= 81 Then TmpPts = 13
            If 90 - (ASin(tmp2 / tmp3)) > 81 And 90 - (ASin(tmp2 / tmp3)) <= 99 Then TmpPts = 6
        End If
        If tmp3 > 1956 Then Exit Sub
        If tmp3 <= 1956 And tmp3 >= 1810 Then
            Call Pts(TmpPts * 2)
            Exit Sub
        End If
        If tmp3 <= 1340 And tmp3 >= 1205 Then
            Pts (TmpPts * 3)
            Exit Sub
        End If
        Pts (TmpPts)
        

End Sub

Public Sub Pts2()
        tmp = (Bar2.Left) - 2760
        tmp2 = ((Bar1.Top)) - 2760
    tmp3 = Sqr((tmp ^ 2) + (tmp2 ^ 2))
        If tmp < tmp2 Then
            If (ASin(tmp / tmp3)) <= 9 Then TmpPts = 3
            If (ASin(tmp / tmp3)) > 9 And (ASin(tmp / tmp3)) <= 27 Then TmpPts = 17
            If (ASin(tmp / tmp3)) > 27 And (ASin(tmp / tmp3)) <= 45 Then TmpPts = 2
            If (ASin(tmp / tmp3)) > 45 And (ASin(tmp / tmp3)) <= 63 Then TmpPts = 15
            If (ASin(tmp / tmp3)) > 63 And (ASin(tmp / tmp3)) <= 81 Then TmpPts = 10
            If (ASin(tmp / tmp3)) > 81 And (ASin(tmp / tmp3)) <= 99 Then TmpPts = 6
            'MsgBox tmp & " " & tmp2 & " " & tmp3 & " -  " & (ASin(tmp / tmp3))
        Else
            'MsgBox tmp & " " & tmp2 & " " & tmp3 & " -  " & 90 - (ASin(tmp2 / tmp3))
            If 90 - (ASin(tmp2 / tmp3)) > 0 And 90 - (ASin(tmp2 / tmp3)) <= 9 Then TmpPts = 3
            If 90 - (ASin(tmp2 / tmp3)) > 9 And 90 - (ASin(tmp2 / tmp3)) <= 27 Then TmpPts = 17
            If 90 - (ASin(tmp2 / tmp3)) > 27 And 90 - (ASin(tmp2 / tmp3)) <= 45 Then TmpPts = 2
            If 90 - (ASin(tmp2 / tmp3)) > 45 And 90 - (ASin(tmp2 / tmp3)) <= 63 Then TmpPts = 15
            If 90 - (ASin(tmp2 / tmp3)) > 63 And 90 - (ASin(tmp2 / tmp3)) <= 81 Then TmpPts = 10
            If 90 - (ASin(tmp2 / tmp3)) > 81 And 90 - (ASin(tmp2 / tmp3)) <= 99 Then TmpPts = 6
        End If
        If tmp3 > 1956 Then Exit Sub
        If tmp3 <= 1956 And tmp3 >= 1810 Then
            Call Pts(TmpPts * 2)
            Exit Sub
        End If
        If tmp3 <= 1340 And tmp3 >= 1205 Then
            Pts (TmpPts * 3)
            Exit Sub
        End If
        Pts (TmpPts)

End Sub


Public Sub Pts3()
        tmp = 2760 - (Bar2.Left)
        tmp2 = ((Bar1.Top)) - 2760
    tmp3 = Sqr((tmp ^ 2) + (tmp2 ^ 2))
        If tmp < tmp2 Then
            If (ASin(tmp / tmp3)) <= 9 Then TmpPts = 3
            If (ASin(tmp / tmp3)) > 9 And (ASin(tmp / tmp3)) <= 27 Then TmpPts = 19
            If (ASin(tmp / tmp3)) > 27 And (ASin(tmp / tmp3)) <= 45 Then TmpPts = 7
            If (ASin(tmp / tmp3)) > 45 And (ASin(tmp / tmp3)) <= 63 Then TmpPts = 16
            If (ASin(tmp / tmp3)) > 63 And (ASin(tmp / tmp3)) <= 81 Then TmpPts = 8
            If (ASin(tmp / tmp3)) > 81 And (ASin(tmp / tmp3)) <= 99 Then TmpPts = 11
            'MsgBox tmp & " " & tmp2 & " " & tmp3 & " -  " & (ASin(tmp / tmp3))
        Else
            'MsgBox tmp & " " & tmp2 & " " & tmp3 & " -  " & 90 - (ASin(tmp2 / tmp3))
            If 90 - (ASin(tmp2 / tmp3)) > 0 And 90 - (ASin(tmp2 / tmp3)) <= 9 Then TmpPts = 3
            If 90 - (ASin(tmp2 / tmp3)) > 9 And 90 - (ASin(tmp2 / tmp3)) <= 27 Then TmpPts = 19
            If 90 - (ASin(tmp2 / tmp3)) > 27 And 90 - (ASin(tmp2 / tmp3)) <= 45 Then TmpPts = 7
            If 90 - (ASin(tmp2 / tmp3)) > 45 And 90 - (ASin(tmp2 / tmp3)) <= 63 Then TmpPts = 16
            If 90 - (ASin(tmp2 / tmp3)) > 63 And 90 - (ASin(tmp2 / tmp3)) <= 81 Then TmpPts = 8
            If 90 - (ASin(tmp2 / tmp3)) > 81 And 90 - (ASin(tmp2 / tmp3)) <= 99 Then TmpPts = 11
        End If
        If tmp3 > 1956 Then Exit Sub
        If tmp3 <= 1956 And tmp3 >= 1810 Then
            Call Pts(TmpPts * 2)
            Exit Sub
        End If
        If tmp3 <= 1340 And tmp3 >= 1205 Then
            Pts (TmpPts * 3)
            Exit Sub
        End If
        Pts (TmpPts)
End Sub

Public Sub Pts4()
Dim TmpPts As Integer
        tmp = 2760 - (Bar2.Left)
        tmp2 = 2760 - ((Bar1.Top))
    tmp3 = Sqr((tmp ^ 2) + (tmp2 ^ 2))
        If tmp < tmp2 Then
            If (ASin(tmp / tmp3)) <= 9 Then TmpPts = 20
            If (ASin(tmp / tmp3)) > 9 And (ASin(tmp / tmp3)) <= 27 Then TmpPts = 5
            If (ASin(tmp / tmp3)) > 27 And (ASin(tmp / tmp3)) <= 45 Then TmpPts = 12
            If (ASin(tmp / tmp3)) > 45 And (ASin(tmp / tmp3)) <= 63 Then TmpPts = 9
            If (ASin(tmp / tmp3)) > 63 And (ASin(tmp / tmp3)) <= 81 Then TmpPts = 14
            If (ASin(tmp / tmp3)) > 81 And (ASin(tmp / tmp3)) <= 99 Then TmpPts = 11
            
        Else
            
            If 90 - (ASin(tmp2 / tmp3)) > 0 And 90 - (ASin(tmp2 / tmp3)) <= 9 Then TmpPts = 20
            If 90 - (ASin(tmp2 / tmp3)) > 9 And 90 - (ASin(tmp2 / tmp3)) <= 27 Then TmpPts = 5
            If 90 - (ASin(tmp2 / tmp3)) > 27 And 90 - (ASin(tmp2 / tmp3)) <= 45 Then TmpPts = 12
            If 90 - (ASin(tmp2 / tmp3)) > 45 And 90 - (ASin(tmp2 / tmp3)) <= 63 Then TmpPts = 9
            If 90 - (ASin(tmp2 / tmp3)) > 63 And 90 - (ASin(tmp2 / tmp3)) <= 81 Then TmpPts = 14
            If 90 - (ASin(tmp2 / tmp3)) > 81 And 90 - (ASin(tmp2 / tmp3)) <= 99 Then TmpPts = 11
        End If
        If tmp3 > 1956 Then Exit Sub
        If tmp3 <= 1956 And tmp3 >= 1810 Then
            Call Pts(TmpPts * 2)
            Exit Sub
        End If
        If tmp3 <= 1340 And tmp3 >= 1205 Then
            Pts (TmpPts * 3)
            Exit Sub
        End If
        Pts (TmpPts)
        
End Sub

Public Sub Pts(Points As Integer)

    If Val(Score(Player).Caption) - Points >= 0 Then
        Score(Player).Caption = Val(Score(Player).Caption) - Points
    End If
    If Score(Player).Caption = 0 Then
        wav.Stop
        wav.filename = App.Path & "\appl1.wav"
        wav.Play
        Command1.Enabled = False
        lblWin = "Player " & Player & " WIN!!!!"
        Command2.Visible = True
        Command3.Visible = True
        cmdCriquet.Visible = True
        Label1.Visible = True
    End If
End Sub

