VERSION 4.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SuNZ Of MaN"
   ClientHeight    =   2430
   ClientLeft      =   3075
   ClientTop       =   1620
   ClientWidth     =   4320
   Height          =   2835
   Icon            =   "Form2.frx":0000
   Left            =   3015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Top             =   1275
   Width           =   4440
   Begin VB.Timer Timer3 
      Interval        =   400
      Left            =   720
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   360
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   480
      Picture         =   "Form2.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   0
      Picture         =   "Form2.frx":1A54
      ScaleHeight     =   555
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   15
         Index           =   0
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   855
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   1720
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
      BevelInner      =   1
      Begin Threed.SSCommand SSCommand3 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Options"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Anti Idle"
         ForeColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "FILE"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   1
         Picture         =   "Form2.frx":4B96
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   975
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   1720
      _StockProps     =   15
      BackColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BevelInner      =   1
      Begin Threed.SSCommand SSCommand9 
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "AOL Hide"
         ForeColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand8 
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Del. Chat"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "AOL show"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Advertise"
         ForeColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "IMz off"
         ForeColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "IMz On"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bauhaus 93"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1455
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   2566
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
      BevelInner      =   1
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bradley Hand ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bradley Hand ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bradley Hand ITC"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         Status Bar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
X = UserSN()
MsgBox "Wuz up " & X & " . This prog is da shit, So don't have to much fun ok???...LOL", vbExclamation, "Welcome!"
StayOnTop Me
SendChat ("<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #000000>•SuNz OF M<html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #999933>aN • BY: PoO<html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #993300>H • LoaDeD•<html></html><html><html></html><html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "PoOH OwN Yo Ass!"
'This Changes Label1.Caption to what you want
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendChat ("<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #000000>• SuNz OF M<html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #999933>aN • BY: PoO<html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #993300>H • UNLoaDeD<html></html><html><html></html><html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Shows u yo status F0o"
'This Changes Label1.Caption to what you want
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Itz Da Time F0o"
'This Changes Label1.Caption to what you want
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Itz Da Date F0o"
'This Changes Label1.Caption to what you want
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Duh! Itz your SN"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand1_Click()
Playwav ("Blip2.wav")
PopupMenu Form3.File
End Sub


Private Sub SSCommand1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Get all da info here"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand2_Click()
MsgBox "This Anti Will Stay On Until You Sign-off", vbExclamation, "Anti Idle"
Modal% = FindWindow("_AOL_Modal", vbNullString)
Stat% = FindChildByClass(Modal%, "_AOL_Static")
If Stat% <> 0 Then
AOIcon% = FindChildByClass(Modal%, "_AOL_Icon")
thetext = Gettext(Stat%)
If thetext = " You have been idle for a while. Do you want to stay online?" Then
Click% = SendMessage(AOIcon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(AOIcon%, WM_LBUTTONUP, 0, 0&)
End If
End If
End Sub


Private Sub SSCommand2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Anti Idle"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand3_Click()
Playwav ("Blip2.wav")
PopupMenu Form3.Opt
End Sub

Private Sub SSCommand3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Other Options"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand4_Click()
Call IMKeyword("$IM_ON", " PoOH OwNS Me")
End Sub

Private Sub SSCommand4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Turn IMz On"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand5_Click()
Call IMKeyword("$IM_ON", " PoOH OwNS Me")
End Sub

Private Sub SSCommand5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Turn IMz Off"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand6_Click()
Playwav ("Blip2.wav")
PopupMenu Form3.adver
End Sub

Private Sub SSCommand6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Advertise Dis Shit"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Show AOL"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Delete's Da Chat"
'This Changes Label1.Caption to what you want
End Sub


Private Sub SSCommand9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Hide AOL"
'This Changes Label1.Caption to what you want
End Sub


Private Sub Timer1_Timer()
Label4.Caption = "User: " + UserSN + " "
End Sub


Private Sub Timer2_Timer()
Label3.Caption = Format$(Now, "mm/dd/yy")
End Sub


Private Sub Timer3_Timer()
Label2.Caption = Format$(Now, "h:mm:ss AM/PM")
End Sub


