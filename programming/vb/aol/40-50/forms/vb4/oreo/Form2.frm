VERSION 4.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OreO 1.0 By: PoOH & FuSH Yu MaNG"
   ClientHeight    =   1590
   ClientLeft      =   3195
   ClientTop       =   1635
   ClientWidth     =   3990
   Height          =   1995
   Left            =   3135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   Top             =   1290
   Width           =   4110
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3840
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   1080
      TabIndex        =   17
      Text            =   "Steve Case"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   0
      Top             =   4560
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   873
      _StockProps     =   15
      BackColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Outline         =   -1  'True
      Begin Threed.SSCommand SSCommand4 
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "MISC."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mead Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "MAIL"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mead Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "CHAT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mead Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "FILE"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Mead Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   1680
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "OreO Punter"
      ForeColor       =   0
      BackColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
      BevelInner      =   1
      Font3D          =   2
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   720
      Width           =   255
   End
   Begin Threed.SSCommand SSCommand15 
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   1920
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Exit Punter"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand14 
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   2520
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Stop"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand13 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Error"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand12 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "LAG"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand11 
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   720
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Punter"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand10 
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   1080
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Lagz"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand9 
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   720
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Fake Prog"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand8 
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Clear Chat"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   720
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Advertise"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   88
      X2              =   88
      Y1              =   32
      Y2              =   88
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "I M'z"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   600
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "IM'z Off"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "IM'z On"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " DATE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Mead Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("_∏.∑∞∑-[ï ÿRÎO π∑∞   ÉÚr ¿ˆL 4∑∫")
timeout 0.2
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("  `∑.∏.∑∞∑--[ï ﬂy : PÙOH & FuSH Yu MaNG")
timeout 0.2
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("       '∑.∏∏.∑∞∑-[ï LÙÂDÎ– ﬂy: " + UserSN() + "")
timeout 1
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("             '∑.∏∏.∑∞∑-[ï !~—ÂBÔäcO~! ")
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("_∏.∑∞∑-[ï ÿRÎO π∑∞   ÉÚr ¿ˆL 4∑∫")
timeout 0.2
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("  `∑.∏.∑∞∑--[ï ﬂy : PÙOH & FuSH Yu MaNG")
timeout 0.2
Playwav ("StreetJeopardy.wav")
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("       '∑.∏∏.∑∞∑-[ï UnlÙÂDÎ– ﬂy: " + UserSN() + "")
timeout 1
SendChat "<i><b><Font Face= Arial Narrow>" & BlackRedBlack("             '∑.∏∏.∑∞∑-[ï !~—ÂBÔäcO~! ")
Form2.Left = 5
timeout (0.1)
Form2.Left = 400
timeout (0.1)
Form2.Left = 700
timeout (0.1)
Form2.Left = 1000
timeout (0.1)
Form2.Left = 2000
timeout (0.1)
Form2.Left = 3000
timeout (0.1)
Form2.Left = 4000
timeout (0.1)
Form2.Left = 5000
timeout (0.1)
Form2.Left = 4000
timeout (0.1)
Form2.Left = 3000
timeout (0.1)
Form2.Left = 2000
timeout (0.1)
Form2.Left = 1000
timeout (0.1)
Form2.Left = 700
timeout (0.1)
Form2.Left = 400
timeout (0.1)
Form2.Left = 5
timeout (0.1)
Form2.Left = 400
timeout (0.1)
Form2.Left = 700
timeout (0.1)
Form2.Left = 1000
timeout (0.1)
Form2.Left = 2000
timeout (0.1)
Form2.Left = 5
timeout (0.1)
Form2.Left = 400
timeout (0.1)
Form2.Left = 700
timeout (0.1)
Form2.Left = 1000
timeout (0.1)
Form2.Left = 2000
timeout (0.1)
Form2.Left = 3000
timeout (0.1)
Form2.Left = 4000
timeout (0.1)
Form2.Left = 5000
timeout (0.1)
Form2.Left = 4000
timeout (0.1)
Form2.Left = 3000
timeout (0.1)
Form2.Left = 2000
timeout (0.1)
Form2.Left = 1000
timeout (0.1)
Form2.Left = 700
timeout (0.1)
Form2.Left = 400
timeout (0.1)
Form2.Left = 5
timeout (0.1)
Form2.Left = 400
timeout (0.1)
Form2.Left = 700
timeout (0.1)
Form2.Left = 1000
timeout (0.1)
Form2.Left = 2000
End
End Sub

Private Sub SSCommand1_Click()
PopupMenu Form3.File
End Sub

Private Sub SSCommand10_Click()
Form13.Show
Form2.Hide
End Sub

Private Sub SSCommand11_Click()

End Sub

Private Sub SSCommand13_Click()
Timer2.Enabled = True
End Sub

Private Sub SSCommand14_Click()
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub SSCommand15_Click()

End Sub

Private Sub SSCommand2_Click()
PopupMenu Form3.ch
End Sub

Private Sub SSCommand3_Click()
PopupMenu Form3.ma
End Sub

Private Sub SSCommand4_Click()
PopupMenu Form3.misc
End Sub

Private Sub SSCommand5_Click()

Call IMKeyword("$IM_ON", "OreO Is Da Shit ")
End Sub

Private Sub SSCommand6_Click()

Call IMKeyword("$IM_OFF", " OreO Is Da Shit")
End Sub

Private Sub SSCommand7_Click()

End Sub

Private Sub SSCommand8_Click()

End Sub

Private Sub SSCommand9_Click()
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format$(Now, "mm/dd/yy")
Label2.Caption = Format$(Now, "h:mm:ss AM/PM")
Label3.Caption = "User: " + UserSN + " "
End Sub


Private Sub Timer2_Timer()
Do
Call IMKeyword(Text1, "<font = 999999999999999999999999999999999999999999999999999999999999999999999999999999 999999999999999999999999999999999999999999999999999999999999999999999999999999 999999999999999999999999999999999999999999999999999999999999999999999999999999 999999999999999999999999999999999999999999999999999999999999999999999999999999 999999999999999999999999999999999999999999999999999999999999999999999999999999 999999999999999999999999999999999999999999999999999999999999999999999999999999 999999999999999999999>")
timeout (0.4)
Loop
End Sub


Private Sub Timer3_Timer()
Do
Call IMKeyword(Text1, "'<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>'<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>'<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>'<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>")
timeout (0.4)
Loop
End Sub


