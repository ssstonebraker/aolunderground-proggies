VERSION 4.00
Begin VB.Form Form7 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OReO 1.0 : Mail PoOH & FuSH"
   ClientHeight    =   3255
   ClientLeft      =   3420
   ClientTop       =   1830
   ClientWidth     =   4590
   Height          =   3660
   Icon            =   "Form7.frx":0000
   Left            =   3360
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Top             =   1485
   Width           =   4710
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Sup PoOH  & FuSH -"
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "SUP OreO FreaKs!!!"
      Top             =   1080
      Width           =   1935
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1508
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
      BevelWidth      =   4
      BorderWidth     =   2
      BevelInner      =   1
      Outline         =   -1  'True
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Send"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Font3D          =   2
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "What to say"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject--->"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Poobear172@Hotmail.com"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form7"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form2.Show
Unload Form7
End Sub


Private Sub SSCommand1_Click()
SendMail "poohbear172@Hotmail.com", Text1, Text2
MsgBox "Thanks for givin' me some feedback!", vbInformation, "Thanx!"
End Sub


