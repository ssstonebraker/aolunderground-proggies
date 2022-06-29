VERSION 4.00
Begin VB.Form Form13 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OreO 1.0 : Lag Scroll"
   ClientHeight    =   1410
   ClientLeft      =   4245
   ClientTop       =   3060
   ClientWidth     =   3390
   Height          =   1815
   Icon            =   "Form13.frx":0000
   Left            =   4185
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Top             =   2715
   Width           =   3510
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Text            =   $"Form13.frx":030A
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Text            =   $"Form13.frx":03FE
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   $"Form13.frx":04F2
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   $"Form13.frx":05E6
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   $"Form13.frx":06DA
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   $"Form13.frx":07CF
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   $"Form13.frx":08C4
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   $"Form13.frx":09B9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Text            =   "Sup Peeps"
      Top             =   720
      Width           =   1335
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1215
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   2143
      _StockProps     =   15
      BackColor       =   12632256
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
      BevelInner      =   1
      Begin Threed.SSCommand SSCommand4 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "OreO's"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "FuSH's"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "PoOHs"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Scroll Lag"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
End
Attribute VB_Name = "Form13"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendChat "<b><Font Face= Arial>" & BlackRedBlack(" I Lagged In The Chat Because I Am Lame-- " + UserSN() + "")
Form2.Show
Unload Form13
End Sub


Private Sub SSCommand1_Click()
SendChat Text1 + Text6 + Text7 + Text8 + Text9
timeout 5
Text1.Text = ""
End Sub

Private Sub SSCommand2_Click()
SendChat "<b><i><s>" & BlackRedBlack(" PoOH OwNs NiGGaz ")
SendChat Text2 + Text3 + Text4 + Text5
timeout 5

End Sub


Private Sub SSCommand3_Click()
SendChat "<b><i><s>" & BlackRedBlack(" FuSH Yu MaNG OwNs AOL")
SendChat Text2 + Text3 + Text4 + Text5
timeout 5
End Sub


Private Sub SSCommand4_Click()
SendChat "<b><i><s>" & BlackRedBlack(" OreO 1.0 for Aol 4.0")
SendChat Text2 + Text3 + Text4 + Text5
End Sub


