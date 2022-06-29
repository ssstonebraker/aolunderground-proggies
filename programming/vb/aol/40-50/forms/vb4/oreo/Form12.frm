VERSION 4.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OreO 1.0 : Scroller"
   ClientHeight    =   1395
   ClientLeft      =   3885
   ClientTop       =   3075
   ClientWidth     =   3660
   Height          =   1800
   Icon            =   "Form12.frx":0000
   Left            =   3825
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Top             =   2730
   Width           =   3780
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Text            =   "**************************************"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Text            =   "$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000C0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "PoOH OwNS Me"
      Top             =   840
      Width           =   1695
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   1931
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
      BevelWidth      =   4
      BevelInner      =   1
      Begin Threed.SSCommand SSCommand6 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "***"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "$$$"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "@@@"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MadAve"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   2
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Stop Spiral"
      ForeColor       =   192
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
   Begin Threed.SSCommand SSCommand2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Spiral Scroll"
      ForeColor       =   0
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Scroll"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MadAve"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
End
Attribute VB_Name = "Form12"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SendChat "<b><Font Face= Arial>" & BlackRedBlack(" I Scrolled In The Chat Because I Am Lame-- " + UserSN() + "")
Form2.Show
Unload Form12
End Sub


Private Sub SSCommand1_Click()
Let RN = Int(Rnd * 1 + 1)
If RN = 1 Then Let ALE$ = Chr(4) & (Text1.Text)

SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.02)
SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.02)

End Sub


Private Sub SSCommand2_Click()
Do
SendChat (Text1)
timeout (0.7)
Dim MyLen As Integer
MyString = Text1.Text
MyLen = Len(MyString)
MyStr = Mid(MyString, 2, MyLen) + Mid(MyString, 1, 1)
Text1.Text = MyStr
Loop Until Label1.caption = "0"
Label1.caption = " "
End Sub


Private Sub SSCommand4_Click()
Let RN = Int(Rnd * 1 + 1)
If RN = 1 Then Let ALE$ = Chr(4) & (Text2.Text)

SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.03)
SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.02)
End Sub

Private Sub SSCommand5_Click()
Let RN = Int(Rnd * 1 + 1)
If RN = 1 Then Let ALE$ = Chr(4) & (Text3.Text)

SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.03)
SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.02)
End Sub


Private Sub SSCommand6_Click()
Let RN = Int(Rnd * 1 + 1)
If RN = 1 Then Let ALE$ = Chr(4) & (Text4.Text)

SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.03)
SendChat ("<pre=" & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) & ALE$ & Chr(4) + "<b><font face='Arial'>OReO 1.0")
timeout (0.02)
End Sub


