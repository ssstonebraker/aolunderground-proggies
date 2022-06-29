VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "shády tøølz¹"
   ClientHeight    =   855
   ClientLeft      =   4020
   ClientTop       =   2805
   ClientWidth     =   2430
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1545
   Icon            =   "Form1.frx":0000
   Left            =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   Top             =   2175
   Width           =   2550
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Text            =   "mailto:suessonline@hotmail.com"
      Top             =   2520
      Width           =   2415
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "shády tøølz¹"
      ForeColor       =   8388608
      BackColor       =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "vb5"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "vb6"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "vb4"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "vb"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "art2"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "art"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
   End
   Begin VB.Menu a0 
      Caption         =   "file"
      Begin VB.Menu a4 
         Caption         =   "about"
      End
      Begin VB.Menu a18 
         Caption         =   "minimize"
      End
      Begin VB.Menu a21 
         Caption         =   "keyword"
         Begin VB.Menu a22 
            Caption         =   "beta test"
         End
         Begin VB.Menu a25 
            Caption         =   "file search"
         End
         Begin VB.Menu a24 
            Caption         =   "hecklers"
         End
         Begin VB.Menu a26 
            Caption         =   "netfind"
         End
      End
   End
   Begin VB.Menu a1 
      Caption         =   "chat "
      Begin VB.Menu a6 
         Caption         =   "advertise"
      End
      Begin VB.Menu a9 
         Caption         =   "attention"
      End
      Begin VB.Menu a7 
         Caption         =   "error room"
      End
      Begin VB.Menu a8 
         Caption         =   "idle bot"
      End
      Begin VB.Menu a10 
         Caption         =   "lagger"
      End
      Begin VB.Menu a11 
         Caption         =   "linker"
      End
      Begin VB.Menu a27 
         Caption         =   "macro kill"
         Begin VB.Menu a29 
            Caption         =   "4 line"
         End
         Begin VB.Menu a30 
            Caption         =   "8 line"
         End
         Begin VB.Menu a31 
            Caption         =   "12 line"
         End
      End
      Begin VB.Menu a12 
         Caption         =   "re-enter"
      End
      Begin VB.Menu a33 
         Caption         =   "scroller"
      End
      Begin VB.Menu a13 
         Caption         =   "sound hell"
      End
   End
   Begin VB.Menu a2 
      Caption         =   "ims "
      Begin VB.Menu a15 
         Caption         =   "ims on"
      End
      Begin VB.Menu a14 
         Caption         =   "ims off"
      End
      Begin VB.Menu a16 
         Caption         =   "im ignore"
      End
      Begin VB.Menu a20 
         Caption         =   "im lagger"
      End
   End
   Begin VB.Menu a3 
      Caption         =   "mail "
      Begin VB.Menu a36 
         Caption         =   "lag a fag"
      End
      Begin VB.Menu a35 
         Caption         =   "room adder"
      End
   End
   Begin VB.Menu a5 
      Caption         =   "exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub a10_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause "1.5"
Call ChatSend("<html></html>.<html></html>.<html></html><html></html>.<html></html>.<html></html>.<html></html>.<html></html>.<html></html><html></html>.<html></html>.<html></html>.")
Pause "1.2"
Call ChatSend("<html></html>.<html></html>.<html></html><html></html>.<html></html>.<html></html>.<html></html>.<html></html>.<html></html><html></html>.<html></html>.<html></html>.")
Pause "1.1"
Call ChatSend("<html></html>.<html></html>.<html></html><html></html>.<html></html>.<html></html>.<html></html>.<html></html>.<html></html><html></html>.<html></html>.<html></html>.")
End Sub

Private Sub a11_Click()
Form5.Show
End Sub


Private Sub a12_Click()
Form6.Show
End Sub

Private Sub a13_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause ".6"
Call ChatSend("{S ygp}")
Pause ".6"
Call ChatSend("{S ygp}")
Pause ".6"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
Pause "1.4"
Call ChatSend("{S ygp}")
End Sub

Private Sub a14_Click()
Call InstantMessage("$IM_OFF", "suess!")
End Sub

Private Sub a15_Click()
Call InstantMessage("$IM_ON", "suess!")
End Sub

Private Sub a16_Click()
Form7.Show
End Sub

Private Sub a18_Click()
Form1.WindowState = 1
End Sub

Private Sub a20_Click()
Form8.Show
End Sub

Private Sub a22_Click()
Call Keyword("beta")
End Sub

Private Sub a24_Click()
Call Keyword(";)")
End Sub

Private Sub a25_Click()
Call Keyword("file search")
End Sub

Private Sub a26_Click()
Call Keyword("netfind")
End Sub

Private Sub a29_Click()
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
End Sub

Private Sub a30_Click()
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause ("1.7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")

End Sub

Private Sub a31_Click()
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause ("1.7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause ("1.7")
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
Pause 0.6
Call ChatSend("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  shády tøølz¹")
End Sub

Private Sub a33_Click()
Form9.Show
End Sub

Private Sub a35_Click()
MsgBox "i didn't ever get this function to work, ;[", 64, "sorry"
End Sub

Private Sub a36_Click()
Call SendMail("tomguy2001", "what's up?", "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>s<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>u<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>e<html></html><html></html>s<html></html>s.....<html></html><html></html><html></html><html></html>p<html></html><html></html>a<html></html><html></html>y<html></html>b<html></html>a<html></html>c<html></html>k<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!<html></html>!")

End Sub

Private Sub a4_Click()
Form2.Show
End Sub

Private Sub a5_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause ".6"
Call ChatSend("•·'¯`v› <B>b</b>ý <b>s</b>uèss")
Pause ".6"
Call ChatSend("•·'¯`v› <B>n</B>øw <B>u</B>nløadéd")
Call FormExitDown(Me)
Unload Me
End Sub



Private Sub a6_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause ".6"
Call ChatSend("•·'¯`v› <B>b</b>ý <b>s</b>uèss")
Pause ".6"
Call ChatSend("•·'¯`v› < a href=" & Text1 & ">request it here!</a>")
End Sub


Private Sub a7_Click()
Call FormOnTop(Me)
Call ChatSend("<font = 99999999999999999999999999999999999999999999999999999999999><B>s</B>hády <B>t</B>øølz¹")
End Sub

Private Sub a8_Click()
Form3.Show
End Sub

Private Sub a9_Click()
Form4.Show
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause ".6"
Call ChatSend("•·'¯`v› <B>b</b>ý <b>s</b>uèss")
Pause ".6"
Call ChatSend("•·'¯`v› <B>n</B>øw <B>l</B>øadéd")
End Sub

Private Sub SSCommand1_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>b</b>áíliñg <b>t</b>ø <b>p</b>r - <b>á</b>rt")
Call PrivateRoom("art")
End Sub


Private Sub SSCommand2_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>b</b>áíliñg <b>t</b>ø <b>p</b>r - <b>á</b>rt²")
Call PrivateRoom("art2")
End Sub

Private Sub SSCommand3_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>b</b>áíliñg <b>t</b>ø <b>p</b>r - <b>v</b>b")
Call PrivateRoom("vb")
End Sub

Private Sub SSCommand4_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>b</b>áíliñg <b>t</b>ø <b>p</b>r - <b>v</b>b4")
Call PrivateRoom("vb4")
End Sub

Private Sub SSCommand5_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>b</b>áíliñg <b>t</b>ø <b>p</b>r - <b>v</b>b6")
Call PrivateRoom("vb6")
End Sub

Private Sub SSCommand6_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>b</b>áíliñg <b>t</b>ø <b>p</b>r - <b>v</b>b5")
Call PrivateRoom("vb5")
End Sub

Private Sub SSPanel1_Click()
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause ".6"
Call ChatSend("•·'¯`v› < a href=" & Text1 & ">request it here!</a>")
End Sub


