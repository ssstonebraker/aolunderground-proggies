VERSION 4.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ím ìgñøre"
   ClientHeight    =   375
   ClientLeft      =   4695
   ClientTop       =   2715
   ClientWidth     =   3270
   Height          =   780
   Icon            =   "Form7.frx":0000
   Left            =   4635
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Top             =   2370
   Width           =   3390
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "ignore"
      BackColor       =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin Threed.SSPanel SSPanel2 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "unignore"
      ForeColor       =   8388608
      BackColor       =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
End
Attribute VB_Name = "Form7"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
Call FormOnTop(Me)
End Sub

Private Sub SSPanel1_Click()
Call InstantMessage("$IM_OFF, " & Person$, "" & Text1.Text & "")
Pause ".6"
Call ChatSend("•·'¯`v› <B>í</B>m <B>ì</B>gñørïng: <I>" & Text1.Text & "")
End Sub


Private Sub SSPanel2_Click()
Call InstantMessage("$IM_ON, " & Person$, "" & Text1.Text & "")
End Sub


