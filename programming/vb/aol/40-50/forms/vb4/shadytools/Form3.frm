VERSION 4.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "idlè bøt"
   ClientHeight    =   375
   ClientLeft      =   3555
   ClientTop       =   3120
   ClientWidth     =   2190
   Height          =   780
   Left            =   3495
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Top             =   2775
   Width           =   2310
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   360
      Top             =   600
   End
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
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "start"
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub SSCommand1_Click()

End Sub


Private Sub Form_Load()
Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
Call FormOnTop(Me)
End Sub


Private Sub SSPanel1_Click()
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
Call FormOnTop(Me)
Call ChatSend("•·'¯`v› <B>s</B>hády <B>t</B>øølz¹")
Pause ".6"
Call ChatSend("•·'¯`v› <B>i</b>dlè <B>b</b>øt")
Pause ".6"
Call ChatSend("•·'¯`v› <B>r</B>êásõñ: <I>" & Text1.Text & " ")
End Sub


