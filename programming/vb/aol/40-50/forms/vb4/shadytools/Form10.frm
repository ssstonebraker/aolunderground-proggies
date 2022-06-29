VERSION 4.00
Begin VB.Form Form10 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "røøm âddér"
   ClientHeight    =   975
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2565
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1380
   Left            =   1080
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   2685
   Begin VB.ListBox List1 
      Columns         =   1
      Height          =   1035
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "add room"
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
Attribute VB_Name = "Form10"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
Call FormOnTop(Me)
End Sub


Private Sub SSPanel1_Click()
Call AddRoomToListbox(List1, False)
End Sub


