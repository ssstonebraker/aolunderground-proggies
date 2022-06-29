VERSION 4.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "ábôut"
   ClientHeight    =   1305
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   4575
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1710
   Icon            =   "Form2.frx":0000
   Left            =   2655
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Top             =   3075
   Width           =   4695
   Begin Threed.SSPanel SSPanel1 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   2355
      _StockProps     =   15
      Caption         =   $"Form2.frx":0442
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
      BevelInner      =   1
      Alignment       =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
Call FormOnTop(Me)
End Sub


Private Sub SSPanel1_Click()
Call FormExitUp(Me)
Unload Form2
End Sub


