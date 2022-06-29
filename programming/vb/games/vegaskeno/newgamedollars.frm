VERSION 5.00
Begin VB.Form newgamedollars 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "newgamedollars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   120
      Top             =   1335
   End
   Begin VB.Image Image2 
      Height          =   1770
      Left            =   15
      Top             =   90
      Width           =   4740
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "newgamedollars.frx":000C
      Top             =   570
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Don't forget to save your game (credits and marks will be saved)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1005
      TabIndex        =   1
      Top             =   540
      Width           =   3645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "You have 50 credits to start with. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   4815
   End
End
Attribute VB_Name = "newgamedollars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
Picture1.Picture = LoadResPicture("600coins", bitmap)
End Sub

Private Sub Image2_Click()
Unload Me
dollars = 50
board.Show
End Sub

Private Sub Timer1_Timer()
Unload Me
dollars = 50
board.Show
End Sub
