VERSION 5.00
Begin VB.Form frmWindowSpy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WindowSpy - Saiñt"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   Icon            =   "frmWindowSpy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblWinModule 
      BackColor       =   &H00000000&
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Label lblPCLass 
      BackColor       =   &H00404040&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label lblWinPText 
      BackColor       =   &H00808080&
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label lblWinPHandle 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label lblWinIDNum 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label lblWinStyle 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label lblWinTxt 
      BackColor       =   &H00808080&
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label lblWinClass 
      BackColor       =   &H00404040&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lblWinHdl 
      BackColor       =   &H00000000&
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmWindowSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call SetWindowPos(frmWindowSpy.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Private Sub Timer1_Timer()
Call WindowSPY(lblWinHdl, lblWinClass, lblWinTxt, lblWinStyle, lblWinIDNum, lblWinPHandle, lblWinPText, lblPCLass, lblWinModule)
End Sub
