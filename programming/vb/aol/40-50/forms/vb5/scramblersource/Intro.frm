VERSION 5.00
Object = "{05589FA0-C356-11CE-BF01-00AA0055595A}#2.0#0"; "AMOVIE.OCX"
Object = "{A847FB88-4D77-11D2-A11D-549F06C10000}#1.0#0"; "PRJSMOOTHBAR.OCX"
Begin VB.Form intro 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5685
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   108
      Left            =   2160
      Top             =   2520
   End
   Begin prjSmoothBar.SmoothProgress SmoothProgress1 
      Height          =   255
      Left            =   600
      Top             =   4080
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
   End
   Begin AMovieCtl.ActiveMovie ActiveMovie1 
      Height          =   5460
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   9631
      ShowDisplay     =   0   'False
      ShowControls    =   0   'False
      ShowTracker     =   0   'False
      AutoStart       =   -1  'True
      FileName        =   "C:\My Documents\Scrambler\Main\test.avi"
   End
End
Attribute VB_Name = "intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
SmoothProgress1.Value = SmoothProgress1.Value + 1
If SmoothProgress1.Value >= 100 Then Timer1.Enabled = False
End Sub
