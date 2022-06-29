VERSION 5.00
Object = "{A847FB88-4D77-11D2-A11D-549F06C10000}#1.0#0"; "PRJSMOOTHBAR.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3630
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   2400
   End
   Begin prjSmoothBar.SmoothProgress SmoothProgress1 
      Height          =   135
      Left            =   1200
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   238
      ColorBar        =   16761024
      ColorProgress   =   8388608
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing..."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
SmoothProgress1.Value = SmoothProgress1.Value + 2
If SmoothProgress1.Value = 100 Then Timer1.Enabled = False: Form2.Hide: Form1.Show
If SmoothProgress1.Value >= 15 And SmoothProgress1.Value < 50 Then Label1.Caption = "Loading Graphics"
If SmoothProgress1.Value >= 60 And SmoothProgress1.Value < 70 Then Label1.Caption = "Loading Categories"
If SmoothProgress1.Value >= 70 And SmoothProgress1.Value < 75 Then Label1.Caption = "Loading Points"
If SmoothProgress1.Value >= 75 And SmoothProgress1.Value < 90 Then Label1.Caption = "Loading Lists and High Scores": timeout (0.3)
If SmoothProgress1.Value >= 90 Then Label1.Caption = "Starting..."
End Sub
