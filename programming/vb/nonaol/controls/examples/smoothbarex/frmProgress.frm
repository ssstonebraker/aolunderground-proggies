VERSION 5.00
Object = "{A847FB88-4D77-11D2-A11D-549F06C10000}#1.0#0"; "prjSmoothBar.ocx"
Begin VB.Form frmProgress 
   Caption         =   "Smooth Progress Bar"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "cmdReset"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   480
   End
   Begin prjSmoothBar.SmoothProgress SmoothProgress1 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdReset_Click()
    SmoothProgress1.Value = 0
End Sub

Private Sub cmdStart_Click()
    Timer1.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    SmoothProgress1.Min = 0
    SmoothProgress1.Max = 120
    SmoothProgress1.Value = 0
End Sub

Private Sub Timer1_Timer()
    If SmoothProgress1.Value >= SmoothProgress1.Max Then
        MsgBox "Done"
        Timer1.Enabled = False
        Exit Sub
    Else
        SmoothProgress1.Value = SmoothProgress1.Value + 1
    End If
End Sub
