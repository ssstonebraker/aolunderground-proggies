VERSION 5.00
Begin VB.Form frmNewScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Congratulations!"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmNewScore.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmNewScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------
' frmScore. This Form is dependent upon the module
' modScore.BAS and the form frmScore.FRM
' ----------------------------------------------------------

Private Sub cmdOK_Click()
' -----------------------------------------------------------
' Close the window and add the name to the High Score list.
' -----------------------------------------------------------
If lngScore > 0 Then
    AddScoreAndSave txtName.Text, lngScore
    
    ' Display the High Score Table
    Load frmScore
    Me.Visible = False
    frmScore.Show vbModal, frmMain
    Unload Me
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub
