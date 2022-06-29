VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Heres a simpel little thing"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Send Mail"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Im'sOff"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Im'sOn"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
IMsOn
Label2.Visible = True
Label1.Visible = False
End Sub

Private Sub Label2_Click()
IMsOff
Label1.Visible = True
Label2.Visible = False
End Sub

Private Sub Label3_Click()
Form2.Show
Form1.Hide
End Sub
