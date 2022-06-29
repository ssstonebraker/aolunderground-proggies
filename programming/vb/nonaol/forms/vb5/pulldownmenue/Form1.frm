VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Pulldown menue example By KnK"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Butten example"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Label Example"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.PopupMenu Form2.buttens, 2
'This one uses a 2 since its the second pull down menue on Form2
End Sub

Private Sub Label1_Click()
Form2.PopupMenu Form2.Label, 1
'This one uses a 1 since its the first pull down menue on Form2
End Sub
