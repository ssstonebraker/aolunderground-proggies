VERSION 4.00
Begin VB.Form Form11 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat"
   ClientHeight    =   225
   ClientLeft      =   1695
   ClientTop       =   1530
   ClientWidth     =   1560
   Height          =   630
   Left            =   1635
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Top             =   1185
   Width           =   1680
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   225
      Left            =   870
      TabIndex        =   1
      Top             =   0
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   885
   End
End
Attribute VB_Name = "Form11"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Unload Form11
End Sub

Private Sub Command2_Click()
End
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


