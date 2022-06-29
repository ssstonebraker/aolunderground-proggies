VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add odd Handle"
   ClientHeight    =   765
   ClientLeft      =   -135
   ClientTop       =   1590
   ClientWidth     =   1665
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Add odd Handle"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Odd Handle"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Form7.Text1.text = Form3.Text1.text
 Form13.Text1.text = Form3.Text1.text

End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub
