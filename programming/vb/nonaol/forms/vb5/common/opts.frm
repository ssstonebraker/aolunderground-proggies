VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2940
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6150
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Send 
         Caption         =   "Send "
      End
      Begin VB.Menu GetHTML 
         Caption         =   "Get HTML"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GetHTML_Click()
Form2.Text1 = Form1.Text2
Form2.Show
End Sub

Private Sub Send_Click()
ChatSend Form1.Text2
End Sub
