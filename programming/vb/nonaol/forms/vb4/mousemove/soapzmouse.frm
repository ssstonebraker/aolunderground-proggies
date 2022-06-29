VERSION 4.00
Begin VB.Form soapzhowto21 
   Caption         =   "soap"
   ClientHeight    =   1005
   ClientLeft      =   2895
   ClientTop       =   2820
   ClientWidth     =   2535
   Height          =   1410
   Left            =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   2535
   Top             =   2475
   Width           =   2655
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Text            =   "soap made this"
      Top             =   540
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   ":]"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "this"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "like"
      Height          =   195
      Left            =   1740
      TabIndex        =   1
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "just"
      Height          =   195
      Left            =   1380
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
   Begin VB.Menu main 
      Caption         =   "main"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu chat 
         Caption         =   "chat"
      End
      Begin VB.Menu ims 
         Caption         =   "ims"
      End
      Begin VB.Menu grr 
         Caption         =   "grr"
      End
   End
End
Attribute VB_Name = "soapzhowto21"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Caption = ":P"

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = "made by soap"
Command1.Caption = ";]"
Label3.ForeColor = "&H000000"
Label1.ForeColor = "&H000000"
Label2.ForeColor = "&H000000"
End Sub


Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Caption = ":]"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = "made by soap"
Command1.Caption = ":]"
Label3.ForeColor = "&H000000"
Label1.ForeColor = "&H000000"
Label2.ForeColor = "&H000000"
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = "made by soap"
Command1.Caption = ":]"
Label1.ForeColor = "&H00FF00"
Label2.ForeColor = "&H000000"
Label3.ForeColor = "&H000000"
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = "made by soap"
Command1.Caption = ":]"
Label2.ForeColor = "&H00FF00"
Label1.ForeColor = "&H000000"
Label3.ForeColor = "&H000000"
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = "made by soap"
Command1.Caption = ":]"
Label3.ForeColor = "&H00FF00"
Label1.ForeColor = "&H000000"
Label2.ForeColor = "&H000000"
End Sub


Private Sub madeby_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = "so gimme props"
Command1.Caption = ":]"
Label3.ForeColor = "&H000000"
Label1.ForeColor = "&H000000"
Label2.ForeColor = "&H000000"
End Sub


Private Sub text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Caption = ":]"
Label2.ForeColor = "&H000000"
Label1.ForeColor = "&H000000"
Label3.ForeColor = "&H000000"
text1.Text = "so gimme some props"

End Sub


