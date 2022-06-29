VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   2145
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Hint "
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "!"
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.text = "Big Willy Style Baby" Then
MsgBox "you got it correct!", vbOKOnly, "Water Rapids"
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ " & UserSN + " got the secret access")
Form6.Show
Form5.Hide
Else
MsgBox "so sowwy you lose!", vbOKOnly, "Water Rapids"
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ " & UserSN + " didn't get the secret access")
End If
End Sub

Private Sub Command2_Click()
Form1.Show
Form5.Hide
End Sub

Private Sub Command3_Click()
MsgBox "Buein venidos a Miami"
End Sub

Private Sub Form_Load()
StayOnTop Me
FormPositionTopLeft Me
ChatSend "" & (" ")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Water Rapids")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ " & UserSN + " is now trying for the secret access")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ will he make it?")
TimeOut 0.3
ChatSend "" & (" ")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MvFrm Me
End Sub
