VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000009&
   Caption         =   "Check"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1920
   LinkTopic       =   "Form4"
   ScaleHeight     =   1140
   ScaleWidth      =   1920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.text = "Diana + Jerry" Then
MsgBox "That's Really Me...heh I'm cool", vbOKOnly, "Water Rapids"
MsgBox "Oh and if you wanna know Diana and Jerry are my parents", vbOKOnly, "Water Rapids"
MsgBox "Yeah IM Da MaN!", vbOKOnly, "Water Rapids"
Else
MsgBox "that ain't me...kill the mutha fucka", vbOKOnly, "Water Rapids"
MsgBox "is this annoying?", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
MsgBox "heh", vbOKOnly, "Water Rapids"
End If
End Sub

Private Sub Form_Load()
StayOnTop Me
CenterForm Me
End Sub
