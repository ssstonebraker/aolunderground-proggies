VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send E-MAIL"
   ClientHeight    =   3855
   ClientLeft      =   3540
   ClientTop       =   1410
   ClientWidth     =   3330
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "KnK-mail.frx":0000
   ScaleHeight     =   3855
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   " Send"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   3135
   End
   Begin VB.OptionButton Other 
      BackColor       =   &H00000000&
      Caption         =   "Other"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Sugguestons"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Complaints"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "KnK-mail.frx":539B
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
If Option1 = True Then
 Call AOLMail("Bill@knk.tierranet.com", "Complaints -KnK Server Helper-", Text1)
 End If
If Option2 = True Then
Call AOLMail("Bill@knk.tierranet.com", "Sugustions -KnK Server helper-", Text1)
End If
If Option3 = True Then
Call AOLMail("Bill@knk.tierranet.com", "Other -KnK Server helper-", Text1)
End If
End If
If aversion$ = "aol4" Then
If Option1 = True Then
 Call SendMail("Bill@knk.tierranet.com", "Complaints -KnK Server Helper-", Text1)
 End If
If Option2 = True Then
Call SendMail("Bill@knk.tierranet.com", "Sugustions -KnK Server helper-", Text1)
End If
If Option3 = True Then
Call SendMail("Bill@knk.tierranet.com", "Other -KnK Server helper-", Text1)
End If
End If
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Option1_Click()
Text1.text = "Needs alot of work"

End Sub

Private Sub Option2_Click()
Text1.text = "Needs ........................."
End Sub

Private Sub Other_Click()
Text1.text = "???"
End Sub
