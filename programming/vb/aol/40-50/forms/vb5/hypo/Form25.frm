VERSION 5.00
Begin VB.Form Form25 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Echo Bot"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "close "
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "WHo to Echo"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Echo"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Form25.Hide
Unload Form25
End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Timer1_Timer()
Text2 = SNFromLastChatLine
Text2 = LCase(Text2)

If LCase(Text1) = Text2 Then
SendChat LastChatLine
 TimeOut 0.9
End If
 End Sub
