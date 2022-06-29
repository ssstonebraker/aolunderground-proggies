VERSION 5.00
Begin VB.Form Welcome 
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   Picture         =   "Welcome.frx":0000
   ScaleHeight     =   4290
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   960
      TabIndex        =   1
      Text            =   "New UseR?"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                 -  x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
'look at questions.frm
If Combo1.Text = "New UseR?" Then
    Questions.Show
    Me.Hide
    FormNotOnTop Me
    Me.Hide
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 0
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BorderStyle = 1
End Sub
