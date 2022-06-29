VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Auto scroll text by nick(txt)"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   3420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton stop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton start 
      Caption         =   "Start"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   600
   End
   Begin VB.TextBox Text 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Text            =   "scrolling text....             "
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "'This is a example written by nick(txt) to show how to Auto scroll text in a text box"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = False
End Sub



Private Sub start_Click()
Timer1.Enabled = True
End Sub

Private Sub stop_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Dim getfirst As String
Dim gettext As String
Dim getcount As String
'this gets the first letter/number of the text
getfirst$ = Mid(Text, 1, 1)
'this gets the lenght of the textbox
getcount$ = Len(Text.Text)
'this will get the text starting at the
'second letter or number.
gettext$ = Mid(Text, 2, getcount$)
'adds the text without the first letter/number
'then adds the first letter/number after it
Text.Text = gettext$ & getfirst$
End Sub
