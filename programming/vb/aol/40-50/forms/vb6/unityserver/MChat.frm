VERSION 5.00
Begin VB.Form MChat 
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   3240
      Width           =   6090
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   3780
      Left            =   0
      Top             =   0
      Width           =   6480
   End
End
Attribute VB_Name = "MChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormOnTop Me    'sets the form on top
Me.Left = FrmMain.Left - ((Me.Width - FrmMain.Width) / 2)   'centers the form on the main form
Me.Top = FrmMain.Top + FrmMain.Height   'places the form on the bottom of the main form
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then   'checks to see if the user hit enter
    If Text2 = "" Then Exit Sub 'if the text is blank then it exits the sub
    ChatSend Text2  'if it's not blank it sends the chat
    Text2 = ""  'sets text 2 to nothing so the user can continue typing
End If  'ends the if
End Sub
Private Sub Timer1_Timer()
Text1 = ""  'clears text1
End Sub
