VERSION 5.00
Begin VB.Form ip 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   615
   ClientLeft      =   1755
   ClientTop       =   1425
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "ip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormTop Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Label1.Caption = "Your Ip Is " & main.win.LocalIP
End Sub
