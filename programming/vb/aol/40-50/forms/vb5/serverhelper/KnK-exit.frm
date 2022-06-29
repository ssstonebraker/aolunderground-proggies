VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   0  'None
   Caption         =   "Form11"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   LinkTopic       =   "Form11"
   Picture         =   "KnK-exit.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Just kidding"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
KnKload$ = GetFromINI("KnKTheme", "KnKload", App.Path + "\KnK.ini")
If KnKload$ = "knk1" Then
Unload Me
Form7.Show
End If
If KnKload$ = "knk2" Then
Unload Me
Form13.Show

End If
End Sub

Private Sub Form_Load()
Playwav (App.Path + "\exit.WAV")
StayOnTop Me
End Sub
