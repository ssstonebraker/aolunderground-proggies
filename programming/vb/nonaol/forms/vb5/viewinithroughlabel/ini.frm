VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Give it a second to work."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Air = FreeFile
Open "c:\windows\win.ini" For Input As Air
Do Until EOF(Air)
Input #Air, x
Label1.Caption = Label1.Caption & x & Chr(13)
DoEvents
Loop


End Sub

Private Sub Timer1_Timer()

End Sub
