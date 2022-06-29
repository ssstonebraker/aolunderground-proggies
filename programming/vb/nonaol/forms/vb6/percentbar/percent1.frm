VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "hope this works"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.ForeColor = RGB(0, 0, 255)     'use blue bar
For i = 0 To 100 Step 2
Pause (0.05)
updateprogress Picture1, i
Next
Picture1.Cls 'clear
End
End Sub
Sub updateprogress(pb As Control, ByVal percent)
Dim num$    'use percent
    If Not pb.AutoRedraw Then
    pb.AutoRedraw = -1
    End If
    pb.Cls
    pb.ScaleWidth = 100
    pb.DrawMode = 10
    num$ = Format$(percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$   'print percent
    pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
    pb.Refresh
End Sub

