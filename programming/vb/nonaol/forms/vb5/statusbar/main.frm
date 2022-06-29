VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   5235
   ClientTop       =   4185
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4440
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'look under mousemove
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "You Moved onto Command1"
'This Changes Label1.Caption to what you want
End Sub

Private Sub Command2_Click()
'look under mousemove
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "You Moved onto Command2"
'This Changes Label1.Caption to what you want
End Sub

Private Sub Command3_Click()
'look under mousemove
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "You Moved onto Command3"
'This Changes Label1.Caption to what you want
End Sub

Private Sub Command4_Click()
'look under mousemove
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "You Moved onto The About Button" 'This Changes Label1.Caption to what you want
'This Changes Label1.Caption to what you want
End Sub

Private Sub Form_Load()
'look under mousemove
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Made By FueLx"
End Sub

Private Sub Frame1_Click()
'look under mousemove
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "The Status Bar"
End Sub

Private Sub Label1_Click()
'look under mousemove
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "The Status Bar"
'This Changes Label1.Caption to what you want
End Sub
