VERSION 5.00
Begin VB.Form FrmMove1 
   Caption         =   "Move Things"
   ClientHeight    =   3195
   ClientLeft      =   1080
   ClientTop       =   555
   ClientWidth     =   4695
   Icon            =   "FrmMove1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -1320
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   -1000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New I-Face"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.OptionButton OptMove 
      Caption         =   "KJL"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   735
      ItemData        =   "FrmMove1.frx":0CCA
      Left            =   2760
      List            =   "FrmMove1.frx":0CD7
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox TxtExmp1 
      Alignment       =   2  'Center
      Height          =   765
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmMove1.frx":0CF1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmMove1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Intxoffset As Integer
Dim Intyoffset As Integer

Private Sub Command1_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "2"
Intxoffset = X
Intyoffset = Y
Else
Text1.Text = "1"
End If
End Sub

Private Sub Command1_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "1"
Command1.Move Command1.Left + X - Intxoffset, Command1.Top + Y - Intyoffset
End If
End Sub

Private Sub Form_Load()
UpSett
Text1.Text = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSett
End Sub

Private Sub List1_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "2"
Intxoffset = X
Intyoffset = Y
Else
Text1.Text = "1"
End If
End Sub

Private Sub List1_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "1"
List1.Move List1.Left + X - Intxoffset, List1.Top + Y - Intyoffset
End If
End Sub
Private Sub Timer1_Timer()
Text1.Text = "1"
End Sub

Private Sub TxtExmp1_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "2"
Intxoffset = X
Intyoffset = Y
Else
Text1.Text = "1"
End If
End Sub

Private Sub TxtExmp1_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "1"
TxtExmp1.Move TxtExmp1.Left + X - Intxoffset, TxtExmp1.Top + Y - Intyoffset
End If
End Sub

Private Sub OptMove_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "2"
Intxoffset = X
Intyoffset = Y
Else
Text1.Text = "1"
End If
End Sub

Private Sub OptMove_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "1" Then
Text1.Text = "1"
OptMove.Move OptMove.Left + X - Intxoffset, OptMove.Top + Y - Intyoffset
End If
End Sub
