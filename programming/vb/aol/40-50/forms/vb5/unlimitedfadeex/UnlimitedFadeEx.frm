VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Unlimited Color Fade Example"
   ClientHeight    =   3744
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3744
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove Color"
      Height          =   615
      Left            =   3360
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Color"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1224
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      LargeChange     =   50
      Left            =   120
      Max             =   254
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   50
      Left            =   120
      Max             =   254
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   50
      Left            =   120
      Max             =   254
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy HTML"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Fade"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim C As Integer, X As Integer, T As Integer, P As Integer
Dim Wavey As Boolean
Dim B As String
Wavey = False
If List1.ListCount = 2 Then Text2.Text = ChatFade(Text1.Text, GetR(List1.List(0)), GetR(List1.List(1)), GetG(List1.List(0)), GetG(List1.List(1)), GetB(List1.List(0)), GetB(List1.List(1)), Wavey): Exit Sub
Text3.Text = Text1.Text
Do
C = Len(Text3.Text) / List1.ListCount - 1
If InStr(C, ".") = 0 Then GoTo NoDot:
Text3.Text = Text3.Text & " "
Loop
NoDot:
B = Text3.Text
T = (Len(Text3) / (List1.ListCount - 1))
Text2.Enabled = True
For X = 0 To List1.ListCount - 2
If X = 0 Then P = 1: GoTo Uno
P = X * T
Uno:
B = Mid(Text3.Text, P, T)
Text2.Text = Text2.Text & ChatFade(B, GetR(List1.List(X)), GetR(List1.List(X + 1)), GetG(List1.List(X)), GetG(List1.List(X + 1)), GetB(List1.List(X)), GetB(List1.List(X + 1)), Wavey)
Next
Text2.Enabled = False
End Sub

Private Sub Command2_Click()
Text2.Enabled = True
Clipboard.SetText Text2.Text
Text2.Enabled = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
List1.AddItem Color(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub Command5_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Form_Load()

End Sub

Private Sub HScroll1_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll1_Scroll()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

End Sub

Private Sub HScroll2_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

End Sub

Private Sub HScroll2_Scroll()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

End Sub

Private Sub HScroll3_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

End Sub

Private Sub HScroll3_Scroll()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

End Sub

Private Sub List1_DblClick()
Dim A, C, D, R, G, B
A = List1.List(List1.ListIndex)
C = InStr(A, " ")
R = Left(A, C - 1)
A = Mid(A, C + 1)
C = InStr(A, " ")
G = Left(A, C - 1)
B = Mid(A, C + 1)

Label1.BackColor = RGB(R, G, B)

End Sub
