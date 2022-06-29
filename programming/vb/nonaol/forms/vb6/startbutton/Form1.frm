VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Change"
   ClientHeight    =   960
   ClientLeft      =   5040
   ClientTop       =   4275
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   2175
   Begin VB.CommandButton Command3 
      Caption         =   "L.click"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "R.click"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "enter"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim shelltraywnd As Long, button As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call SendMessageByString(button, WM_SETTEXT, 0&, Text1.Text)

End Sub



Private Sub Command2_Click()
Dim shelltraywnd As Long, button As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call SendMessageLong(button, WM_RBUTTONDOWN, 0&, 0&)
Call SendMessageLong(button, WM_RBUTTONUP, 0&, 0&)
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Dim shelltraywnd As Long, button As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
Call SendMessageLong(button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(button, WM_LBUTTONUP, 0&, 0&)
End Sub

Private Sub Form_Load()
If Text1.Text = "" Then
MsgBox "Please enter a name in the textbox", vbOKOnly
End If
End Sub
