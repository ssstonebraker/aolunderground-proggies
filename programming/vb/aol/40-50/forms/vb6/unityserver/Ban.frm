VERSION 5.00
Begin VB.Form Ban 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Ban.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   1620
      TabIndex        =   1
      Top             =   225
      Width           =   1410
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1980
      MouseIcon       =   "Ban.frx":2BD5
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1215
      Width           =   675
   End
   Begin VB.Label startbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   675
      MouseIcon       =   "Ban.frx":2D27
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1215
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   0
      Picture         =   "Ban.frx":2E79
      Top             =   0
      Width           =   3240
   End
End
Attribute VB_Name = "Ban"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormOnTop Me
AddRoomToListBox List1, False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub
Private Sub Label1_Click()
AddRoomToListBox List1, False
End Sub
Private Sub List1_DblClick()
On Error Resume Next
For x = 0 To List2.ListCount
    If List2.List(x) = List1.List(List1.ListIndex) Then Exit Sub
Next x
List2.AddItem List1.List(List1.ListIndex)
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & List1.List(List1.ListIndex) & ", You Have Been Banned<font face=" & Chr(34) & "verdana" & Chr(34) & ">" & RAscii
End Sub
Private Sub List2_DblClick()
On Error Resume Next
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & List2.List(List2.ListIndex) & ", You Have Been UnBanned<font face=" & Chr(34) & "verdana" & Chr(34) & ">" & RAscii
List2.RemoveItem List2.ListIndex
End Sub
Private Sub startbut_Click()
Me.Hide
End Sub
