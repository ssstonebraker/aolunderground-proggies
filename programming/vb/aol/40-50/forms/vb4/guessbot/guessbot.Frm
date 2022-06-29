VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThaAbbot Guess Bot"
   ClientHeight    =   1365
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2745
   Height          =   1770
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   2865
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear List"
      Height          =   240
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Scroll Winners"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1070
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OFF"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ON"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Winners:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "The Number:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "none"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Albertus Extra Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize Timer
thenum$ = Int(Rnd * 100 + 1)
Label1.Caption = thenum$
SendChat "ThaAbbot Guess Bot enabled"
Timeout (0.7)
SendChat "Type ""/guess"" plus a number from 1-100"
Timer1.Enabled = True
End Sub


Private Sub Command2_Click()
Timer1.Enabled = False
Label1.Caption = "none"
SendChat "ThaAbbot Guess Bot Disabled"
End Sub


Private Sub Command3_Click()
SendChat "<B></I><U>ThaAbbot Guess Bot Winners:"
Timeout (0.8)
If List1.ListCount = 0 Then
Do: DoEvents: Loop
End If
List1.Enabled = False
i = List1.ListCount - 1
List1.ListIndex = 0
For X = 0 To i
List1.ListIndex = X
SendChat "<B></I></U>" + List1 + ""
Timeout (0.8)
Next X
List1.Enabled = True
Timeout (0.7)
SendChat "<B></I></U>That's all the winners!"
End Sub

Private Sub Command4_Click()
List1.Clear
End Sub

Private Sub Form_Load()
StayOnTop Me
MsgBox "ThaAbbot Guess Bot by ThaAbbot" + Chr$(13) + Chr$(10) + "©1998 ThaAbbot" + Chr$(13) + Chr$(10) + "Anyone can use these codes if you give me credit!" + Chr$(13) + Chr$(10) + "Please don't be lame and take the code and say they're yours" + Chr$(13) + Chr$(10) + "Dedicated to ProGs InC" + Chr$(13) + Chr$(10) + "http://come.to/progsinc", 0, "ThaAbbot Guess Bot by ThaAbbot"
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()
it$ = Label1.Caption
If LastChatLine Like ("/guess *") Then
per$ = SNFromLastChatLine
GoTo send
End If
Exit Sub
send:
If Text1 = LastChatLineWithSN Then
Exit Sub
End If
Number$ = Mid(LastChatLine, InStr(LastChatLine, "/") + 7)
If it$ > Number$ Then
SendChat ("" + per$ + " Guess higher!")
End If
If it$ < Number$ Then
SendChat ("" + per$ + " Guess lower!")
End If
If Number$ = it$ Then
SendChat ("" + per$ + " That's right!")
List1.AddItem per$
For i = 0 To List1.ListCount - 1
For E = 0 To List1.ListCount - 1

If LCase(List1.List(i)) Like LCase(List1.List(E)) And i <> E Then

List1.RemoveItem (E)
End If

Next E
Next i
End If
Text1.Text = LastChatLineWithSN
End Sub


