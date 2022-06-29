VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "INH AFK"
   ClientHeight    =   2070
   ClientLeft      =   1005
   ClientTop       =   1515
   ClientWidth     =   2130
   FillColor       =   &H00FFFFFF&
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "AGaramond Bold"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Height          =   2475
   Left            =   945
   LinkTopic       =   "Form1"
   Picture         =   "INHAFK.frx":0000
   ScaleHeight     =   2070
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   2250
   Begin VB.ListBox List2 
      Height          =   315
      Left            =   3000
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   2400
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Interval        =   4
      Left            =   2640
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1080
   End
   Begin VB.ListBox List1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Reason:"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Msg's"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Start AFK"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mins."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "AGaramond Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "         INH AFK"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Call Center(Form1)
Call FormOnTop(Form1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
Label2.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
End Sub


Private Sub Form_Unload(Cancel As Integer)
Label9.Caption = "0"
Label4.Caption = "0"
List1.Clear
Text1.Text = "Reason:"
list2.Clear
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
aa = Button
If aa <> 0 Then
dd = (X - bb)
ee = (Y - cc)
Left = Left + (dd)
Top = Top + (ee)
Exit Sub
End If
bb = X
cc = Y
End Sub


Private Sub Label2_Click()
Form4.WindowState = 1
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H808080
Label7.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
End Sub


Private Sub Label3_Click()
End
End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H808080
Label2.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
End Sub


Private Sub Label6_Click()
If Label6.Caption = "Start AFK" Then
Timer2.Enabled = True
timer3.Enabled = True
BlackBlue ("`·. . -·´) INH •[AFK-On]•[Reason: " + Text1.Text + "]•")
 
BlackBlue ("  ¸,.–·´ •[" + LCase(UserSN) + "]•[AFK]•[" + Label4.Caption + "]•")
 
BlackBlue ("`·. . -·´) •[Type -" + LCase(UserSN) + " And a Message]•")
Label4.Caption = (Label4) + 1
Label6.Caption = "Stop AFK"
Exit Sub
End If
If Label6.Caption = "Stop AFK" Then
Timer2.Enabled = False
timer3.Enabled = False
Pause (0.01)
BlackBlue ("`·. . -·´) INH •[AFK-Off]•[Reason: " + Text1.Text + "]•")
Pause (0.01)
BlackBlue ("  ¸,.–·´ •[" + LCase(UserSN) + "]•[AFK]•[" + Label4.Caption + "]•")
Pause (0.01)
Label6.Caption = "Start AFK"
Do
DoEvents
Loop
Exit Sub
End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &H808080
Label2.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label7.ForeColor = &HFFFFFF
End Sub


Private Sub Label7_Click()
Form3.PopupMenu Form3.Options, 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &H808080
Label2.ForeColor = &HFFFFFF
Label3.ForeColor = &HFFFFFF
Label6.ForeColor = &HFFFFFF
End Sub


Private Sub List1_Click()
MsgBox "" + list2.List(List1.ListIndex) + "", , "Im Not Here AFK Message"
End Sub

Private Sub Timer1_Timer()
Label9.Caption = List1.ListCount
End Sub


Private Sub Timer2_Timer()
Do
Pause (60)

Pause (0.01)
BlackBlue ("`·. . -·´) INH •[AFK-On]•[Reason: " + Text1.Text + "]•")
Pause (0.01)
BlackBlue ("  ¸,.–·´ •[" + LCase(UserSN) + "]•[AFK]•[" + Label4.Caption + "]•")
Pause (0.01)
BlackBlue ("`·. . -·´) •[Type -" + LCase(UserSN) + " And a Message]•")
Label4.Caption = (Label4) + 1
Loop
End Sub


Private Sub timer3_Timer()
On Error Resume Next
FindChatRoom




If LCase(Label10.Caption) Like LCase("-" + LCase(UserSN) + "*") Then
blah$ = Len(Label10.Caption) - Len("" + LCase(UserSN) + "")
hah$ = Right(Label10.Caption, blah$)
For Z = 1 To 11
If Mid$(hah$, Z, 1) = " " Then
wow$ = Len(hah$) - Len("e")
message = Right(hah$, wow$)


End If
Next Z

list2.AddItem message
List1.AddItem SNFromLastChatLine
Pause (0.1)
BlackBlue ("`·. . -·´) INH •[AFK-On]•[Reason: " + Text1.Text + "]•")
Pause (0.01)
BlackBlue ("  ¸,.–·´ •[" + LCase(SNFromLastChatLine) + "]•[Message Recorded]•")
Pause (1)
 
End If
 
End Sub

Private Sub Timer4_Timer()
Label10.Caption = LastChatLine
Call AntiIdle4o
End Sub


