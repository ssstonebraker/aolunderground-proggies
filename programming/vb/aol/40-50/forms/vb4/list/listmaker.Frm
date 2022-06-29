VERSION 4.00
Begin VB.Form Form3 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "List Maker"
   ClientHeight    =   4035
   ClientLeft      =   5760
   ClientTop       =   510
   ClientWidth     =   3705
   Height          =   4440
   Icon            =   "listmaker.frx":0000
   Left            =   5700
   LinkTopic       =   "Form3"
   Picture         =   "listmaker.frx":0442
   ScaleHeight     =   4035
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Top             =   165
   Width           =   3825
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   720
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   4200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Text            =   "-List Me"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5280
      Top             =   1560
   End
   Begin VB.TextBox Text4 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Text            =   "‹(®)›  Ravage Inc. List Maker  ‹(®)›"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Text            =   "Screen Name"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2520
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00400000&
      Caption         =   "Turn Off"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   600
      Top             =   4560
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.Label Label11 
      BackColor       =   &H00400000&
      Caption         =   "List Bot"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "People:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "(0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Room"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Send List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Name"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Name"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "List Maker"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If What_Said = Text1.text Then
List1.AddItem " " & Screen_Name & " "
Label9.Caption = Str(List1.ListCount)
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case 2: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case 3: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case 4: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case 5: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case 6: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case 7: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
Case Else: SendChat "‹(" + Label9 + ")›" & Screen_Name & " was added to the list‹(" + Label9 + ")›"
End Select
End If
End Sub


Private Sub Form_Load()
StayOnTop Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = QBColor(15)
Label2.ForeColor = QBColor(15)
Label1.ForeColor = QBColor(15)
Label5.ForeColor = QBColor(15)
Label6.ForeColor = QBColor(15)
Label7.ForeColor = QBColor(15)
Label3.ForeColor = QBColor(15)
End Sub

Private Sub Image1_Click()
Chat1.ScanOn
timer2.Enabled = True
Timer3.Enabled = False

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = QBColor(12)
End Sub

Private Sub Label12_Click()
Timer3.Enabled = True
timer2.Enabled = False
Chat1.ScanOff
LGreen_DGreen_LGreen "  ‹(®)›Ravage Inc. •List Maker•‹(®)›"
timeout (0.5)
LGreen_DGreen_LGreen "‹(®)›Ravage Inc. Say • Bot Off •‹(®)›"
End Sub

Private Sub Label2_Click()

Form3.WindowState = 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = QBColor(12)
End Sub

Private Sub Label3_Click()
SendChat "‹(®)›Ravage Inc. •List Maker•‹(®)›"
timeout (0.5)
SendChat "‹(®)›Unloaded BY •" + UserSN() + "‹(®)›"

Unload Me


End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = QBColor(12)
End Sub

Private Sub Label4_Click()
If Text2 = "" Then
MsgBox "You Dont Have A Name To Add", vbCritical, "Ravage Inc."
Else
List1.AddItem Text2
SendChat "‹(®)›" + Text2.text + " You Were Added‹(®)›"
End If
Text2 = ""
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = QBColor(12)
End Sub

Private Sub Label5_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = QBColor(12)
End Sub

Private Sub Label6_Click()
List1.clear
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = QBColor(12)
End Sub

Private Sub Label7_Click()
LGreen_DGreen_LGreen "‹(®)› Sending List ‹(®)›"
On Error Resume Next
For i = 0 To List1.ListCount - 1
names$ = names$ + "  " + "" + List1.List(i) + " , " + " "
Next i
Call SendMail(names$, text3, Text4)
timeout (0.5)
LGreen_DGreen_LGreen "‹(®)› List Sent ‹(®)›"
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = QBColor(12)
End Sub

Private Sub Label8_Click()
 Call AddRoomToListBox(List1)
LGreen_DGreen_LGreen "‹(®)›Was Added To List‹(®)›"
End Sub

Private Sub Timer1_Timer()
Label9.Caption = Str(List1.ListCount)
End Sub

Private Sub timer2_Timer()
LGreen_DGreen_LGreen "  ‹(®)›Ravage Inc. •List Maker•‹(®)›"
timeout (0.5)
LGreen_DGreen_LGreen "‹(®)›Ravage Inc. Say • " + Text1.text + "•‹(®)›"
timeout (60#)

End Sub


Private Sub Timer3_Timer()
timer2.Enabled = False
End Sub


