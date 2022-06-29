VERSION 5.00
Begin VB.Form Create 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2808
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   6504
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Create.frx":0000
   ScaleHeight     =   2808
   ScaleWidth      =   6504
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3960
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   2400
   End
   Begin VB.ListBox List5 
      Height          =   2352
      ItemData        =   "Create.frx":460C
      Left            =   4440
      List            =   "Create.frx":460E
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.ListBox List4 
      Height          =   2352
      ItemData        =   "Create.frx":4610
      Left            =   4440
      List            =   "Create.frx":4612
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.ListBox List3 
      Height          =   2352
      ItemData        =   "Create.frx":4614
      Left            =   4440
      List            =   "Create.frx":4616
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.ListBox List2 
      Height          =   2352
      ItemData        =   "Create.frx":4618
      Left            =   4440
      List            =   "Create.frx":461A
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.ListBox List1 
      Height          =   2352
      ItemData        =   "Create.frx":461C
      Left            =   240
      List            =   "Create.frx":461E
      TabIndex        =   1
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   2400
      TabIndex        =   15
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "List's Made"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   132
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   132
      Left            =   4440
      TabIndex        =   13
      Top             =   120
      Width           =   252
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   372
      Left            =   4320
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   852
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   372
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   372
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   1692
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   372
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   3720
      TabIndex        =   4
      Top             =   2640
      Width           =   492
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   372
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "               Fast Server Toolz"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   4.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   132
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   2172
   End
End
Attribute VB_Name = "Create"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormOnTop Me
List2.Visible = True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label10_Click()
List3.Visible = False
List4.Visible = True
Label10.Visible = False
Label11.Visible = True
End Sub

Private Sub Label11_Click()
List4.Visible = False
List5.Visible = True
Label11.Visible = False
Label12.Visible = True
End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label2_Click()
If GetUser = "" Then
MsgBox "You Must Be Signed On To Use This"
Else
MailOp.Show
End If
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_Click()
If List2.Visible Then
List2.RemoveItem List2.ListIndex
Else
End If
If List3.Visible Then
List3.RemoveItem List3.ListIndex
Else
End If
If List4.Visible Then
List4.RemoveItem List4.ListIndex
Else
End If
If List5.Visible Then
List5.RemoveItem List5.ListIndex
Else
End If
End Sub

Private Sub Label5_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Label9.Caption = "0"
End Sub

Private Sub Label6_Click()
Timer1.Enabled = True
End Sub

Private Sub Label7_Click()
List2.Clear
List2.Visible = True

End Sub

Private Sub Timer1_Timer()
If List2.Visible = True Then
For a = 0 To List1.ListCount
Label9.Caption = Val(Label9) + 1
List2.AddItem (Label9.Caption & ") " & List1.List(a))
Next a
Pause (0.1)
Timer1.Enabled = False
End If
If Label9.Caption = "1000" Then
List2.Visible = False
List3.Visible = True
Else
End If
If List3.Visible = True Then
For a = 0 To List1.ListCount
Label9.Caption = Val(Label9) + 1
List3.AddItem (Label9.Caption & ") " & List1.List(a))
Next a
Pause (0.1)
Timer1.Enabled = False
End If
If Label9.Caption = "2000" Then
List3.Visible = False
List4.Visible = True
Else
End If
If List4.Visible = True Then
For a = 0 To List1.ListCount
Label9.Caption = Val(Label9) + 1
List4.AddItem (Label9.Caption & ") " & List1.List(a))
Pause (0.1)
Next a
Timer1.Enabled = False
End If
If Label9.Caption = "3000" Then
List4.Visible = False
List5.Visible = True
Else
End If
If List5.Visible = True Then
For a = 0 To List1.ListCount
Label9.Caption = Val(Label9) + 1
List5.AddItem (Label9.Caption & ") " & List1.List(a))
Pause (0.1)
Next a
Timer1.Enabled = False
End If
If Label9.Caption = "4000" Then
MsgBox "All Your List's Are Full"
Else
End If
End Sub

Private Sub Timer2_Timer()
If Label9.Caption = "0" Then
Label13.Caption = "No"
Else
End If
If Label9.Caption = "1" Then
Label13.Caption = "1"
Else
End If
If Label9.Caption = "2" Then
Label13.Caption = "1"
Else
End If
If Label9.Caption = "1001" Then
Label13.Caption = "2"
Else
End If
If Label9.Caption = "1002" Then
Label13.Caption = "2"
Else
End If
If Label9.Caption = "2001" Then
Label13.Caption = "3"
Else
End If
If Label9.Caption = "2002" Then
Label13.Caption = "3"
Else
End If
If Label9.Caption = "3001" Then
Label13.Caption = "All"
Else
End If
If Label9.Caption = "3002" Then
Label13.Caption = "All"
Else
End If
If List2.ListIndex > 0 Then
Label13.Caption = "1"
Else
End If
If List3.ListIndex > 0 Then
Label13.Caption = "2"
Else
End If
If List4.ListIndex > 0 Then
Label13.Caption = "3"
Else
End If
If List5.ListIndex > 0 Then
Label13.Caption = "All"
Else
End If
End Sub
