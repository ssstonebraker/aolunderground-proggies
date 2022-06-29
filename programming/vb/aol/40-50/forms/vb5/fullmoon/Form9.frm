VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Server"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   3300
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   6240
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   6240
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6240
      Top             =   960
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   21
      Text            =   "Sn Here"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   4560
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6240
      Top             =   240
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ListBox lstList 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox lstmail 
      Height          =   1425
      Left            =   11640
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5760
      TabIndex        =   24
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Banned:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Finds:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Requests:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Lists:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Send List"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Make List"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Moon Server"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
End Sub

Private Sub Form_Load()

StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label16_Click()
Call SendChat("+Ttext1+")
End Sub

Private Sub Label17_Click()
Unload Me
End Sub

Private Sub Label18_Click()
Form9.WindowState = 1
End Sub

Private Sub Label2_Click()
Call MakeList
End Sub

Private Sub Label5_Click()
Form2.PopupMenu Form2.Label8, 2
'This one uses a 1 since its the first pull down menue on Form2Label6.ForeColor = QBColor(12)
End Sub

Private Sub Label6_Click()
If lstList.ListCount = 0 Then Exit Sub
who$ = ""
For X = 0 To lstList.ListCount - 1
who$ = who$ + "((" + lstList.List(X) + ")),"
Next X
who$ = Left$(who$, Len(who$) - 1)
lstList.Clear
Call ChatSend("Processing Requested Lists")
Call MailSend(who$, "Full Moon Server List", maillist$)
Call ChatSend("Requested Lists Sent")

End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Label8.Caption = Str(lstList.ListCount)
End Sub

Private Sub Timer2_Timer()
Label11.Caption = Str(List1.ListCount)
End Sub

Private Sub Timer3_Timer()
Label12.Caption = Str(List2.ListCount)
End Sub

Private Sub Timer4_Timer()
Label15.Caption = Str(List3.ListCount)
End Sub
