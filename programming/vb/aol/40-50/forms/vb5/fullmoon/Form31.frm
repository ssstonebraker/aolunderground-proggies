VERSION 5.00
Begin VB.Form Form31 
   BorderStyle     =   0  'None
   Caption         =   "M-Chat"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   LinkTopic       =   "Form31"
   Picture         =   "Form31.frx":0000
   ScaleHeight     =   3810
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7800
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   7920
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   5760
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7800
      Top             =   600
   End
   Begin VB.TextBox Text4 
      Height          =   2895
      Left            =   10560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   2055
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
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
      Height          =   2295
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "People:"
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
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   6600
      TabIndex        =   12
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label5 
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
      Left            =   6240
      TabIndex        =   11
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label4 
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
      Left            =   6120
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label3 
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
      Left            =   7200
      TabIndex        =   8
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   7440
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Moon M-Chat"
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
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendChat("" + Text5 + "")
Text5 = ""
End Sub

Private Sub Form_Load()

 Call AddRoomToListBox(List1)
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Form31.WindowState = 1
End Sub

Private Sub Label4_Click()
Form2.PopupMenu Form2.Label9, 2
'This one uses a 1 since its the first pull down menue on Form2Label6.ForeColor = QBColor(12)
End Sub

Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

Text4.Text = LastChatLineWithSN
If Text3.Text = Text4.Text Then Exit Sub
Text2.Text = "" + Text2.Text + "" & Chr(13) & "" & Chr(10) & "" & LastChatLineWithSN & ""
Text3.Text = LastChatLineWithSN
End Sub

Private Sub Timer2_Timer()
Label6.Caption = Str(List1.ListCount)
End Sub
