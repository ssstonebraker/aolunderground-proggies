VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Macro Shop"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   5355
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   -240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.ListBox List1 
      Columns         =   99
      Height          =   645
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label6 
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
      Left            =   960
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Macro Shop"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines used:"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   855
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
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
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
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For X% = 32 To 255
List1.AddItem Chr(X%)
Next X%
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Form4.WindowState = 1
End Sub

Private Sub Label5_Click()
Form6.Show
End Sub

Private Sub Label6_Click()
Form2.PopupMenu Form2.Label7, 2
'This one uses a 1 since its the first pull down menue on Form2Label6.ForeColor = QBColor(12)
End Sub

Private Sub Label8_Click()
Form2.PopupMenu Form2.Label6, 2
'This one uses a 1 since its the first pull down menue on Form2Label6.ForeColor = QBColor(12)
End Sub




Private Sub Label9_Click()

End Sub

Private Sub List1_Click()
Text1.SelText = List1.List(List1.ListIndex)
End Sub

Private Sub Text1_Change()
Dim lineCount As Long
    On Local Error Resume Next
    'get/show the number of lines in the edit control
    lineCount = SendMessageLong(Text1.hwnd, EM_GETLINECOUNT, 0&, 0&)
    Label1 = Format$(lineCount, "##,###")
End Sub
