VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Restarter 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Pw 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   765
      Width           =   3390
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox cPath 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   3390
   End
   Begin VB.Label OkBut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Left            =   495
      MouseIcon       =   "Restarter.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   810
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aol Dir:"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1170
      TabIndex        =   2
      Top             =   135
      Width           =   510
   End
   Begin VB.Label ChangeBut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
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
      Left            =   315
      MouseIcon       =   "Restarter.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   450
      Width           =   660
   End
End
Attribute VB_Name = "Restarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeBut_Click()
CommonDialog1.Filter = "Executable (*.Exe)|*.exe"   'sets the common dialogue filter to executables
CommonDialog1.ShowOpen  'tells common dialogue to show the form for opening a file
cPath = CommonDialog1.FileName  'sets the path to the file designated by the user
End Sub
Private Sub Form_Load()
FormOnTop Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub
Private Sub OkBut_Click()
AolRestartPw = Pw   'sets the variable to the users password
AolDir = cPath  'sets the path to the users path
Me.Hide 'hides the form
End Sub
