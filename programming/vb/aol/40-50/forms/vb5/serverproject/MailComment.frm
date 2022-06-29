VERSION 5.00
Begin VB.Form MailComment 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Mail Comments"
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   20
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4395
      TabIndex        =   6
      Top             =   2280
      WhatsThisHelpID =   20
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   120
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1680
      WhatsThisHelpID =   20
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      WhatsThisHelpID =   20
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      WhatsThisHelpID =   20
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      WhatsThisHelpID =   20
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      WhatsThisHelpID =   20
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-•´)•–  X-Treme Server '99  Mail Comments -•´)•– "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      WhatsThisHelpID =   20
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   -120
      Picture         =   "MailComment.frx":0000
      Top             =   360
      WhatsThisHelpID =   20
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   0
      Picture         =   "MailComment.frx":054D
      Top             =   1440
      WhatsThisHelpID =   20
      Width           =   2430
   End
   Begin VB.Label loaded 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      WhatsThisHelpID =   20
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2040
      Picture         =   "MailComment.frx":0CB9
      Top             =   2040
      WhatsThisHelpID =   20
      Width           =   1500
   End
End
Attribute VB_Name = "MailComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
MenuForm.itemComments.Checked = True
mComm$ = Text1
If Text3.Text = "" Then mChatText = False
mChat$ = Text3
If Text1 = "" Then MenuForm.itemComments.Checked = False
Unload Me
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Unload Me
End Sub

Private Sub Command3_Click()
Text1.SetFocus
Text1 = ""
Text3 = ""
End Sub

Private Sub Form_Load()
StayOnTop Me
loaded.Caption = Str(Server.List2.ListCount)
CenterForm Me
End Sub

