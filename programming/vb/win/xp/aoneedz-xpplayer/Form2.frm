VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1590
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3120
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   195
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1350
      Left            =   1440
      MousePointer    =   1  'Arrow
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   195
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      TabIndex        =   3
      Top             =   500
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   3090
      TabIndex        =   1
      Top             =   20
      Width           =   3120
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xp amp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   -30
         Width           =   510
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   900
      TabIndex        =   6
      Top             =   1395
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   5
      Top             =   1400
      Width           =   300
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xp amp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
Me.File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
If GetFromINI("Settings", "Directory", "\player.oca") = "" Then
Load Me
Else
Form1.File1.Path = GetFromINI("Settings", "Directory", "\player.oca")
Me.Hide
Form1.Show
End If
StayOnTop Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormNoCaption Me
End Sub

Private Sub Label2_Click()
If Me.File1.ListCount = 0 Then
Exit Sub
Else
Call WriteToINI("Settings", "Directory", "" & Dir1.Path, "\player.oca")
Form1.File1.Path = GetFromINI("Settings", "Directory", "\player.oca")
If Form1.WindowState = 1 Then
Form2.Hide
Form1.WindowState = 0
Else
Me.Hide
Form1.Show
End If
End If
End Sub

Private Sub Label3_Click()
Me.WindowState = 1
Unload Me
Form1.WindowState = 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveFormNoCaption Me
End Sub
