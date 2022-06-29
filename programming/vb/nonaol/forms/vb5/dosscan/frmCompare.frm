VERSION 5.00
Begin VB.Form frmCompare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "comparing procedures"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtInfo 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmCompare.frx":0000
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "my chatsend"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtCode 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmCompare.frx":065E
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "copy of chatsend"
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox txtCopied 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmCompare.frx":0833
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   4800
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
      Begin VB.CommandButton cmdCompare 
         Caption         =   "compare"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdReplaceTitle 
         Caption         =   "replace title"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdKillDims 
         Caption         =   "kill dims"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdReplaceArguments 
         Caption         =   "replace arg."
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDoAll 
         Caption         =   "do all"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdReplaceVariables 
         Caption         =   "replace var."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdKillComments 
         Caption         =   "kill comments"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdKillTabs 
         Caption         =   "kill tabs"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&file"
      Begin VB.Menu mnuExit 
         Caption         =   "&exit"
      End
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCompare_Click()
    Dim CodeStr As String, CopyStr As String, IsSame As Boolean
    CodeStr$ = LCase(txtCode.Text)
    CopyStr$ = LCase(txtCopied.Text)
    CodeStr$ = Replace(CodeStr$, "&", "*")
    CodeStr$ = Replace(CodeStr$, "[", "*")
    CodeStr$ = Replace(CodeStr$, "]", "*")
    CodeStr$ = Replace(CodeStr$, "$", "*")
    CodeStr$ = Replace(CodeStr$, "%", "*")
    CodeStr$ = Replace(CodeStr$, "@", "*")
    CodeStr$ = Replace(CodeStr$, "!", "*")
    CodeStr$ = Replace(CodeStr$, "#", "*")
    CopyStr$ = Replace(CopyStr$, "&", "")
    CopyStr$ = Replace(CopyStr$, "[", "")
    CopyStr$ = Replace(CopyStr$, "]", "")
    CopyStr$ = Replace(CopyStr$, "$", "")
    CopyStr$ = Replace(CopyStr$, "%", "")
    CopyStr$ = Replace(CopyStr$, "@", "")
    CopyStr$ = Replace(CopyStr$, "!", "")
    CopyStr$ = Replace(CopyStr$, "#", "")
    CopyStr$ = Replace(CopyStr$, "**", "*")
    IsSame = CopyStr$ Like CodeStr$
    If CodeStr$ = CopyStr$ Or IsSame = True Then
        MsgBox "code scans copied"
    Else
        MsgBox "code does not scan copied"
    End If
End Sub

Private Sub cmdDoAll_Click()
    txtCode.Text = ReplaceTitle(txtCode.Text)
    txtCopied.Text = ReplaceTitle(txtCopied.Text)
    txtCode.Text = KillTabs(txtCode.Text)
    txtCopied.Text = KillTabs(txtCopied.Text)
    txtCode.Text = KillComments(txtCode.Text)
    txtCopied.Text = KillComments(txtCopied.Text)
    txtCode.Text = ReplaceVariables(txtCode.Text)
    txtCopied.Text = ReplaceVariables(txtCopied.Text)
    txtCode.Text = ReplaceArguements(txtCode.Text)
    txtCopied.Text = ReplaceArguements(txtCopied.Text)
    txtCode.Text = KillDims(txtCode.Text)
    txtCopied.Text = KillDims(txtCopied.Text)
End Sub

Private Sub cmdKillComments_Click()
    txtCode.Text = KillComments(txtCode.Text)
    txtCopied.Text = KillComments(txtCopied.Text)
End Sub

Private Sub cmdKillDims_Click()
    txtCode.Text = KillDims(txtCode.Text)
    txtCopied.Text = KillDims(txtCopied.Text)
End Sub

Private Sub cmdKillTabs_Click()
    txtCode.Text = KillTabs(txtCode.Text)
    txtCopied.Text = KillTabs(txtCopied.Text)
End Sub

Private Sub cmdReplaceArguments_Click()
    txtCode.Text = ReplaceArguements(txtCode.Text)
    txtCopied.Text = ReplaceArguements(txtCopied.Text)
End Sub

Private Sub cmdReplaceTitle_Click()
    txtCode.Text = ReplaceTitle(txtCode.Text)
    txtCopied.Text = ReplaceTitle(txtCopied.Text)
End Sub

Private Sub cmdReplaceVariables_Click()
    txtCode.Text = ReplaceVariables(txtCode.Text)
    txtCopied.Text = ReplaceVariables(txtCopied.Text)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
