VERSION 5.00
Begin VB.Form frmTricks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "finding variables, etc."
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6495
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
   ScaleHeight     =   2415
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStuff 
      Height          =   1095
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtStuff 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtInfo 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmTricks.frx":0000
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
      Begin VB.TextBox txtCode 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmTricks.frx":0153
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&file"
      Begin VB.Menu mnuExit 
         Caption         =   "&exit"
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "&show"
      Begin VB.Menu mnuProcedureTitle 
         Caption         =   "&procedure title"
      End
      Begin VB.Menu mnuComments 
         Caption         =   "&comments"
      End
      Begin VB.Menu mnuVariables 
         Caption         =   "&variables"
      End
      Begin VB.Menu mnuArguments 
         Caption         =   "&arguments"
      End
   End
End
Attribute VB_Name = "frmTricks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuArguments_Click()
    txtStuff.Text = GetArguments(txtCode.Text)
    fraStuff.Caption = "arguments"
End Sub

Private Sub mnuComments_Click()
    txtStuff.Text = GetComments(txtCode.Text)
    fraStuff.Caption = "comments"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuProcedureTitle_Click()
    txtStuff.Text = GetSubTitle(txtCode.Text)
    fraStuff.Caption = "procedure title"
End Sub

Private Sub mnuVariables_Click()
    txtStuff.Text = GetVariables(txtCode.Text)
    fraStuff.Caption = "variables"
End Sub
