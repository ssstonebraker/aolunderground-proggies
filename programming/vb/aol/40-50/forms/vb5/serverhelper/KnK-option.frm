VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1560
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   " AOL Version"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1575
      Begin VB.OptionButton Option4 
         Caption         =   "AOL 4.o"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "AOL95"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "No"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Yes"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Show Intro art and play wave at startup."
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Check2 = False

End Sub

Private Sub Check2_Click()
Check1 = False
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub
