VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aol 7.0 Send Chat Example By Source"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Send Label Caption"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "source's aol 7.0 send chat example"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Textbox"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Fixed Text"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "contact on aim: cia source"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "source's aol 7.0 send chat example"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   1560
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendChat ("source's aol 7.0 send chat example")
End Sub

Private Sub Command2_Click()
SendChat ("" + Text1.Text + "")
End Sub

Private Sub Command3_Click()
SendChat ("" + Label1.Caption + "")
End Sub
