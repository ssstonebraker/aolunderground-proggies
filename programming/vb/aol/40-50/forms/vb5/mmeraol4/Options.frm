VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Mailbox"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000B&
         Caption         =   "Old Mail"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000B&
         Caption         =   "Sent Mail"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000B&
         Caption         =   "New Mail"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000B&
         Caption         =   "FlashMail"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remove ""fwd: """
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
End Sub
Private Sub Form_Load()
Call StayOnTop(Me)
End Sub
