VERSION 4.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Write and read .INI example By DeVil"
   ClientHeight    =   945
   ClientLeft      =   795
   ClientTop       =   4200
   ClientWidth     =   4380
   Height          =   1350
   Left            =   735
   LinkTopic       =   "Form4"
   ScaleHeight     =   945
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Top             =   3855
   Width           =   4500
   Begin VB.CommandButton Command2 
      Caption         =   "Read it"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "What it will change the .exm file to"
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write it"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Write and read .INI example"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
r% = WritePrivateProfileString("Example", "Textbox", Text1.Text, "C:\initexam.exm")
End Sub

Private Sub Command2_Click()
Text1.Text = GetFromINI("Example", "Textbox", "C:\initexam.exm")
End Sub


Private Sub Form_Load()
MsgBox "there will be a file named initexam.exm in you C:\ drive made for THIS example delete it but it will be made every time this starts. Kewl, eh?", 0, "DeViL'S Example"
MsgBox "Oh yes i forgot, when you write to a file see knkfounders.bas it will tell you but i added a note to it so take a peek", 0, "DeViL'S Example"
End Sub


