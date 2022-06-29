VERSION 5.00
Begin VB.Form FrmCstmIt1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costimize the I-Face"
   ClientHeight    =   1365
   ClientLeft      =   4320
   ClientTop       =   4830
   ClientWidth     =   2895
   Icon            =   "FrmCstmIt1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2895
   Begin VB.ComboBox CmbCmdBG1 
      Height          =   315
      ItemData        =   "FrmCstmIt1.frx":0CCA
      Left            =   1440
      List            =   "FrmCstmIt1.frx":0CD1
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox CmbTxtBG1 
      Height          =   315
      ItemData        =   "FrmCstmIt1.frx":0CE0
      Left            =   1440
      List            =   "FrmCstmIt1.frx":0CE7
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame FramCstm 
      Caption         =   "Options"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.ComboBox CmbTxtClr1 
         Height          =   315
         ItemData        =   "FrmCstmIt1.frx":0CF4
         Left            =   120
         List            =   "FrmCstmIt1.frx":0CFB
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CmbClrs1 
         Height          =   315
         ItemData        =   "FrmCstmIt1.frx":0D05
         Left            =   120
         List            =   "FrmCstmIt1.frx":0D0C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdUpdate1 
         Caption         =   "Update Settings"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmCstmIt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbClrs1_Click()
Select Case CmbClrs1.Text
Case "Red"
FrmMove1.BackColor = &HFF&
FrmMove1.OptMove.BackColor = &HFF&
Case "Green"
FrmMove1.BackColor = &HFF00&
FrmMove1.OptMove.BackColor = &HFF00&
Case "Blue"
FrmMove1.BackColor = &HFF0000
FrmMove1.OptMove.BackColor = &HFF0000
Case "Yellow"
FrmMove1.BackColor = &HFFFF&
FrmMove1.OptMove.BackColor = &HFFFF&
Case "Black"
FrmMove1.BackColor = &H0&
FrmMove1.OptMove.BackColor = &H0&
Case "Gray"
FrmMove1.BackColor = &HC0C0C0
FrmMove1.OptMove.BackColor = &HC0C0C0
Case "White"
FrmMove1.BackColor = &HFFFFFF
FrmMove1.OptMove.BackColor = &HFFFFFF
End Select
End Sub

Private Sub CmbClrs1_DropDown()
If CmbClrs1.ListCount = 1 Then
CmbClrs1.Clear
CmbClrs1.AddItem ("Red")
CmbClrs1.ListIndex = 0
CmbClrs1.AddItem ("Green")
CmbClrs1.AddItem ("Blue")
CmbClrs1.AddItem ("Yellow")
CmbClrs1.AddItem ("Black")
CmbClrs1.AddItem ("Gray")
CmbClrs1.AddItem ("White")
End If
End Sub

Private Sub CmbCmdBG1_Click()
Select Case CmbCmdBG1.Text
Case "Red"
FrmMove1.Command1.BackColor = &HFF&
Case "Green"
FrmMove1.Command1.BackColor = &HFF00&
Case "Blue"
FrmMove1.Command1.BackColor = &HFF0000
Case "Yellow"
FrmMove1.Command1.BackColor = &HFFFF&
Case "Black"
FrmMove1.Command1.BackColor = &H0&
Case "Gray"
FrmMove1.Command1.BackColor = &HC0C0C0
Case "White"
FrmMove1.Command1.BackColor = &HFFFFFF
End Select
End Sub

Private Sub CmbCmdBG1_DropDown()
If CmbCmdBG1.ListCount = 1 Then
CmbCmdBG1.Clear
CmbCmdBG1.AddItem ("Red")
CmbCmdBG1.ListIndex = 0
CmbCmdBG1.AddItem ("Green")
CmbCmdBG1.AddItem ("Blue")
CmbCmdBG1.AddItem ("Yellow")
CmbCmdBG1.AddItem ("Black")
CmbCmdBG1.AddItem ("Gray")
CmbCmdBG1.AddItem ("White")
End If
End Sub

Private Sub CmbTxtBG1_Click()
Select Case CmbTxtBG1.Text
Case "Red"
FrmMove1.TxtExmp1.BackColor = &HFF&
FrmMove1.List1.BackColor = &HFF&
Case "Green"
FrmMove1.TxtExmp1.BackColor = &HFF00&
FrmMove1.List1.BackColor = &HFF00&
Case "Blue"
FrmMove1.TxtExmp1.BackColor = &HFF0000
FrmMove1.List1.BackColor = &HFF0000
Case "Yellow"
FrmMove1.TxtExmp1.BackColor = &HFFFF&
FrmMove1.List1.BackColor = &HFFFF&
Case "Black"
FrmMove1.TxtExmp1.BackColor = &H0&
FrmMove1.List1.BackColor = &H0&
Case "Gray"
FrmMove1.TxtExmp1.BackColor = &HC0C0C0
FrmMove1.List1.BackColor = &HC0C0C0
Case "White"
FrmMove1.TxtExmp1.BackColor = &HFFFFFF
FrmMove1.List1.BackColor = &HFFFFFF
End Select
End Sub

Private Sub CmbTxtBG1_DropDown()
If CmbTxtBG1.ListCount = 1 Then
CmbTxtBG1.Clear
CmbTxtBG1.AddItem ("Red")
CmbTxtBG1.ListIndex = 0
CmbTxtBG1.AddItem ("Green")
CmbTxtBG1.AddItem ("Blue")
CmbTxtBG1.AddItem ("Yellow")
CmbTxtBG1.AddItem ("Black")
CmbTxtBG1.AddItem ("Gray")
CmbTxtBG1.AddItem ("White")
End If
End Sub

Private Sub CmbTxtClr1_Click()
Select Case CmbTxtClr1.Text
Case "Red"
FrmMove1.TxtExmp1.ForeColor = &HFF&
FrmMove1.OptMove.ForeColor = &HFF&
FrmMove1.List1.ForeColor = &HFF&
Case "Green"
FrmMove1.TxtExmp1.ForeColor = &HFF00&
FrmMove1.OptMove.ForeColor = &HFF00&
FrmMove1.List1.ForeColor = &HFF00&
Case "Blue"
FrmMove1.TxtExmp1.ForeColor = &HFF0000
FrmMove1.OptMove.ForeColor = &HFF0000
FrmMove1.List1.ForeColor = &HFF0000
Case "Yellow"
FrmMove1.TxtExmp1.ForeColor = &HFFFF&
FrmMove1.OptMove.ForeColor = &HFFFF&
FrmMove1.List1.ForeColor = &HFFFF&
Case "Black"
FrmMove1.TxtExmp1.ForeColor = &H0&
FrmMove1.OptMove.ForeColor = &H0&
FrmMove1.List1.ForeColor = &H0&
Case "Gray"
FrmMove1.TxtExmp1.ForeColor = &HC0C0C0
FrmMove1.OptMove.ForeColor = &HC0C0C0
FrmMove1.List1.ForeColor = &HC0C0C0
Case "White"
FrmMove1.TxtExmp1.ForeColor = &HFFFFFF
FrmMove1.OptMove.ForeColor = &HFFFFFF
FrmMove1.List1.ForeColor = &HFFFFFF
End Select
End Sub

Private Sub CmbTxtClr1_DropDown()
If CmbTxtClr1.ListCount = 1 Then
CmbTxtClr1.Clear
CmbTxtClr1.AddItem ("Red")
CmbTxtClr1.ListIndex = 0
CmbTxtClr1.AddItem ("Green")
CmbTxtClr1.AddItem ("Blue")
CmbTxtClr1.AddItem ("Yellow")
CmbTxtClr1.AddItem ("Black")
CmbTxtClr1.AddItem ("Gray")
CmbTxtClr1.AddItem ("White")
End If
End Sub

Private Sub CmdUpdate1_Click()
SaveSett
UpdateSettings
End Sub

Private Sub Form_Load()
CmbClrs1.ListIndex = 0
CmbTxtClr1.ListIndex = 0
CmbTxtBG1.ListIndex = 0
CmbCmdBG1.ListIndex = 0
End Sub
