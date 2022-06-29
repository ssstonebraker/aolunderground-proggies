VERSION 5.00
Begin VB.Form FrmExmp1 
   Caption         =   "Press the new i-face button"
   ClientHeight    =   3195
   ClientLeft      =   6795
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmExmp1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.ListBox List1 
      Height          =   735
      ItemData        =   "FrmExmp1.frx":0CCA
      Left            =   2760
      List            =   "FrmExmp1.frx":0CD7
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton OptMove 
      Caption         =   "KJL"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New I-Face"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox TxtExmp1 
      Alignment       =   2  'Center
      Height          =   765
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmExmp1.frx":0CF1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmExmp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
FrmMove1.Show
FrmCstmIt1.Show
End Sub

Private Sub Form_Load()
UpdateSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

