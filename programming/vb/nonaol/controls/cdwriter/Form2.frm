VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   345
      Left            =   5190
      TabIndex        =   2
      Top             =   2235
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "StartScan"
      Height          =   360
      Left            =   525
      TabIndex        =   1
      Top             =   2220
      Width           =   1260
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   195
      TabIndex        =   0
      Top             =   315
      Width           =   6675
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.VFcdwriter1.ScanSCSIBus 0

End Sub

Private Sub Command2_Click()
Unload Form2
End Sub
