VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About prjRandom"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   2100
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "About"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Label Label2 
         Caption         =   $"frmAbout.frx":0000
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmAbout.Hide
End Sub
