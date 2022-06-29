VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Set To Guest"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Set To Guest"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1155
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   $"Frmguest.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
Call GuestSetToGuest
End Sub

