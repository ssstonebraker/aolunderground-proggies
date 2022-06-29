VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIntro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2355
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmIntro.frx":0000
      Top             =   60
      Width           =   4965
   End
   Begin VB.Label lblOkay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Okay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2085
      TabIndex        =   1
      Top             =   2445
      Width           =   915
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2730
      Left            =   0
      Top             =   0
      Width           =   5070
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub lblOkay_Click()
    Unload Me
End Sub
