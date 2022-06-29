VERSION 5.00
Begin VB.Form Snap 
   Caption         =   "SnapShot"
   ClientHeight    =   4464
   ClientLeft      =   2004
   ClientTop       =   2244
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4464
   ScaleWidth      =   6480
   Begin VB.PictureBox Picture1 
      Height          =   732
      Left            =   0
      ScaleHeight     =   684
      ScaleWidth      =   2124
      TabIndex        =   0
      Top             =   0
      Width           =   2172
   End
End
Attribute VB_Name = "Snap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Resize()
    Picture1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


