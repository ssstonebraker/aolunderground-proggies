VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbFile 
      Height          =   315
      ItemData        =   "cmb.frx":0000
      Left            =   0
      List            =   "cmb.frx":000A
      TabIndex        =   0
      Text            =   "File"
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbFile_Click()
If cmbFile.Text = "" Then
  cmbFile.Text = "File"
End If
' have ALL the words u want listed n the combo box n the propertys box....THEN use thise 2 code
' them
If cmbFile.Text = "Bla" Then
'  wut is gonna do!
End If
If cmbFile.Text = "Bla2" Then
' the next thing the coding is gonna do
End If
End Sub
'continue the rest how ya want

