VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menu 
      Caption         =   "menu"
      Begin VB.Menu mypc 
         Caption         =   "&My PC"
      End
      Begin VB.Menu desktop 
         Caption         =   "&Desktop"
      End
      Begin VB.Menu mp3stylist 
         Caption         =   "&Mp3 Stylist"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub desktop_Click()

End Sub

Private Sub mp3stylist_Click()
Shell (App.Path & "\mp3sylist.exe")
End Sub

Private Sub mypc_Click()
CommonDialog1.Filter = "Excutable Files | *.exe"
CommonDialog1.ShowOpen
FileName = CommonDialog1.FileName
If CommonDialog1.FileName = "" Then
End
End If
Shell (FileName)
End Sub
