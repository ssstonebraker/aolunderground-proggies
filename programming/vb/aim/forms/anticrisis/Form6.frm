VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   165
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1935
   LinkTopic       =   "Form6"
   ScaleHeight     =   165
   ScaleWidth      =   1935
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu ONE 
         Caption         =   "Send To Chat"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ONE_Click()
Call Chat_Send(text$)
End Sub
