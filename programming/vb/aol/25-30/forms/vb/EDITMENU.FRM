VERSION 2.00
Begin Form frmMain 
   Caption         =   "Menu Shortcuts"
   ClientHeight    =   2535
   ClientLeft      =   2955
   ClientTop       =   4155
   ClientWidth     =   5535
   Height          =   3225
   Left            =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5535
   Top             =   3525
   Width           =   5655
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFOPen 
         Caption         =   "&Open"
      End
      Begin Menu mnuFSave 
         Caption         =   "&Save"
      End
      Begin Menu mnuFSaveAs 
         Caption         =   "Save &As"
      End
   End
   Begin Menu mnuEdit 
      Caption         =   "&Edit"
      Begin Menu mnuECut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin Menu mnuECopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin Menu mnuEPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Option Explicit

Sub mnuECopy_Click ()

    msgbox "You have selected Copy from the Edit menu"

End Sub

Sub mnuECut_Click ()

    msgbox "You selected Cut from the Edit menu"

End Sub

Sub mnuEPaste_Click ()

    MsgBox "You have selected Paste from the Edit menu"

End Sub

