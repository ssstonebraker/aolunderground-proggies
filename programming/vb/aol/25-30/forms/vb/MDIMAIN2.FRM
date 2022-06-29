VERSION 2.00
Begin MDIForm frmMain 
   Caption         =   "MDIForm1"
   ClientHeight    =   5820
   ClientLeft      =   1530
   ClientTop       =   2235
   ClientWidth     =   7365
   Height          =   6510
   Left            =   1470
   LinkTopic       =   "MDIForm1"
   Top             =   1605
   Width           =   7485
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFNew 
         Caption         =   "&New"
      End
      Begin Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin Menu mnuWindow 
      Caption         =   "&Window"
      Begin Menu mnuTile 
         Caption         =   "&Tile"
      End
      Begin Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin Menu mnuWDash1 
         Caption         =   "-"
      End
      Begin Menu mnuWList 
         Caption         =   "Windows"
         WindowList      =   -1  'True
      End
   End
End
Option Explicit

Sub mnuArrangeIcons_Click ()

    frmMain.Arrange 3

End Sub

Sub mnuCascade_Click ()

    frmMain.Arrange 0

End Sub

Sub mnuFNew_Click ()

    Dim x As New frmCHild

    x.Show


End Sub

Sub mnuTile_Click ()

    frmMain.Arrange 1

End Sub

