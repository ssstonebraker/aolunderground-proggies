VERSION 2.00
Begin MDIForm frmMain 
   Caption         =   "MDIForm1"
   ClientHeight    =   5820
   ClientLeft      =   2505
   ClientTop       =   2385
   ClientWidth     =   7365
   Height          =   6510
   Left            =   2445
   LinkTopic       =   "MDIForm1"
   Top             =   1755
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
End
Option Explicit

Sub mnuFNew_Click ()

    Dim OurNewForm As New frmChild
    OurNewForm.Show


End Sub

