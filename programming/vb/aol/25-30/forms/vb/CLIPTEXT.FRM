VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Dealing with the clipboard"
   ClientHeight    =   3465
   ClientLeft      =   735
   ClientTop       =   3930
   ClientWidth     =   7365
   Height          =   4155
   Left            =   675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7365
   Top             =   3300
   Width           =   7485
   Begin PictureBox Picture1 
      Height          =   3105
      Left            =   225
      ScaleHeight     =   3075
      ScaleWidth      =   6870
      TabIndex        =   1
      Top             =   150
      Width           =   6900
      Begin TextBox txtData 
         Height          =   2835
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   6645
      End
   End
   Begin Menu mnuEdit 
      Caption         =   "&Edit"
      Begin Menu mnuECut 
         Caption         =   "&Cut"
      End
      Begin Menu mnuECopy 
         Caption         =   "C&opy"
      End
      Begin Menu mnuEPaste 
         Caption         =   "&Paste"
      End
   End
End
Option Explicit

Sub mnuECopy_Click ()

    Clipboard.Clear
    Clipboard.SetText txtData.SelText

End Sub

