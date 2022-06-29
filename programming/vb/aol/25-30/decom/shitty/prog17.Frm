VERSION 4.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   4140
   ClientLeft      =   1455
   ClientTop       =   2250
   ClientWidth     =   5505
   Height          =   4830
   Left            =   1395
   LinkTopic       =   "Form16"
   ScaleHeight     =   4140
   ScaleWidth      =   5505
   Top             =   1620
   Width           =   5625
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuKill 
         Caption         =   "Kill Wait"
      End
   End
End
Attribute VB_Name = "Form16"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Rof% = FindWindow("Thunderform", "Form16")
Call ParentChange(AOL%, Rof%)
End Sub

Private Sub mnuKill_Click()
Screen.MousePointer = 11
End Sub


