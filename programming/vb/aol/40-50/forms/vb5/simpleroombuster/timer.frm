VERSION 5.00
Begin VB.Form time 
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   3030
   ClientTop       =   4800
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   1560
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   0
   End
End
Attribute VB_Name = "time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()

Call KeyWord("aol://2719:2-2-" & main.Combo1)
main.Label10.Caption = Val(main.Label10.Caption) + 1
If FindRoom& <> 0 Then Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Call WaitForOKOrRoom(main.Combo1)
End Sub
