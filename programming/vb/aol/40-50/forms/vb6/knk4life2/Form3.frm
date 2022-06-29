VERSION 5.00
Begin VB.Form frmRoombust 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Room Buster"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1620
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Attemps -"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmRoombust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
Label1.Caption = "0"
Call CloseWindow(FindRoom&)
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False


End Sub

Private Sub Command3_Click()
Combo3.AddItem Combo3.Text

End Sub

Private Sub Command4_Click()
Combo3.RemoveItem Combo3.ListIndex

'Call SaveComboBox(App.Path + "\room.lst", rmNames)
End Sub

Private Sub Form_Load()
FormOnTop Me
On Local Error Resume Next
Call LoadComboBox(App.Path + "\room.lst", Combo3)
If Err Then
Beep
'MsgBox "Error: File Not Found!", vbExclamation, "Error"
End If
Timer2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SaveComboBox(App.Path + "\room.lst", Combo3)
End Sub

Private Sub Timer1_Timer()
Timer2.Enabled = True
Call Keyword("aol://2719:2-2-" & Combo3)
Label1.Caption = Val(Label1.Caption) + 1
If FindRoom& <> 0 Then Timer1.Enabled = False & Timer2.Enabled = False
End Sub

Private Sub Timer2_Timer()
Call WaitForOKOrRoom(Combo3)
End Sub
