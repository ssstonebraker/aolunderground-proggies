VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Warez Requestor"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1785
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   0
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   1320
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "ReVeNgE"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Item to request"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders")
TimeOut (0.15)
SendChat BlackGreenBlack("«-×´¯`°  Request Bot ")
TimeOut (0.15)
SendChat BlackGreenBlack("«-×´¯`°  Requested Item:") + Text1.text + ("")
TimeOut (0.15)
SendChat BlackGreenBlack("«-×´¯`°  Say '/I got it' if you have it")
Timer2 = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
SendChat BlackGreenBlack("«-×´¯`°  Request Bot: Stoped")
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
If LastChatLine = "/I got it" Then
    
    If List2.ListCount = 0 Then List2.AddItem SNFromLastChatLine
List1.AddItem SNFromLastChatLine

For i = 0 To List2.ListCount - 1
num = List2.List(i)
If num = SNFromLastChatLine Then Exit Sub
Next i
List1.AddItem SNFromLastChatLine
List2.AddItem SNFromLastChatLine
    
 End If
End Sub

Private Sub Timer2_Timer()
SendChat BlackGreenBlack("«-×´¯`°  KnK Founders")
TimeOut (0.15)
SendChat BlackGreenBlack("«-×´¯`°  Request Bot ")
TimeOut (0.15)
SendChat BlackGreenBlack("«-×´¯`°  Requested Item:") + Text1.text + ("")
TimeOut (0.15)
SendChat BlackGreenBlack("«-×´¯`°  Say '/I got it' if you have it")

End Sub

Private Sub Timer3_Timer()
If Timer1.Enabled = False Then Exit Sub
If List1.ListCount = 0 Then
End If
If List1.ListCount <> 0 Then
Dim i As Integer
For i = 0 To List1.ListCount - 1
SendChat ("" + List1.List(i) + "Can you Send?")
List1.RemoveItem i
Next i
 End If
 
End Sub
