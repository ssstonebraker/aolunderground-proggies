VERSION 5.00
Object = "{92EDEF56-A415-11D2-BBA6-BA26EE701995}#4.0#0"; "QUIRKAIM.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "beavs team scrambler"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbOff 
      Caption         =   "off"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "on"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstBlue 
      Height          =   1065
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox lstRed 
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin QuirkAIMOCX.AIM qrkAIM 
      Left            =   1200
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblBlueScr 
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblRedScr 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblScores 
      Alignment       =   2  'Center
      Caption         =   "scores"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lblOutput 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblWord 
      Caption         =   "word"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblBlueCnt 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblRedCnt 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbOff_Click()
qrkAIM.ChatOff
qrkAIM.ChatSend ("team scramble is over, red team has " & lblRedScr & " points, and blue has " & lblBlueScr & " points")
lstRed.Clear
lblRedCnt = "0"
lstBlue.Clear
lblBlueCnt = "0"
txtWord.Enabled = True
End Sub

Private Sub cmdOn_Click()
qrkAIM.ChatOn
lblRedScr = "0"
lblBlueScr = "0"
lstRed.Clear
lblRedCnt = "0"
lstBlue.Clear
lblBlueCnt = "0"
qrkAIM.ChatSend ("team scrambler is now ON!")
End Sub

Private Sub Command1_Click()
txtWord = Reverse("monkey")
End Sub

Private Sub Form_Load()
qrkAIM.ChatSend ("team scrambler, type " & Chr(34) & "/red" & Chr(34) & " to join red team, and " & Chr(34) & "/blue" & Chr(34) & " to join the blue team")
txtWord.Text = " "
End Sub

Private Sub qrkAIM_ChatLastLine(Who As String, What As String)
strWhat$ = Left(What$, 4)
If strWhat$ = "/red" Then
Call Remove(lstBlue, Who$)
Call Add(lstRed, Who$)
lblRedCnt = lstRed.ListCount
Call qrkAIM.ChatSend(Who$ & " has joined the red team, they have " & lblRedScr & " points, and " & lblRedCnt & " players")
ElseIf strWhat$ = "/blu" Then
Call Remove(lstRed, Who$)
Call Add(lstBlue, Who$)
Call qrkAIM.ChatSend(Who$ & " has joined the blue team, they have " & lblBlueScr & " points, and " & lblBlueCnt & " players")
lblBlueCnt = lstBlue.ListCount
End If
If What$ = txtWord Then
strWinner$ = WhichTeam(lstRed, lstBlue, Who$)
If strWinner$ = "1" Then
 qrkAIM.ChatSend (Who$ & " of the red team, you got it write! it was " & txtWord)
 lblRedScr = Val(lblRedScr) + 5
 MsgBox "put a new word and type enter"
 txtWord.Enabled = True
 txtWord.Text = " "
End If
If strWinner$ = "2" Then
 qrkAIM.ChatSend (Who$ & " of the blue team, you got it write! it was " & txtWord)
 lblBlueScr = Val(lblBlueScr) + 5
 MsgBox "put a new word and type enter"
 txtWord.Enabled = True
 txtWord.Text = " "
End If
If strWinner$ = "0" Then
 qrkAIM.ChatSend (Who$ & " join a team first")
 qrkAIM.ChatSend ("the word was " & txtWord & " but you can't get it cause dumbass " & Who$ & " said it without being on a team")
 MsgBox "put a new word and type enter"
 txtWord.Enabled = True
 txtWord.Text = " "
End If
End If
End Sub

Private Sub txtWord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lblOutput = Scramble(txtWord)
qrkAIM.ChatSend ("team scrambler, new word " & Chr(34) & lblOutput & Chr(34))
txtWord.Enabled = False
KeyAscii = 0
End If
End Sub
