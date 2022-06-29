VERSION 5.00
Object = "{655D25A2-69EB-11D2-A9C1-D94536B35B75}#2.0#0"; "SCOREKEEPER3.OCX"
Begin VB.Form Form1 
   Caption         =   "Scrambler Score Keeper v1.2 Example"
   ClientHeight    =   4350
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "name1"
      Top             =   1320
      Width           =   975
   End
   Begin ScoreKeeper.Score Score1 
      Height          =   1935
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3413
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fake Chat Room:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Menu mnuSubs 
      Caption         =   "Subs"
      Begin VB.Menu mnuAddnameandscore 
         Caption         =   "AddNameAndScore"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuSendscore 
         Caption         =   "SendScores"
      End
      Begin VB.Menu mnuSortscore 
         Caption         =   "SortScore"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClse 
         Caption         =   "Close Example"
      End
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "Propterties"
      Begin VB.Menu mnuEnabled 
         Caption         =   "Enabled"
      End
      Begin VB.Menu mnuHowmanypeople 
         Caption         =   "HowManyPeople"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text3.Text = "Online Host:" & (Chr(13) + Chr(10)) & "Online Host:     You are now in ***Fake Chat Room***" & (Chr(13) + Chr(10)) & "Online Host:"
Score1.AddNameAndScore "name1", 2
Score1.AddNameAndScore "name2", 4
Score1.AddNameAndScore "name3", 1
Score1.AddNameAndScore "name4", 8
Score1.AddNameAndScore "name5", 5
Score1.AddNameAndScore "name6", 12
End Sub

Private Sub mnuAddnameandscore_Click()
'some code to check if text2 is an integer
'would be good
Score1.AddNameAndScore Text1, Text2
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub mnuEnabled_Click()
If Score1.Enabled = True Then
Score1.Enabled = False
ElseIf Score1.Enabled = False Then
Score1.Enabled = True
End If
End Sub

Private Sub mnuRefresh_Click()
Score1.Refresh
End Sub

Private Sub mnuSendscore_Click()
Score1.SendScore
End Sub

Private Sub mnuSortscore_Click()
Score1.SortScore
End Sub

Private Sub Score1_SendScore(Score As String)
If Score = "" Then
Else
Text3.Text = Text3 & (Chr(13) + Chr(10)) & "Person:  " & Score
End If
End Sub
