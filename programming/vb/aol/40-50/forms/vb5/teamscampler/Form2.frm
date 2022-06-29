VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bismuth Team Bot"
   ClientHeight    =   2850
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB5Chat2.Chat Chat1 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.ListBox List3 
      Height          =   645
      ItemData        =   "Form2.frx":0000
      Left            =   3480
      List            =   "Form2.frx":0002
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtSn 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   5520
      Top             =   360
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5520
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Team2"
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1035
      ItemData        =   "Form2.frx":0004
      Left            =   1680
      List            =   "Form2.frx":0006
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Form2.frx":0008
      Left            =   0
      List            =   "Form2.frx":000A
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Team1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "People Left:"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Time Left:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "n/a"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "n/a"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "0 people"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "0 people"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "0 people total"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnuPhile 
      Caption         =   "Phile"
      Begin VB.Menu mnuExample 
         Caption         =   "Example"
      End
      Begin VB.Menu mnuLine99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "Back"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuGountil 
         Caption         =   "Set Go Until"
      End
      Begin VB.Menu mnuTimelimit 
         Caption         =   "Set Time Limit"
      End
      Begin VB.Menu mnuUseplayerslastgame 
         Caption         =   "Players From Last Game"
      End
      Begin VB.Menu mnuLine78 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIfnotteam 
         Caption         =   "Tell Player if not a Team"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadplayer 
         Caption         =   "Load Player List"
      End
      Begin VB.Menu mnuSaveplayers 
         Caption         =   "Save Current Players"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsinghelp 
         Caption         =   "Using Help"
      End
      Begin VB.Menu mnuUsingteambot 
         Caption         =   "Using Team Bot"
      End
      Begin VB.Menu mnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnumailme 
         Caption         =   "Mail Me"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public thename As String
Public thechat As String
Public peoplethatcanplay As String
Private Sub Command1_Click()
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM Bø† *Active*")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(Type the team you want on")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(Teams: " & Text1 & " & " & Text2)
If Label5.Caption = "n/a" Then
Timer1.Enabled = True
ElseIf Label5.Caption <> "n/a" Then
Timer2.Enabled = True
peoplethatcanplay = Label5.Caption
End If
mnuPhile.Enabled = False
mnuOptions.Enabled = False
mnuHelp.Enabled = False
Command2.Enabled = True
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
End Sub
Private Sub Command2_Click()
Timer1.Enabled = False
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM Bø† *Inactive*")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(" & Text1 & " has " & GetLabelNum(Label2) & " players")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(" & Text2 & " has " & GetLabelNum(Label3) & " players")
mnuPhile.Enabled = True
mnuOptions.Enabled = True
mnuHelp.Enabled = True
Command2.Enabled = False
Command1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
End Sub
Private Sub Form_Load()
StayOnTop Me
Command2.Enabled = False
Command1.Enabled = True
mnuIfnotteam.Checked = False
List1.AddItem ("Team1 Players:")
List2.AddItem ("Team2 Players:")
End Sub
Private Sub mnuBack_Click()
Dim a As Integer
List1.RemoveItem (0)
List2.RemoveItem (0)
Form1.Label6.Caption = Text1
Form1.Label7.Caption = Text2
For a = 0 To List1.ListCount - 1
Form1.List1.AddItem (List1.List(a))
Form1.List2.AddItem ("0")
Next a
For a = 0 To List2.ListCount - 1
Form1.List3.AddItem (List2.List(a))
Form1.List4.AddItem ("0")
Next a
Unload Me
Form1.Visible = True
End Sub
Private Sub mnuExample_Click()
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM Bø† Example:")
Chat1.ChatSend ("  ")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">" & Text1)
End Sub
Private Sub mnuGountil_Click()
MsgBox "Reminder:  Add check if # code!!!", vbExclamation, "Note to self!!!!"
Label5.Caption = InputBox("Please enter the amount of players you would like added before the bot automatically turns off", "Bismuth Team Bot")
End Sub
Private Sub mnuIfnotteam_Click()
If mnuIfnotteam.Checked = True Then
mnuIfnotteam.Checked = False
ElseIf mnuIfnotteam.Checked = False Then
mnuIfnotteam.Checked = True
End If
End Sub
Private Sub mnuTimelimit_Click()
MsgBox "Reminder:  Add check if # code!!!", vbExclamation, "Note to self!!!!"
Label4.Caption = InputBox("Please enter the time (in seconds) that you would like to continue to add users before the bot automatically turns off", "Bismuth Team Bot")
End Sub
Private Sub Timer1_Timer()
thechat = LastChatLineWithSN
If ChatOnly(thechat) = Text1 Then
thename = NameOnly(thechat)
addperson1 ("Timer1")
ElseIf ChatOnly(thechat) = Text2 Then
thename = NameOnly(thechat)
addperson2 ("Timer1")
End If
End Sub
Public Sub addperson1(thetimer As String)
Dim a As Integer
Dim b As Integer
Dim G As Boolean
For a = 1 To List1.ListCount
If thename = List1.List(a) Then
G = True
End If
Next a
For a = 1 To List2.ListCount
If thename = List2.List(a) Then
G = True
End If
Next a
If G = False Then
List1.AddItem (thename)
b = GetLabelNum(Label2) + 1
Label2.Caption = b & " people"
If thetimer = "Timer1" Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & thename & " added to " & Text2 & ", player #" & b)
ElseIf thetimer = "Timer2" Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & thename & " added to " & Text2 & ", player #" & b)
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(" & Val(Label5.Caption) - 1 & " more players can join")
Label5.Caption = Val(Label5.Caption) - 1
End If
ElseIf G = False Then
'check if supposed to tell
End If
b = 0
a = 0
End Sub
Public Sub addperson2(thetimer As String)
Dim a As Integer
Dim b As Integer
Dim G As Boolean
For a = 1 To List2.ListCount
If thename = List2.List(a) Then
G = True
End If
Next a
For a = 1 To List1.ListCount
If thename = List1.List(a) Then
G = True
End If
Next a
If G = False Then
List2.AddItem (thename)
b = GetLabelNum(Label3) + 1
Label3.Caption = b & " people"
If thetimer = "Timer1" Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & thename & " added to " & Text2 & ", player #" & b)
ElseIf thetimer = "Timer2" Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & thename & " added to " & Text2 & ", player #" & b)
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(" & Val(Label5.Caption) - 1 & " more players can join")
Label5.Caption = Val(Label5.Caption) - 1
End If
ElseIf G = True Then
'asdf
End If

a = 0
b = 0
End Sub
Private Sub Timer2_Timer()
If Val(GetLabelNum(Label2)) + Val(GetLabelNum(Label3)) = peoplethatcanplay Then
Timer2.Enabled = False
Timeout (0.1)
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM Bø† *Inactive*")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(" & Text1 & " has " & GetLabelNum(Label2) & " players")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(" & Text2 & " has " & GetLabelNum(Label3) & " players")
mnuPhile.Enabled = True
mnuOptions.Enabled = True
mnuHelp.Enabled = True
Command2.Enabled = False
Command1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Else
thechat = LastChatLineWithSN
If ChatOnly(thechat) = Text1 Then
thename = NameOnly(thechat)
addperson1 ("Timer2")
ElseIf ChatOnly(thechat) = Text2 Then
thename = NameOnly(thechat)
addperson2 ("Timer2")
End If
End If
End Sub
