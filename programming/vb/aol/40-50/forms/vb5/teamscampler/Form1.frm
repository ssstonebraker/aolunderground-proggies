VERSION 5.00
Object = "{5B033ECF-098E-11D1-A4B2-444553540000}#1.0#0"; "SUBCLASS.OCX"
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bismuth Team Scrambler"
   ClientHeight    =   3690
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3540
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB5Chat2.Chat Chat1 
      Left            =   3120
      Top             =   3720
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Pause"
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Text            =   "Word to Scramble"
      Top             =   0
      Width           =   2055
   End
   Begin SubclassCtl.Subclass Subclass1 
      Left            =   2880
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Top 3"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Manual"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   1440
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Auto"
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "All"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Personal"
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Team"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List4 
      Height          =   840
      ItemData        =   "Form1.frx":000C
      Left            =   3120
      List            =   "Form1.frx":000E
      TabIndex        =   5
      Top             =   2640
      Width           =   375
   End
   Begin VB.ListBox List3 
      Height          =   840
      ItemData        =   "Form1.frx":0010
      Left            =   1800
      List            =   "Form1.frx":0012
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   840
      ItemData        =   "Form1.frx":0014
      Left            =   1320
      List            =   "Form1.frx":0016
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "Form1.frx":0018
      Left            =   0
      List            =   "Form1.frx":001A
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2160
      Max             =   250
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   1
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Subject"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1440
   End
   Begin VB.Label Label8 
      Caption         =   "Scroll Scores :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Team 2"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Team 1"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "0 players"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "0 players"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1 point"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Menu mnuPhile 
      Caption         =   "Phile"
      Begin VB.Menu mnuAdverstise 
         Caption         =   "Advertise"
      End
      Begin VB.Menu mnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubjectbot 
         Caption         =   "Subject Bot"
      End
      Begin VB.Menu mnuTeambot2 
         Caption         =   "Team Bot"
      End
      Begin VB.Menu mnuVotebot 
         Caption         =   "Vote Bot"
      End
      Begin VB.Menu mnuLine19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustombot 
         Caption         =   "Custom Bot"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRoombust 
         Caption         =   "Room Bust"
      End
      Begin VB.Menu mnuIMz 
         Caption         =   "IMz"
         Begin VB.Menu mnuOn 
            Caption         =   "On"
         End
         Begin VB.Menu mnuOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu mnuLine22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuLine6732 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuTems 
      Caption         =   "Teams"
      Begin VB.Menu mnuReset3 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTeambot 
         Caption         =   "Team Bot"
      End
      Begin VB.Menu mnuLint7 
         Caption         =   "-"
      End
      Begin VB.Menu mnBalanceteams 
         Caption         =   "Balance Teams"
      End
      Begin VB.Menu mnuLine7456 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTellplayer 
         Caption         =   "Tell Player if not on team"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLine800 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddPlayer 
         Caption         =   "Add Player"
      End
      Begin VB.Menu mnuRemoveplayer 
         Caption         =   "Remove Player"
      End
      Begin VB.Menu mnuChangeteamname 
         Caption         =   "Change Team Name"
      End
   End
   Begin VB.Menu mnuScoring 
      Caption         =   "Scoring"
      Begin VB.Menu mnuBywordlenght 
         Caption         =   "By Word Length"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuManually 
         Caption         =   "User Defined"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFixscore 
         Caption         =   "Fix Score"
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshscore 
         Caption         =   "Refresh Score"
      End
      Begin VB.Menu mnuAnotherline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTeam 
         Caption         =   "Team Scores"
      End
      Begin VB.Menu mnuPersonal 
         Caption         =   "Personal Scores"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "All Scores"
      End
      Begin VB.Menu mnutopforteam 
         Caption         =   "Top For Each Team"
      End
      Begin VB.Menu mnuTop3 
         Caption         =   "Top 3 Over All"
      End
      Begin VB.Menu mnuLine18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighestever 
         Caption         =   "Highest Ever"
      End
      Begin VB.Menu mnuToptenteams 
         Caption         =   "Top Ten Teams"
      End
      Begin VB.Menu mnuToptenpeepz 
         Caption         =   "Top Ten Peepz"
      End
      Begin VB.Menu mnuLine78 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewtopten 
         Caption         =   "View High Scores"
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuLine14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadlist 
         Caption         =   "Load List"
      End
      Begin VB.Menu mnuUnloadlist 
         Caption         =   "Unload List"
      End
      Begin VB.Menu mnuLine20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListconverter 
         Caption         =   "List Converter"
      End
      Begin VB.Menu mnuLInte15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListeditor 
         Caption         =   "List Editor"
      End
      Begin VB.Menu mnuReset2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSecondclue 
         Caption         =   "Second Clue"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSpellinghint 
         Caption         =   "Spelling Hint"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuLine9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChattoplay 
         Caption         =   "How to Play Scrambler"
      End
      Begin VB.Menu mnuLine9856 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnnouncehost 
         Caption         =   "Announce Host"
      End
      Begin VB.Menu mnuEndgame 
         Caption         =   "End Game"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuABout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuGreetz 
         Caption         =   "Greetz"
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnudisclaimer 
         Caption         =   "Disclaimer"
      End
      Begin VB.Menu mnuUsinghelp 
         Caption         =   "Using Help"
      End
      Begin VB.Menu mnuUsingatm 
         Caption         =   "Using Atm"
      End
      Begin VB.Menu mnuLine12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuABoutfft 
         Caption         =   "About FFT"
      End
      Begin VB.Menu mnuApp 
         Caption         =   "App"
      End
      Begin VB.Menu mnuLine13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNextversion 
         Caption         =   "Next Vers."
      End
      Begin VB.Menu mnuMailme 
         Caption         =   "Mail Me"
      End
      Begin VB.Menu mnuLine17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSupertones 
         Caption         =   "Supertones!"
      End
      Begin VB.Menu mnuLine29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStoopidthoughts 
         Caption         =   "Stoopid Thoughts"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Windows declarations
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'Windows constants
Private Const WM_SYSCOMMAND = &H112
Private Const MF_STRING = &H0
Private Const MF_SEPARATOR = &H800
'ID for new About command (must be < &HF000)
Private Const IDM_ABOUT = 10
Private Const IDM_MAIL = 11
Public thename As String
Public thechat As String



Private Sub Command1_Click()
If Option1.Value = True Then
'automatic mode
ElseIf Option2.Value = True Then
'manual mode
FOrmOff
Command7.Enabled = True
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM §ç®åMb£è®")
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(Uscramble: " & ScrambleText(Text1))
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">º°`(Subject: " & Text2 & " - Points: " & GetScore)
Timer1.Enabled = True
Else
MsgBox "You must select either Automatic or Manual!", vbExclamation, "bISMUTH"
End If
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
FormOn
Command7.Enabled = False
SendChat ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM §ç®åMb£è® Disabled")
End Sub
Private Sub Command7_Click()
If Command7.Caption = "Pause" Then
Command7.Caption = "Pause (On)"
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM §ç®åMb£è® Paused")
Timer1.Enabled = False
ElseIf Command7.Caption = "Pause (On)" Then
Command7.Caption = "Pause"
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(Bì§Mµ†h †êåM §ç®åMb£è® UnPaused")
Timer1.Enabled = True
End If
End Sub

Private Sub Form_Load()
StayOnTop Me
Option2.Value = True
mnuBywordlenght.Checked = False
mnuManually.Checked = True
    Dim i As Long, hMenu As Long
    'Add "About..." command to system menu
    hMenu = GetSystemMenu(Me.hwnd, False)
    i = AppendMenu(hMenu, MF_SEPARATOR, 0, 0&)
    i = AppendMenu(hMenu, MF_STRING, IDM_ABOUT, "&About Bismuth")
    i = AppendMenu(hMenu, MF_STRING, IDM_MAIL, "E-Mail Me")
    'Setup Subclass
    Subclass1.hwnd = Me.hwnd
    Subclass1.Messages(WM_SYSCOMMAND) = True
Command7.Enabled = False
End Sub



Private Sub HScroll1_Change()
Label1.Caption = HScroll1.Value & " points"
End Sub






Private Sub mnuAddPlayer_Click()
Form3.Visible = True
Form1.Visible = False
End Sub

Private Sub mnuAdverstise_Click()
MsgBox "|" & ChatOnly(LastChatLineWithSN)
End Sub
Private Sub mnuBywordlenght_Click()
mnuBywordlenght.Checked = True
mnuManually.Checked = False
HScroll1.Enabled = False
End Sub
Private Sub mnuClose_Click()
End
End Sub
Private Sub mnuListconverter_Click()
MsgBox "Well sad to say, this will be in the next version. So why did i add it on the menu? i don't know.  Anywayz what it will do is take for example a Raydr Log file an convert it to be compatible with Bismuth Team Scrambler.  But i need to know what scrambler logs i should add to this.  If your a programmer, [spying on the competition?] and you have a scrambler that uses logs, and you would like me to add a converter from my log to yours let me know, or if you want your logs to be converted to my format let me know.  Skanky DooDle!", vbInformation, "About List Converter"
End Sub
Private Sub mnuManually_Click()
mnuManually.Checked = True
mnuBywordlenght.Checked = False
HScroll1.Enabled = True
End Sub
Private Sub mnuTellplayer_Click()
If mnuTellplayer.Checked = True Then
 mnuTellplayer.Checked = False
ElseIf mnuTellplayer.Checked = False Then
 mnuTellplayer.Checked = True
End If
End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    'Look for WM_SYSCOMMAND message with About command
    If Msg = WM_SYSCOMMAND Then
        If wParam = IDM_ABOUT Then
            'about box code here
            Exit Sub
        ElseIf wParam = IDM_MAIL Then
            MsgBox "you clicked E-mail Me!", vbInformation, "Kewl"
            Exit Sub
        End If
    End If
    'Pass along to default handler if message not processed
    Result = Subclass1.CallWndProc(Msg, wParam, lParam)
End Sub
Private Sub mnuRoombust_Click()
MsgBox "Someone do this, i don't want to", vbInformation, "Not my area"
End Sub

Private Sub mnuStoopidthoughts_Click()
MsgBox "I'm not suffering from insanity, I'm enjoying ever minute of it!", , "Stoopid Thought"
MsgBox "Do you believe in love at first sight, or should i walk by again?", , "Stoopid Thought"
MsgBox "Your Unique! just like everyone else", , "Stoopid Thought"
MsgBox "All of those were found on the internet somewhere, so credit to the peepz who thought them up!  If you want one added you know where to e-Mail me.", , "Stoopid Thoughts"
End Sub

Private Sub mnuTeambot_Click()
Form1.Visible = False
Form2.Visible = True
End Sub

Private Sub mnuTeambot2_Click()
Form2.Visible = True
Form2.Visible = False
End Sub

Public Sub FormOn()
Command1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
If mnuBywordlenght.Checked = True Then
HScroll1.Enabled = False
ElseIf mnuManually.Checked = True Then
HScroll1.Enabled = True
End If
Command7.Enabled = True
Command6.Enabled = True
Command5.Enabled = True
Command4.Enabled = True
Command3.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
mnuPhile.Enabled = True
mnuTems.Enabled = True
mnuGame.Enabled = True
mnuScoring.Enabled = True
mnuHelp.Enabled = True
End Sub

Public Sub FOrmOff()
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
HScroll1.Enabled = False
Command7.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
Command4.Enabled = False
Command3.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
mnuPhile.Enabled = False
mnuTems.Enabled = False
mnuGame.Enabled = False
mnuScoring.Enabled = False
mnuHelp.Enabled = False
End Sub

Public Function GetScore()
If mnuManually.Checked = True Then
GetScore = HScroll1.Value
ElseIf mnuBywordlenght.Checked = True Then
GetScore = Len(Text1)
End If
End Function

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
thechat = ChatOnly(LastChatLineWithSN)
If thechat = Text1 Then
thename = NameOnly(LastChatLineWithSN)
AddPoints
End If
End Sub

Public Sub AddPoints()
Dim a As Integer
Dim G As Boolean

For a = 0 To List1.ListCount
If thename = List1.List(a) Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(WTG! " & thename & " got it!" & " [" & Label6.Caption & "," & Val(List2.List(a)) + 1 & "," & Label4.Caption + Val(List2.List(a)) + 1 & "]")
List1.AddItem (thename)
List2.AddItem (Val(List2.List(a)) + 1)
List1.RemoveItem (a)
List2.RemoveItem (a)
Label4.Caption = Val(Label4.Caption) + Val(List2.List(a))
G = True
FormOn
End If
Next a
For a = 0 To List3.ListCount
If thename = List3.List(a) Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(WTG! " & thename & " got it!" & " [" & Label7.Caption & "," & Val(List4.List(a)) + 1 & "," & Label5.Caption + Val(List4.List(a)) + 1 & "]")
List3.AddItem (thename)
List4.AddItem (Val(List4.List(a)) + 1)
List3.RemoveItem (a)
List4.RemoveItem (a)
Label5.Caption = Val(Label5.Caption) + Val(List4.List(a))
G = True
FormOn
End If
Next a
If G = False Then
If mnuTellplayer.Checked Then
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & thename & ", your not on a team!")
FormOn
Else
MsgBox thename & " isn't on a team!"
FormOn
End If
End If
End Sub

