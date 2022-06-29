VERSION 5.00
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "ANIGIF.OCX"
Object = "{84A2E5B4-473D-11D1-BABB-0C0909C10000}#5.0#0"; "K_STRAY32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HyPO For AOL 4.0"
   ClientHeight    =   1950
   ClientLeft      =   3105
   ClientTop       =   3330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5385
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Text            =   "<A HREF=""mailto:Toastny@hotmail.com"">Mail ToaST</A>"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   0
      TabIndex        =   20
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Macro Kill 1"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Advertise"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin k_STray32.k_STray k_STray1 
      Left            =   2400
      Top             =   720
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   720
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin AniGIFCtrl.AniGIF AniGIF10 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":0000
      ExtendWidth     =   2355
      ExtendHeight    =   661
   End
   Begin AniGIFCtrl.AniGIF AniGIF9 
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":1F08
      ExtendWidth     =   3201
      ExtendHeight    =   873
   End
   Begin AniGIFCtrl.AniGIF AniGIF8 
      Height          =   615
      Left            =   3480
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":2A28
      ExtendWidth     =   3413
      ExtendHeight    =   1085
   End
   Begin AniGIFCtrl.AniGIF AniGIF7 
      Height          =   615
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":7FA8
      ExtendWidth     =   3413
      ExtendHeight    =   1085
   End
   Begin AniGIFCtrl.AniGIF AniGIF6 
      Height          =   615
      Left            =   1440
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":ADF8
      ExtendWidth     =   3625
      ExtendHeight    =   1085
   End
   Begin AniGIFCtrl.AniGIF AniGIF5 
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":CB88
      ExtendWidth     =   3625
      ExtendHeight    =   873
   End
   Begin AniGIFCtrl.AniGIF AniGIF4 
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":DC44
      ExtendWidth     =   3625
      ExtendHeight    =   873
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      Picture         =   "Form1.frx":11454
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":136B1
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1440
      Picture         =   "Form1.frx":155E5
      ScaleHeight     =   585
      ScaleWidth      =   2025
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      Picture         =   "Form1.frx":17CCC
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1440
      Picture         =   "Form1.frx":19CCE
      ScaleHeight     =   465
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      Picture         =   "Form1.frx":1C29D
      ScaleHeight     =   465
      ScaleWidth      =   1905
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1440
      Picture         =   "Form1.frx":1E27B
      ScaleHeight     =   465
      ScaleWidth      =   2025
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin AniGIFCtrl.AniGIF AniGIF2 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   5415
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":20891
      ExtendWidth     =   9551
      ExtendHeight    =   450
   End
   Begin AniGIFCtrl.AniGIF AniGIF1 
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":2324D
      ExtendWidth     =   9551
      ExtendHeight    =   450
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form1.frx":25C09
      ScaleHeight     =   1635
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.Menu bots 
      Caption         =   "Bots"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuBotsScreenNameDecoder 
         Caption         =   "Screen Name Decoder"
      End
      Begin VB.Menu mnuBotsEchoBot 
         Caption         =   "Echo Bot"
      End
      Begin VB.Menu mnuBotsCoolChat 
         Caption         =   "CoolChat"
      End
      Begin VB.Menu mnuBotsWarezRequester 
         Caption         =   "Warez Requester"
      End
      Begin VB.Menu mnuBotsFakeTermer 
         Caption         =   "Fake Termer"
      End
      Begin VB.Menu mnuBotsPhader 
         Caption         =   "Phader"
      End
      Begin VB.Menu mnuBotsVoter 
         Caption         =   "Voter"
      End
      Begin VB.Menu mnuBotsNumberGuessBot 
         Caption         =   "Number Guess Bot"
      End
      Begin VB.Menu mnuBotsCloseSpiningAOL 
         Caption         =   "Close Spining AOL"
      End
      Begin VB.Menu mnuBotsGuideBots 
         Caption         =   "Guide Bots"
         Index           =   1
      End
      Begin VB.Menu mnuBotsWarezGroupCreator 
         Caption         =   "Warez Group Creator"
         Index           =   2
      End
      Begin VB.Menu mnuBotsAnnoyBots 
         Caption         =   "Annoy Bots"
         Index           =   3
      End
      Begin VB.Menu mnuBotsFakeRoom 
         Caption         =   "Fake Room"
         Index           =   4
      End
      Begin VB.Menu mnuBotsMassImer 
         Caption         =   "Mass Imer"
         Index           =   5
      End
      Begin VB.Menu mnuBotsIdleBot 
         Caption         =   "Idle Bot"
         Index           =   6
      End
      Begin VB.Menu mnuBotsRoomFreezer 
         Caption         =   "Room Freezer"
         Index           =   7
      End
      Begin VB.Menu mnuBotsUpChat 
         Caption         =   "Up Chat"
      End
      Begin VB.Menu mnuAttentionBot 
         Caption         =   "Attention Bot"
         Index           =   8
      End
   End
   Begin VB.Menu punter 
      Caption         =   "Punter"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnupunterErrorPunter 
         Caption         =   "Error Punter"
      End
      Begin VB.Menu mnupunterMailPunter 
         Caption         =   "Mail Punter"
      End
      Begin VB.Menu mnupunterPunter 
         Caption         =   "Punter"
      End
   End
   Begin VB.Menu mnuSystemTray 
      Caption         =   "menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Program"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Program"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Tea$ = "-æ-[ HyPO ™ ]-æ-Advertised by " + UserSN
 fnt$ = "10"
a = Len(Tea$)
For w = 1 To a Step 4
    r$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    s$ = Mid$(Tea$, w + 2, 1)
    t$ = Mid$(Tea$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & r$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavYChaTRBb = p$
SendChat WavYChaTRBb
TimeOut 0.5
'
'
Tea$ = "···÷••(¯`·._=-HyPO ver.¹·º 4 AOL 4 -=_.·´¯)••÷"

a = Len(Tea$)
For w = 1 To a Step 4
    PimPimp5$ = Mid$(Tea$, w, 1)
    Pimp2$ = Mid$(Tea$, w + 1, 1)
    Pimp3$ = Mid$(Tea$, w + 2, 1)
    Pimp4$ = Mid$(Tea$, w + 3, 1)
    Pimp5$ = Pimp5$ & "<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><b>" & PimPimp5$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & Pimp2$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & Pimp3$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & Pimp4$
Next w
WavYChaTRBbPiMp = Pimp5$
SendChat WavYChaTRBbPiMp
TimeOut 0.5
'
'
'
ToaST$ = "···÷••(¯`·._By ToaST _.·´¯)••÷ "

a = Len(ToaST$)
For w = 1 To a Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sup><b>" & ToaSTR$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
'···÷••(¯`·._   _.·´¯)••÷
SendChat WavYChaTRBb
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Image1_Click()
MsgBox "OK"
End Sub

Private Sub AniGIF10_Click()
'<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><b>" & PimPimp5$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & Pimp2$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & Pimp3$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & Pimp4$
IMsOn
Tea$ = "-æ-[ HyPO ™ ]-æ-" + UserSN + "'s IM's On "
 fnt$ = "10"
a = Len(Tea$)
For w = 1 To a Step 4
    r$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    s$ = Mid$(Tea$, w + 2, 1)
    t$ = Mid$(Tea$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & r$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & t$
Next w
WavYChaTRBb = p$
SendChat WavYChaTRBb
End Sub

Private Sub AniGIF4_Click()
Form16.Show
End Sub

Private Sub AniGIF5_Click()
Form13.Show
End Sub

Private Sub AniGIF6_Click()
PopupMenu punter
End Sub

Private Sub AniGIF7_Click()
IMsOff
Tea$ = "-æ-[ HyPO ™ ]-æ-" + UserSN + "'s IM's Off "
 fnt$ = "10"
a = Len(Tea$)
For w = 1 To a Step 4
    r$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    s$ = Mid$(Tea$, w + 2, 1)
    t$ = Mid$(Tea$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & r$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavYChaTRBb = p$
SendChat WavYChaTRBb
End Sub

Private Sub AniGIF8_Click()

 PopupMenu bots
End Sub

Private Sub AniGIF9_Click()
End
End Sub

Private Sub Command3_Click()
r = String(80, Chr(64))
D = 80 - Len(Text1)
C$ = Left(r, D)
'SendChat ("<font=#000ff>" & "" & C$ & "</font>")
SendChat ("<font=#FF0000>" & "" & C$ & "")
lonh = String(90, Chr(64))
D = 90 - Len(Text1)
C$ = Left(r, D)
'SendChat ("" & "" & C$ & "")
'SendChat ("<font=#000ff>" & "" & C$ & "</font>")
End Sub

Private Sub Form_Activate()
Do
Text1 = Time
DoEvents
Loop
End Sub

Private Sub Form_Deactivate()
Form1.WindowState = 1
End Sub

Private Sub Form_Load()
 k_STray1.AddTrayIcon Me
Tea$ = "-æ-[ HyPO ™ ]-æ-loaded by " + UserSN
 fnt$ = "10"
a = Len(Tea$)
For w = 1 To a Step 4
    r$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    s$ = Mid$(Tea$, w + 2, 1)
    t$ = Mid$(Tea$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & r$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavYChaTRBb = p$
SendChat WavYChaTRBb
TimeOut 0.5
'
'
Tea$ = "···÷••(¯`·._=-HyPO ver.¹·º 4 AOL 4 -=_.·´¯)••÷"

a = Len(Tea$)
For w = 1 To a Step 4
    PimPimp5$ = Mid$(Tea$, w, 1)
    Pimp2$ = Mid$(Tea$, w + 1, 1)
    Pimp3$ = Mid$(Tea$, w + 2, 1)
    Pimp4$ = Mid$(Tea$, w + 3, 1)
    Pimp5$ = Pimp5$ & "<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><b>" & PimPimp5$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & Pimp2$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & Pimp3$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & Pimp4$
Next w
WavYChaTRBbPiMp = Pimp5$
SendChat WavYChaTRBbPiMp
TimeOut 0.5
'

'
'
ToaST$ = "···÷••(¯`·._By ToaST _.·´¯)••÷ "

a = Len(ToaST$)
For w = 1 To a Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sup><b>" & ToaSTR$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
'···÷••(¯`·._   _.·´¯)••÷
SendChat WavYChaTRBb
TimeOut 1
toster$ = """MAILTO:ToaSTNy@hotmail.com"""
 
SendChat Text2



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  k_STray1.Form_MouseMove Me, X
End Sub

Private Sub Form_Resize()
  k_STray1.Form_Resize Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 k_STray1.DelTrayIcon Me
End Sub

Private Sub mnuAttentionBot_Click(Index As Integer)
Form6.Show
End Sub

Private Sub mnuBotsAnnoyBots_Click(Index As Integer)
Form12.Show
End Sub

Private Sub mnuBotsCloseSpiningAOL_Click()
KillGlyph
End Sub

Private Sub mnuBotsCoolChat_Click()
Form23.Show
End Sub

Private Sub mnuBotsEchoBot_Click()
Form25.Show
End Sub

Private Sub mnuBotsFakeRoom_Click(Index As Integer)
Form14.Show
End Sub

Private Sub mnuBotsFakeTermer_Click()
Form21.Show
End Sub

Private Sub mnuBotsGuideBots_Click(Index As Integer)
Form10.Show
End Sub

Private Sub mnuBotsIdleBot_Click(Index As Integer)
Form7.Show
End Sub

Private Sub mnuBotsMassImer_Click(Index As Integer)
Form3.Show
End Sub

Private Sub mnuBotsNumberGuessBot_Click()
Form18.Show
End Sub

Private Sub mnuBotsPhader_Click()
Form22.Show
End Sub

Private Sub mnuBotsRoomFreezer_Click(Index As Integer)
Form8.Show
End Sub

Private Sub mnuBotsScreenNameDecoder_Click()
Form26.Show
End Sub

Private Sub mnuBotsUpChat_Click()
Form2.Show
End Sub

Private Sub mnuBotsVoter_Click()
Form19.Show
End Sub

Private Sub mnuBotsWarezGroupCreator_Click(Index As Integer)
Form11.Show
End Sub

Private Sub mnuBotsWarezRequester_Click()
Form20.Show
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnupunterErrorPunter_Click()
Form9.Show
End Sub

Private Sub mnupunterMailPunter_Click()
Form4.Show
End Sub

Private Sub mnupunterPunter_Click()
Form5.Show
End Sub

Private Sub mnuSystemTrayExit_Click()
Unload Me
End Sub

Private Sub mnuSystemTrayRestore_Click()
 k_STray1.Restore Me
End Sub

Private Sub mnuRestore_Click()
     k_STray1.Restore Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

AniGIF4.Visible = False
AniGIF5.Visible = False
AniGIF6.Visible = False
AniGIF7.Visible = False
AniGIF8.Visible = False
AniGIF9.Visible = False
AniGIF10.Visible = False

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
AniGIF4.Visible = True
 
AniGIF5.Visible = False
AniGIF6.Visible = False
AniGIF7.Visible = False
AniGIF8.Visible = False
AniGIF9.Visible = False
AniGIF10.Visible = False
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AniGIF9.Visible = True
AniGIF4.Visible = False
AniGIF5.Visible = False
AniGIF6.Visible = False
AniGIF7.Visible = False
AniGIF8.Visible = False
 
AniGIF10.Visible = False
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AniGIF5.Visible = True
AniGIF4.Visible = False
 
AniGIF6.Visible = False
AniGIF7.Visible = False
AniGIF8.Visible = False
AniGIF9.Visible = False
AniGIF10.Visible = False
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AniGIF8.Visible = True
AniGIF4.Visible = False
AniGIF5.Visible = False
AniGIF6.Visible = False
AniGIF7.Visible = False
 
AniGIF9.Visible = False
AniGIF10.Visible = False
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AniGIF6.Visible = True
AniGIF4.Visible = False
AniGIF5.Visible = False
 
AniGIF7.Visible = False
AniGIF8.Visible = False
AniGIF9.Visible = False
AniGIF10.Visible = False
End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AniGIF10.Visible = True
AniGIF4.Visible = False
AniGIF5.Visible = False
AniGIF6.Visible = False
AniGIF7.Visible = False
AniGIF8.Visible = False
AniGIF9.Visible = False
 
End Sub

Private Sub Picture8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AniGIF7.Visible = True
AniGIF4.Visible = False
AniGIF5.Visible = False
AniGIF6.Visible = False
 
AniGIF8.Visible = False
AniGIF9.Visible = False
AniGIF10.Visible = False
End Sub

Private Sub Text1_Change()
AniGIF7.Visible = True
AniGIF4.Visible = False
AniGIF5.Visible = False
AniGIF6.Visible = False
 AniGIF7.Visible = False
AniGIF8.Visible = False
AniGIF9.Visible = False
AniGIF10.Visible = False
End Sub

