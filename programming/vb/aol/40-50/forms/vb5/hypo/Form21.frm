VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H80000007&
   Caption         =   "Fake Termer"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form21"
   ScaleHeight     =   1275
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Text            =   "Steve Case"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scare them"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Tea$ = "-æ-[ HyPO ™ ]-æ-Termer "
 fnt$ = "10"
A = Len(Tea$)
For w = 1 To A Step 4
    R$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    S$ = Mid$(Tea$, w + 2, 1)
    T$ = Mid$(Tea$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & R$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRBb = P$
SendChat WavYChaTRBb
TimeOut 0.5
'
'
Tea$ = "···÷••(¯`·._TERMING """ + Text1 + """ _.·´¯)••÷"

A = Len(Tea$)
For w = 1 To A Step 4
    PimPimp5$ = Mid$(Tea$, w, 1)
    Pimp2$ = Mid$(Tea$, w + 1, 1)
    Pimp3$ = Mid$(Tea$, w + 2, 1)
    Pimp4$ = Mid$(Tea$, w + 3, 1)
    Pimp5$ = Pimp5$ & "<FONT COLOR=" & Chr$(34) & "#FFCL25" & Chr$(34) & "><sup><b>" & PimPimp5$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & Pimp2$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & Pimp3$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#FFD700" & Chr$(34) & ">" & Pimp4$
Next w
WavYChaTRBbPiMp = Pimp5$
SendChat WavYChaTRBbPiMp
TimeOut 1
'

'
'
ToaST$ = "···÷••(¯`·._BYE BYE _.·´¯)••÷ "

A = Len(ToaST$)
For w = 1 To A Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sup><b>" & ToaSTR$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
SendChat WavYChaTRBb
TimeOut 1
SendChat "<font color=#0000FF><b>10-- Opening Tos Window"
TimeOut 1
SendChat "<font color=#0033FF><b>9-- Finding Violation button"
TimeOut 1
SendChat "<font color=#0066FF><b>8-- Button Found.. Standbye"
TimeOut 1
SendChat "<font color=#0099FF><b>7-- Writing Report..."
TimeOut 1
SendChat "<font color=#00CCFF><b>6-- Sending .. Please Wait"
TimeOut 1
SendChat "<font color=#FF0066><b>5-- Checking Mail Status"
TimeOut 1
SendChat "<font color=#FF0033><b>4-- Mail Status--TosGeneral-- Has read the mail"
TimeOut 1
SendChat "<font color=#FF3300><b>3-- Account terminated in"
TimeOut 1
SendChat "<font color=#FF3300><b>2-- Is this great or what =)"
TimeOut 1
SendChat "<font color=#FF9900><b>1-- Account Termed"
 SendChat "-æ-[ HyPO ™ ]-æ-Termer By The one and Only...ToaST"

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

