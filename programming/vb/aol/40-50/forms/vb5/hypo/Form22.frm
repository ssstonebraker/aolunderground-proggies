VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phader"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Text            =   "HyPO Phader"
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "Form22.frx":0000
      Left            =   120
      List            =   "Form22.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Pick Your Color"
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Tea$ = "-æ-[ HyPO ™ ]-æ-Incoming Phade so STFU "
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
 If Combo1.Text = "Pick Your Color" Then
 TimeOut 0.7
  SendChat "<font color=#FF0000><b>Im a little slow i forgot to pick a color"
 TimeOut 0.7
SendChat " -æ-[ HyPO ™ ]-æ-" + UserSN + " Is a little Retarded today =)"
  End If
If Combo1.Text = "Red to Blue" Then
  TimeOut 0.7
  SendChat "<font color=#FF0000><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#FF0033><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#FF0066><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#FF0099><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#990099><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#9933CC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#9966FF><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#9999FF><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#99CCFF><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#99FFFF><b>" + Text1
End If
If Combo1.Text = "Black to Green" Then
  TimeOut 0.7
  SendChat "<font color=#000000><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#003300><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#006600><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#009900><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#00CC00><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#00FF00><b>" + Text1
 
End If
If Combo1.Text = "Green to Blue" Then
  TimeOut 0.7
  SendChat "<font color=#00CC00><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#00CC33><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#00CC66><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#00CC99><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#00CCCC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#00CCFF><b>" + Text1
  TimeOut 0.7
 SendChat "<font color=#00CCFF><b>" + Text1
  TimeOut 0.7
 SendChat "<font color=#0099FF><b>" + Text1
  TimeOut 0.7
 SendChat "<font color=#0066FF><b>" + Text1
  TimeOut 0.7
 SendChat "<font color=#0033FF><b>" + Text1
   TimeOut 0.7
 SendChat "<font color=#0000FF><b>" + Text1
 
End If

If Combo1.Text = "Black to Blue" Then
  TimeOut 0.7
  SendChat "<font color=#000000><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#000033><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#000066><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#000099><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#0000CC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#0000FF><b>" + Text1
  End If
  If Combo1.Text = "Hot Green to Light Blue" Then
  TimeOut 0.7
  SendChat "<font color=#99FF00><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#99FF33><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#99FF66><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#99FF99><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#99FFCC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#99FFFF><b>" + Text1
  End If
    If Combo1.Text = "Dark Red to Purple" Then
  TimeOut 0.7
  SendChat "<font color=#990000><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#990033><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#990066><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#990099><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#9900CC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#9900FF><b>" + Text1
  End If
    If Combo1.Text = "Red to Purple" Then
  TimeOut 0.7
  SendChat "<font color=#CC0000><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#CC0033><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#CC0066><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#CC0099><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#CC00CC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#CC00FF><b>" + Text1
  End If
      If Combo1.Text = "Yellow to White" Then
  TimeOut 0.7
  SendChat "<font color=#FFFF00><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#FFFF33><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#FFFF66><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#FFFF99><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#FFFFCC><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#FFFFFF><b>" + Text1
  End If
  
        If Combo1.Text = "Yellow to Black" Then
  TimeOut 0.7
  SendChat "<font color=#FFFF00><b>" + Text1
  TimeOut 0.7
   SendChat "<font color=#CCFF00><b>" + Text1
    TimeOut 0.7
 SendChat "<font color=#99FF00><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#66FF00><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#33FF00><b>" + Text1
 TimeOut 0.7
 SendChat "<font color=#006600><b>" + Text1
  TimeOut 0.7
 SendChat "<font color=#006633><b>" + Text1
  TimeOut 0.7
 SendChat "<font color=#003300><b>" + Text1
   TimeOut 0.7
 SendChat "<font color=#000000><b>" + Text1
  End If
 

 
End Sub

Private Sub Command2_Click()
Form22.Hide
Unload Form22

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
Combo1.AddItem "Red to Blue"
Combo1.AddItem "Black to Green"
Combo1.AddItem "Green to Blue"
Combo1.AddItem "Black to Blue"
Combo1.AddItem "Hot Green to Light Blue"
Combo1.AddItem "Dark Red to Purple"
Combo1.AddItem "Red to Purple"
Combo1.AddItem "Yellow to White"
Combo1.AddItem "Yellow to Black"
 

End Sub
