VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form BattleMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Battle Net Chat "
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   Picture         =   "BattleMain.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   7680
      Picture         =   "BattleMain.frx":43842
      ScaleHeight     =   3000
      ScaleWidth      =   1650
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   1650
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Who is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Insert in Loop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Add in list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rejoin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unban Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ban Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   1680
      TabIndex        =   0
      Top             =   6360
      Width           =   4695
   End
   Begin VB.ListBox Roomlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3405
      ItemData        =   "BattleMain.frx":45AF1
      Left            =   6720
      List            =   "BattleMain.frx":45AF3
      MousePointer    =   1  'Arrow
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   9
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Whisper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6900
      TabIndex        =   8
      Top             =   6720
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   435
      Index           =   6
      Left            =   8040
      Picture         =   "BattleMain.frx":45AF5
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   435
      Index           =   5
      Left            =   6720
      Picture         =   "BattleMain.frx":49037
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1275
   End
   Begin VB.Label Channnel 
      BackStyle       =   0  'Transparent
      Caption         =   "Channel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   280
      TabIndex        =   7
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Option 
      BackStyle       =   0  'Transparent
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Quit 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Disconnect 
      BackStyle       =   0  'Transparent
      Caption         =   "Discon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Connect 
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   4
      Left            =   120
      Picture         =   "BattleMain.frx":4C579
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   3
      Left            =   120
      Picture         =   "BattleMain.frx":4FABB
      Top             =   5160
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   2
      Left            =   120
      Picture         =   "BattleMain.frx":52FFD
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   1
      Left            =   120
      Picture         =   "BattleMain.frx":5653F
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   0
      Left            =   120
      Picture         =   "BattleMain.frx":59A81
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label Channel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chat 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   0
      Picture         =   "BattleMain.frx":5CFC3
      Top             =   1920
      Width           =   9600
   End
End
Attribute VB_Name = "BattleMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ind As Long
Dim Channelstring
Dim OldIndex
Private Sub Channnel_Click()
FrmChannel.Show 1
End Sub
Private Sub Connect_Click()
  On Error Resume Next
      With Winsock1
           .Protocol = sckTCPProtocol
           .RemoteHost = "Battle.net"
           .RemotePort = "6112"
           .Connect
           ConnectedToServer = True
      End With
     EnableDisableCommand False, True, True, True
     FrmLogin.Show 1
End Sub
Private Sub Disconnect_Click()
  ConnectedToServer = False
  Timer1.Interval = 0
  EnableDisableCommand True, False, False, True
  Roomlist.Clear
  Text2.Text = ""
  DoEvents
  Winsock1.Close
End Sub
Private Sub EnableDisableCommand(p1, p2, p3, p4)
  Connect.Enabled = p1
  Disconnect.Enabled = p3
  Quit.Enabled = p4
  Channnel.Enabled = p2
End Sub
Private Sub Form_Load()
 EnableDisableCommand True, False, False, True
End Sub

Public Sub ProceedLogin()
   MousePointer = 11
   With Winsock1
     .SendData Chr(3) & Chr(4) & U_ID & Chr(13) & Chr(10) & U_Pass & Chr(13) & Chr(10)
     .SendData Username & Chr(13) & Chr(10)
     .SendData Password & Chr(13) & Chr(10)
   End With
DoEvents
Timer1.Interval = 500
End Sub

Private Sub Image1_Click()
Picture1.Visible = False
End Sub

Private Sub Label3_Click()
ExecuteCode ("/Whisper " & Roomlist.List(Roomlist.ListIndex) & " " & Text1.Text)
End Sub
Private Sub ExecuteCode(codep)
Text1.Text = codep
Send_Click
End Sub
Private Sub Label4_Click()
Picture1.Visible = Not Picture1.Visible
End Sub

Private Sub Label5_Click(Index As Integer)
Select Case Index
 Case 0: ExecuteCode ("/ban " & Roomlist.List(Roomlist.ListIndex))
 Case 1: ExecuteCode ("/unban " & Roomlist.List(Roomlist.ListIndex))
 Case 2: ExecuteCode ("/rejoin ")
 Case 3:
 Case 4:
 Case 5: ExecuteCode ("/stats " & Roomlist.List(Roomlist.ListIndex)) & " STAR"
 Case 6: ExecuteCode ("/whois " & Roomlist.List(Roomlist.ListIndex))
End Select

End Sub


Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> OldIndex Then
 Label5(OldIndex).BackStyle = 0
 Label5(Index).BackStyle = 1
 OldIndex = Index
End If
End Sub

Private Sub Option_Click()
MsgBox "FrmOpt.Show 1"
End Sub

Private Sub Quit_Click()
  End
End Sub

Public Sub Send_Click()
  If Left(Text1.Text, 1) <> "/" Then
    Dim ok As Boolean
    ok = False
    For i = 0 To Roomlist.ListCount - 1
     If Roomlist.Selected(i) Then
      ok = True
      textsend = "/Whisper " & Roomlist.List(i) & " " & Text1.Text
      Winsock1.SendData textsend & Chr(13) & Chr(10)
      Text2.Text = Text2.Text & Username & "TO: " & Text1.Text & Chr(13) & Chr(10)
      Text2.SelStart = Len(Text2)
     End If
    Next i
   If Not ok Then
     Winsock1.SendData Text1.Text & Chr(13) & Chr(10)
     If Left(Text1.Text, 2) = "/W" Then
       Text2.Text = Text2.Text & Username & "TO: " & Mid(Text1.Text, 10, Len(Text1.Text)) & Chr(13) & Chr(10)
     Else
       Text2.Text = Text2.Text & Username & ": " & Text1.Text & Chr(13) & Chr(10)
     End If
     Text2.SelStart = Len(Text2)
   End If
  Else
  Winsock1.SendData Text1.Text & Chr(13) & Chr(10)
  End If
  Text1.Text = ""
End Sub
Private Function RetreiveCodePos(strdata) As Long
  mypost = InStr(1, strdata, "0010", vbBinaryCompare)
  If mypost <> 0 Then RetreiveCodePos = mypost
  mypost = InStr(1, strdata, "0000")
  If mypost <> 0 Then RetreiveCodePos = mypost
  mypost = InStr(1, strdata, "0002")
  If mypost <> 0 Then RetreiveCodePos = mypost
  mypost = InStr(1, strdata, "0012")
  If mypost <> 0 Then RetreiveCodePos = mypost
End Function

Private Sub Roomlist_Click()
Picture1.Visible = False
End Sub

Private Sub Text1_Click()
Picture1.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 Send_Click
End If
End Sub
Private Sub Text2_Click()
Picture1.Visible = False
End Sub

Private Sub Timer1_Timer()
 If Winsock1.State <> 7 And ConnectedToServer Then
  MsgBox "You have lost your connection to Battle.net"
  Disconnect_Click
 End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  Dim BattleData As String
  Winsock1.GetData BattleData, vbString
  DoEvents
  ExecuteData BattleData
  DoEvents
  BattleData = ""
 End Sub
Private Sub ExecuteData(ByVal batdat)
  Dim start As Integer
  Dim strcheck As String
  Dim strcheck2 As String
  start = 1
  Do
    pos1 = InStr(start, batdat, Chr(13), vbBinaryCompare)
    If pos1 <> 0 Then
      strcheck = Mid(batdat, start, pos1 - start + 1)
      X = pos1
      start = pos1 + 2
      strcheck2 = Left(strcheck, 4)
      Select Case strcheck2
        Case "1001": CodeUser strcheck       'User already in the Room
        Case "1002": CodeJoin strcheck       'User Join the Room
        Case "1003": CodeLeave strcheck      'User leave the Room
        Case "1004": CodeWhisper strcheck    'User whisper you
        Case "1005": CodeTalk strcheck       'User Talk (chat text)
        Case "1007": CodeChannel strcheck    'Current Channel
        Case "1018": CodeInfo strcheck       'Info from Battle.net
        Case "1023": CodeEmote strcheck      'Emote user
        Case "1019":                         'Error
        Case "2000":                         'NULL
        Case "2010":                         'Your logged name
      End Select
      DoEvents
    Else
      Exit Do
    End If
  Loop Until X = Len(batdat) - 1
End Sub

'=============================================================
'************* BATTLE NET WINSOCK CODE FUNCTION **************
'=============================================================
Private Sub CodeTalk(strdata)
  mypos = RetreiveCodePos(strdata)
  strname = Mid(strdata, 11, mypos - 11)
  strchat = Mid(strdata, mypos + 6, Len(strdata) - (mypos + 6) - 1)
  Text2.Text = Text2.Text & strname & ": " & strchat & Chr(13) & Chr(10)
  Text2.SelStart = Len(Text2.Text)
End Sub
Private Sub CodeWhisper(strdata)
   mypos = RetreiveCodePos(strdata)
   mypos = mypos - 13
   strname = Mid(strdata, 13, mypos)
   chatstart = InStr(1, strdata, Chr(34))
   chatend = InStr(chatstart + 1, strdata, Chr(34))
   chatend = chatend - chatstart
   strchat = Mid(strdata, chatstart + 1, chatend - 1)
   Text2.Text = Text2.Text & strname & ": " & "(whisper) " & strchat & Chr(13) & Chr(10)
   Text2.SelStart = Len(Text2.Text)
End Sub
Private Sub CodeInfo(strdata)
  On Error Resume Next
  mypos = InStr(1, strdata, Chr(34))
  mypos2 = InStr(mypos + 1, strdata, Chr(34))
  mypos2 = mypos2 - mypos
  strchat = Mid(strdata, mypos + 1, mypos2 - 1)
  Text2.Text = Text2.Text & strchat & Chr(13) & Chr(10)
  Text2.SelStart = Len(Text2.Text)
End Sub
Private Sub CodeEmote(strdata)
  mypos = RetreiveCodePos(strdata)
  strname = Mid(strdata, 12, mypos - 12)
  strchat = Mid(strdata, mypos + 6, Len(strdata) - (mypos + 6) - 1)
  Text2.Text = Text2.Text & strname & " " & strchat & Chr(13) & Chr(10)
  Text2.SelStart = Len(Text2.Text)
End Sub
Private Sub CodeJoin(strdata)
  mypos = RetreiveCodePos(strdata)
  strname = Mid(strdata, 11, mypos - 11) '& Mid(strdata, mypos + 4, Len(strdata) - (mypos + 6) + 2)
  Roomlist.AddItem strname
  Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
End Sub
Private Sub CodeLeave(strdata)
'On Error Resume Next
  Dim founded As Boolean
  founded = False
  mypos = RetreiveCodePos(strdata)
  strname = Trim(Mid(strdata, 11, mypos - 11))
  For i = 0 To Roomlist.ListCount
    If strname = Trim(Roomlist.List(i)) Then
      founded = True
      Exit For
    End If
  Next i
  If founded Then Roomlist.RemoveItem i
  Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
End Sub
Private Sub CodeUser(strdata)
  mypos = RetreiveCodePos(strdata)
  strname = Mid(strdata, 11, mypos - 11) '& Mid(strdata, mypos + 4, Len(strdata) - (mypos + 6) + 2)
  Roomlist.AddItem strname
  Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
End Sub

Private Sub CodeChannel(strdata)
  Roomlist.Clear
  MousePointer = 1
  strname = Mid(strdata, 15, Len(strdata) - 16)
  Channelstring = strname
  Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
End Sub

