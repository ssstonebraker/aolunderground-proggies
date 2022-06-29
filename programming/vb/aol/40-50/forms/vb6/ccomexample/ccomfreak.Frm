VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   4140
   ClientTop       =   2145
   ClientWidth     =   2940
   Height          =   1170
   Left            =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   2940
   Top             =   1770
   Visible         =   0   'False
   Width           =   3060
   Begin VB.Label Label1 
      Caption         =   "C:\mp3s\"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   240
      Top             =   240
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If LCase(What_Said) Like LCase(".ownz *") Then
What_Said = Mid(What_Said, 7)
Chat1.ChatSend "ˆˆ²~›" & GetUser() & " owns " & What_Said
End If
If LCase(What_Said) Like LCase(".play *") Then
What_Said = Mid(What_Said, 7)
PlayMP3 "" & Label1.Caption & What_Said & ".mp3"
Chat1.ChatSend "ˆˆ²~›now playing: " & What_Said
End If
If LCase(What_Said) Like LCase(".stop *") Then
What_Said = Mid(What_Said, 7)
StopMP3 "" & Label1.Caption & What_Said & ".mp3"
Chat1.ChatSend "ˆˆ²~›now Stoping: " & What_Said
End If
If LCase(What_Said) Like LCase(".x *") Then
What_Said = Mid(What_Said, 4)
IMIgnore "ˆˆ²~›" & What_Said
Chat1.ChatSend "ˆˆ²~›" & What_Said & " Has Been [x]'d so STFU!"
End If
If LCase(What_Said) Like LCase(".unx *") Then
What_Said = Mid(What_Said, 6)
IMUnIgnore "" & What_Said
Chat1.ChatSend "ˆˆ²~›" & What_Said & ", Use ur speach wisely!"
End If
If LCase(What_Said) Like LCase(".lol") Then
Chat1.ChatSend "ˆˆ²~›That was funny, LoLz"
End If
If LCase(What_Said) Like LCase(".lmao") Then
Chat1.ChatSend "ˆˆ²~›Cant....."
Pause 0.5
Chat1.ChatSend "ˆˆ²~›stop....."
Pause 0.5
Chat1.ChatSend "ˆˆ²~›LAUGHING, LoLz!!"
End If
If LCase(What_Said) Like LCase(".rofl") Then
Chat1.ChatSend "ˆˆ²~›Mphfmphf...."
Pause 0.5
Chat1.ChatSend "ˆˆ²~›Mfph..."
Pause 0.5
Chat1.ChatSend "ˆˆ²~›Ah hahaha ah haha, LoLz!"
End If
If LCase(What_Said) Like LCase(".dir *") Then
What_Said = Mid(What_Said, 6)
Label1.Caption = What_Said
End If
If LCase(What_Said) Like LCase(".link color red") Then
Chat1.ChatSend "<body link=#FF0000>ˆˆ²~›links RED "
End If
If LCase(What_Said) Like LCase(".Kick *") Then
What_Said = Mid(What_Said, 7)
Chat1.ChatSend "" + GetUser + " kicked: " + What_Said + " in the nuts!"
Pause 0.5
Chat1.ChatSend "Ouch"
End If

End Sub


Private Sub Form_Load()
Chat1.ScanOn
Chat1.ChatSend "ˆˆ²~›Freaka Comz 2K"
Chat1.ChatSend "ˆˆ²~›By ProgtasticFreaka"
Chat1.ChatSend "ˆˆ²~›Eat my Comz!"
End Sub




Private Sub Form_Unload(Cancel As Integer)
Chat1.ChatSend "ˆˆ²~›Freaka Comz 2K "
Chat1.ChatSend "ˆˆ²~›U 8 me!"
End Sub


