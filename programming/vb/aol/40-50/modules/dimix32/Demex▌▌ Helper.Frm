VERSION 4.00
Begin VB.Form Dîmîx³²Helper 
   BorderStyle     =   0  'None
   Caption         =   "Dîmîx³² Helper"
   ClientHeight    =   3030
   ClientLeft      =   3135
   ClientTop       =   1995
   ClientWidth     =   5865
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   3435
   Icon            =   "Dîmîx³² Helper.frx":0000
   Left            =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Top             =   1650
   Width           =   5985
   Begin VB.CommandButton Command32 
      Caption         =   "x"
      Height          =   255
      Left            =   5640
      TabIndex        =   48
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command28 
      Caption         =   "_"
      Height          =   255
      Left            =   5400
      TabIndex        =   47
      Top             =   0
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   1920
      TabIndex        =   45
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command31 
      Caption         =   "add room"
      Height          =   255
      Left            =   960
      TabIndex        =   44
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   330
      Left            =   3600
      TabIndex        =   43
      Text            =   "message"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "mass im"
      Height          =   255
      Left            =   2640
      TabIndex        =   42
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      Caption         =   "im me"
      Height          =   255
      Left            =   2040
      TabIndex        =   41
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command27 
      Caption         =   "on"
      Height          =   255
      Left            =   1080
      TabIndex        =   40
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      Caption         =   "off"
      Height          =   255
      Left            =   1560
      TabIndex        =   39
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command25 
      Caption         =   "mail me"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   3000
      TabIndex        =   36
      Text            =   "Subject"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   3960
      TabIndex        =   35
      Text            =   "Msg"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   2040
      TabIndex        =   34
      Text            =   "Person"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "send an email"
      Height          =   375
      Left            =   960
      TabIndex        =   33
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      Caption         =   "italic"
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "underline"
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "bold"
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Caption         =   "dot"
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "wavy"
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Attention"
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "blank sound"
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "bold other"
      Height          =   255
      Left            =   960
      TabIndex        =   24
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "bold first"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "crack"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "send text"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   1920
      TabIndex        =   20
      Text            =   "Text"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "lastchatmsg"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "last sn"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "last line"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "get chat text"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   9360
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   8880
      Top             =   3600
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   4440
      TabIndex        =   13
      Text            =   "Reason"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "off"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "on"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "off"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "on"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   3480
      TabIndex        =   6
      Text            =   "Reason"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "off"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "on"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "off"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "on"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "                                                     Dîmîx³² Helper"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "IMz"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   38
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "                                    Chat Stuff"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Request Bot"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Bot Food"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Afk Bot"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Idle Bot"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Dîmîx³²Helper"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call bot_idle(Label5, True)

End Sub

Private Sub Command10_Click()
MsgBox chat_lastline
End Sub

Private Sub Command11_Click()
MsgBox chat_lastsn
End Sub

Private Sub Command12_Click()
MsgBox chat_lastlinemsg
End Sub


Private Sub Command13_Click()
chatsend Text3
End Sub

Private Sub Command14_Click()
chat_crack Text3
End Sub

Private Sub Command15_Click()
chat_boldfirst Text3
End Sub

Private Sub Command16_Click()
chat_bold2 Text3
End Sub

Private Sub Command17_Click()
chat_blanksound Text3
End Sub

Private Sub Command18_Click()
chat_attention Text3
End Sub

Private Sub Command19_Click()
chat_wavy Text3
End Sub

Private Sub Command2_Click()
Call bot_idle(Label5, False)

End Sub


Private Sub Command20_Click()
chat_dot Text3
End Sub

Private Sub Command21_Click()
chat_bold
End Sub

Private Sub Command22_Click()
chat_underline
End Sub

Private Sub Command23_Click()
chat_italic
End Sub

Private Sub Command26_Click()
im_off
End Sub

Private Sub Command27_Click()
im_on
End Sub

Private Sub Command28_Click()
Me.WindowState = 1

End Sub

Private Sub Command29_Click()
im "macex2000", "Sup MaceX I'm Testing Your Bas"
End Sub

Private Sub Command3_Click()
Call bot_afk(Text1, True)
End Sub

Private Sub Command30_Click()
im_mass List1, Text8
End Sub

Private Sub Command31_Click()
chat_addroom List1
End Sub

Private Sub Command32_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call bot_afk(Text1, False)

End Sub


Private Sub Command5_Click()
chatsend "</b><font color=#ff0000>•^v·• <font color=#000000>Dimix32 ·• Food Bot on"
Timer1 = True
End Sub

Private Sub Command6_Click()
chatsend "</b><font color=#ff0000>•^v·• <font color=#000000>Dimix32 ·• Food Bot off"
Timer1 = False
End Sub


Private Sub Command7_Click()

Timer2 = True
End Sub

Private Sub Command8_Click()

Timer2 = False
End Sub

Private Sub Command9_Click()
MsgBox getchattext
End Sub

Private Sub Form_Load()
chatsend "</b><font color=#ff0000>•^v·• <font color=#000000>Dimix32 ·• Helper Loaded"
form_ontop Me
form_center Me
End Sub


Private Sub Form_Paint()

 form_fadeice Me


End Sub

Private Sub Form_Unload(Cancel As Integer)
chatsend "</b><font color=#ff0000>•^v·• <font color=#000000>Dimix32 ·• Helper UnLoaded"
End
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
form_move Me
End Sub


Private Sub Timer1_Timer()
bot_food
End Sub


Private Sub Timer2_Timer()
bot_request Text2
End Sub


