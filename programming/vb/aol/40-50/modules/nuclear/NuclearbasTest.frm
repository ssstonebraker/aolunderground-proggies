VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command32 
      Caption         =   "PrBust"
      Height          =   375
      Left            =   960
      TabIndex        =   33
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command31 
      Caption         =   "SignOff"
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Chat2"
      Height          =   375
      Left            =   1800
      TabIndex        =   31
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Unload"
      Height          =   375
      Left            =   960
      TabIndex        =   30
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command28 
      Caption         =   "ExitUp"
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command27 
      Caption         =   "ExitDown"
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command26 
      Caption         =   "ExitLeft"
      Height          =   375
      Left            =   960
      TabIndex        =   27
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command25 
      Caption         =   "ExitRight"
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Flash"
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command23 
      Caption         =   "CoolEx"
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command22 
      Caption         =   "IM2"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command21 
      Caption         =   "AddBud"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command20 
      Caption         =   "ShowWel"
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "KillWel"
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "MailSend"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "AddComb"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "ChatSN"
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "LastLine"
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "LimeMSG"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "IMLast"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "IsImOpen"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "IMBud"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SN"
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sn Im"
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "IMText"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "IM"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ChatText"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "AddRoom"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "AOLCnge"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ChatChang"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Chat"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "InRoom"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If InRoom Then
    MsgBox "Your In A Room", vbOKOnly, "Room"
    End If
    If InRoom = 0 Then
    MsgBox "Enter a room", vbOKOnly, "Room"
    End If
    
End Sub
Private Sub Command10_Click()
    SN
    MsgBox SN
    
End Sub
Private Sub Command11_Click()
    IM_Buddy 1, "test"
End Sub
Private Sub Command12_Click()
     If IsImOpen Then
         MsgBox "IM Found", vbOKOnly, "NuclearV1.bas"
     End If
     If IsImOpen = 0 Then
         MsgBox "No IM Found", vbOKOnly, "Nuclear32.bas"
     End If
End Sub
Private Sub Command13_Click()
    MsgBox IMLastMessage
    
End Sub
Private Sub Command14_Click()
    MsgBox ChatLine
    
End Sub
Private Sub Command15_Click()
    MsgBox ChatLineWithSN
    
End Sub
Private Sub Command16_Click()
    MsgBox ChatLineSN
    
End Sub
Private Sub Command17_Click()
    AddRoomToCombo List1, Combo1
    
End Sub
Private Sub Command18_Click()
    MailSender "t3terminat", "test", "Testing mail sender"
    
End Sub
Private Sub Command19_Click()
    HideWelcomeWindow
    
End Sub
Private Sub Command2_Click()
    Chat ("Test")
    TimeOut 0.5
    Chat ("NuclearV1.bas Test")
    
End Sub
Private Sub Command20_Click()
    ShowWelcomeWindow
    
End Sub
Private Sub Command21_Click()
    Add_Bud_ToList "t3terminat"
End Sub
Private Sub Command22_Click()
    IM2 "t3terminat", "test"
End Sub
Private Sub Command23_Click()
    Form_CoolExit Form1
End Sub
Private Sub Command24_Click()
    Form_FlashTitleBar Form1, 100, 120
      
End Sub
Private Sub Command25_Click()
    Form_ExitRight Form1
    
End Sub
Private Sub Command26_Click()
    Form_ExitLeft Form1
End Sub
Private Sub Command27_Click()
    Form_ExitDown Form1
    
End Sub
Private Sub Command28_Click()
    Form_ExitUp Form1
End Sub

Private Sub Command29_Click()
    Chat_Wavy ("Test")
    TimeOut (0.1)
    Chat_Wavy (" Nuclear32.bas Test")
    
End Sub

Private Sub Command3_Click()
    ChangeAOLRoomCaption ("t3terminat")
    
End Sub

Private Sub Command30_Click()
    Form_Implode Form1, 1000
    Unload Form1
End Sub

Private Sub Command31_Click()
    Call Sign_Off(True)
    
End Sub

Private Sub Command32_Click()
    PR_Bust "vb"
    
End Sub

Private Sub Command33_Click()
    Load_Aol
End Sub

Private Sub Command4_Click()
     ChangeAOLCaption ("t3terminat")
End Sub
Private Sub Command5_Click()
     AddRoomToList List1
     
End Sub
Private Sub Command6_Click()
    ChatText
    MsgBox ChatText
    
End Sub
Private Sub Command7_Click()
     IM "t3terminat", "test"
     
End Sub
Private Sub Command8_Click()
    MsgBox IMLastMessageWithSN
    
End Sub

Private Sub Command9_Click()
    ImSn
    MsgBox ImSn
    
End Sub

Private Sub Form_Load()
    Form_Explode Form1, 1000
    
End Sub
