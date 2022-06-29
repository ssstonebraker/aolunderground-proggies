VERSION 5.00
Begin VB.Form frmPokemonCenter 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Instant Messages"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   2385
      TabIndex        =   8
      Top             =   2070
      Width           =   2100
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Caption         =   "Send Instant Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   2385
      TabIndex        =   7
      Top             =   1830
      Width           =   1920
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2460
      Left            =   0
      Top             =   0
      Width           =   5070
   End
   Begin VB.Image imgBall1 
      Height          =   480
      Left            =   2355
      Picture         =   "frmPokemonCenter.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "Talk to Breeder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2370
      TabIndex        =   5
      Top             =   1605
      Width           =   2010
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Talk to Nurse Joy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2370
      TabIndex        =   4
      Top             =   1365
      Width           =   2010
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "Talk to Trainer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2355
      TabIndex        =   3
      Top             =   1125
      Width           =   2010
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cable Club"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2370
      TabIndex        =   2
      Top             =   885
      Width           =   2010
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Heal Benched Pokémon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2370
      TabIndex        =   1
      Top             =   645
      Width           =   2010
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   2070
      Width           =   135
   End
   Begin VB.Shape shpBgBorder 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   2220
      Left            =   120
      Top             =   120
      Width           =   2220
   End
   Begin VB.Image imgBg 
      Height          =   2160
      Left            =   150
      Picture         =   "frmPokemonCenter.frx":030A
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmPokemonCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgBg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub lbl1_Click()
    lbl1.ForeColor = &HFFFFFF
    lbl2.ForeColor = &HFF0000
    lbl3.ForeColor = &HFF0000
    lbl4.ForeColor = &HFF0000
    lbl5.ForeColor = &HFF0000
    lbl6.ForeColor = &HC0C0C0
    lbl7.ForeColor = &HC0C0C0
End Sub
Private Sub lbl1_DblClick()
    lstBuffer.Clear
    LoadBench lstBuffer
    If lstBuffer.ListCount = 0 Then
        MsgBoxA Me, "No benched Pokémon to heal!"
    Else
        If lstBuffer.ListCount = 1 Then
            MaxHealth lstBuffer.ItemData(0)
        End If
        If lstBuffer.ListCount = 2 Then
            MaxHealth lstBuffer.ItemData(0)
            MaxHealth lstBuffer.ItemData(1)
        End If
        If lstBuffer.ListCount = 3 Then
            MaxHealth lstBuffer.ItemData(0)
            MaxHealth lstBuffer.ItemData(1)
            MaxHealth lstBuffer.ItemData(2)
        End If
        If lstBuffer.ListCount = 4 Then
            MaxHealth lstBuffer.ItemData(0)
            MaxHealth lstBuffer.ItemData(1)
            MaxHealth lstBuffer.ItemData(2)
            MaxHealth lstBuffer.ItemData(3)
        End If
        If lstBuffer.ListCount = 5 Then
            MaxHealth lstBuffer.ItemData(0)
            MaxHealth lstBuffer.ItemData(1)
            MaxHealth lstBuffer.ItemData(2)
            MaxHealth lstBuffer.ItemData(3)
            MaxHealth lstBuffer.ItemData(4)
        End If
        If lstBuffer.ListCount = 6 Then
            MaxHealth lstBuffer.ItemData(0)
            MaxHealth lstBuffer.ItemData(1)
            MaxHealth lstBuffer.ItemData(2)
            MaxHealth lstBuffer.ItemData(3)
            MaxHealth lstBuffer.ItemData(4)
            MaxHealth lstBuffer.ItemData(5)
        End If
        MsgBoxA Me, "Benched Pokémon Healed!"
    End If
End Sub
Private Sub lbl2_Click()
    lbl1.ForeColor = &HFF0000
    lbl2.ForeColor = &HFFFFFF
    lbl3.ForeColor = &HFF0000
    lbl4.ForeColor = &HFF0000
    lbl5.ForeColor = &HFF0000
    lbl6.ForeColor = &HC0C0C0
    lbl7.ForeColor = &HC0C0C0
End Sub
Private Sub lbl2_DblClick()
    Me.Hide
    frmCableClub.Show
End Sub
Private Sub lbl3_Click()
    lbl1.ForeColor = &HFF0000
    lbl2.ForeColor = &HFF0000
    lbl3.ForeColor = &HFFFFFF
    lbl4.ForeColor = &HFF0000
    lbl5.ForeColor = &HFF0000
    lbl6.ForeColor = &HC0C0C0
    lbl7.ForeColor = &HC0C0C0
End Sub
Private Sub lbl3_DblClick()
    intQuote = Random(1, 7)
    Select Case intQuote
    Case 1
        MsgBoxA Me, "Be sure to collect evolutions of different Pokémon!"
    Case 2
        MsgBoxA Me, "You must beat all 8 gym leaders before going up against the new Pokémon league."
    Case 3
        MsgBoxA Me, "Don't tell anyone, but I heard Team Rocket is controlling most of the gyms."
    Case 4
        MsgBoxA Me, "If anyone asks, tell them I don't know anything about Team Rocket."
    Case 5
        MsgBoxA Me, "Viridian Forest is filled with Team Rocket Pokémon bunkers, be careful!"
    Case 6
        MsgBoxA Me, "Team Rocket's objective is to control Pokémon Island, but thats not possible if we don't let them!"
    Case 7
        MsgBoxA Me, "Pss.... did Oak send you? Yes? Good, theres an ATR Base in Cerulean City."
    End Select
End Sub
Private Sub lbl4_Click()
    lbl1.ForeColor = &HFF0000
    lbl2.ForeColor = &HFF0000
    lbl3.ForeColor = &HFF0000
    lbl4.ForeColor = &HFFFFFF
    lbl5.ForeColor = &HFF0000
    lbl6.ForeColor = &HC0C0C0
    lbl7.ForeColor = &HC0C0C0
End Sub
Private Sub lbl4_DblClick()
    intQuote = Random(1, 4)
    Select Case intQuote
    Case 1
        MsgBoxA Me, "Sorry " & frmMain.Player & ", but I can't talk right now."
    Case 2
        MsgBoxA Me, "Come and visit me or any of my sisters to heal your benched Pokémon!"
    Case 3
        MsgBoxA Me, "Chansey is the most friendly Pokémon, get it as soon as you can!"
    Case 4
        MsgBoxA Me, "If you ever need to Trade or Battle someone, here is the place to do it!"
    End Select
End Sub
Private Sub lbl5_Click()
    lbl1.ForeColor = &HFF0000
    lbl2.ForeColor = &HFF0000
    lbl3.ForeColor = &HFF0000
    lbl4.ForeColor = &HFF0000
    lbl5.ForeColor = &HFFFFFF
    lbl6.ForeColor = &HC0C0C0
    lbl7.ForeColor = &HC0C0C0
End Sub
Private Sub lbl5_DblClick()
    intQuote = Random(1, 4)
    Select Case intQuote
    Case 1
        MsgBoxA Me, "Feeding Pokémon the right food makes them happy!"
    Case 2
        MsgBoxA Me, "When you breed a Mew with another speices, it becomes a Mewtwo no matter what."
    Case 3
        MsgBoxA Me, "Brock from Pewter city makes the best Pokémon food."
    Case 4
        MsgBoxA Me, "Misty from Cerulean City is one of the few people who have bikes on Pokémon Island."
    End Select
End Sub
'Private Sub lbl6_Click()
'    lbl1.ForeColor = &HFF0000
'    lbl2.ForeColor = &HFF0000
'    lbl3.ForeColor = &HFF0000
'    lbl4.ForeColor = &HFF0000
'    lbl5.ForeColor = &HFF0000
'    lbl6.ForeColor = &HFFFFFF
'    lbl7.ForeColor = &HFF0000
'End Sub
'Private Sub lbl7_Click()
'    lbl1.ForeColor = &HFF0000
'    lbl2.ForeColor = &HFF0000
'    lbl3.ForeColor = &HFF0000
'    lbl4.ForeColor = &HFF0000
'    lbl5.ForeColor = &HFF0000
'    lbl6.ForeColor = &HFF0000
'    lbl7.ForeColor = &HFFFFFF
'End Sub
Private Sub lblExit_Click()
    Unload Me
End Sub
