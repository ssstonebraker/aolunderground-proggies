VERSION 5.00
Begin VB.Form frmSwitch 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   1290
      TabIndex        =   7
      Top             =   1965
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstBench 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1200
      ItemData        =   "frmSwitch.frx":0000
      Left            =   210
      List            =   "frmSwitch.frx":0002
      TabIndex        =   0
      Top             =   585
      Width           =   1290
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2580
      Left            =   0
      Top             =   0
      Width           =   1665
   End
   Begin VB.Label lblCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   750
      TabIndex        =   6
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label lblSwitch 
      BackStyle       =   0  'Transparent
      Caption         =   "Switch"
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
      Left            =   60
      TabIndex        =   5
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Switch Pokémon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bench:"
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
      Left            =   210
      TabIndex        =   3
      Top             =   330
      Width           =   1290
   End
   Begin VB.Label lblHealth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
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
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   1800
      Width           =   1290
   End
   Begin VB.Label lblHeaDAT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Top             =   2010
      Width           =   1290
   End
End
Attribute VB_Name = "frmSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Load()
    If Me.Tag = "No" Then
        lblCancel.Visible = False
    End If
    lstBench.Clear
    LoadBench lstBench
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmBattle.Show
End Sub
Private Sub lblCancel_Click()
    Unload Me
End Sub
Private Sub lblSwitch_Click()
    If lstBench.ListIndex = -1 Then
        MsgBoxA Me, "Select a Pokémon to switch to!"
    ElseIf lstBench.ItemData(lstBench.ListIndex) = frmBattle.Pokemon1 Then
        MsgBoxA Me, "That Pokémon is already in battle, choose another!"
    ElseIf GetHealth(lstBench.ItemData(lstBench.ListIndex)) = 0 Then
        MsgBoxA Me, "That Pokémon has fainted, choose another!"
    Else
        If frmBattle.TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "STE-" & lstBench.ItemData(lstBench.ListIndex) & "," & GetHealth(lstBench.ItemData(lstBench.ListIndex))
            frmBattle.Pokemon1 = lstBench.ItemData(lstBench.ListIndex)
            frmBattle.HP1 = GetHealth(lstBench.ItemData(lstBench.ListIndex))
            frmBattle.lblPokemon1.Caption = N2N(frmBattle.Pokemon1)
            frmBattle.imgPokemon1.Picture = frmPokedex.imgList.ListImages.Item(frmBattle.Pokemon1).Picture
            frmBattle.AddHP1 0
            lstBuffer.Clear
            movelist LCase(lstBench.Text), lstBuffer
            frmBattle.lblMove1 = lstBuffer.List(0)
            frmBattle.lblMove2 = lstBuffer.List(1)
            frmBattle.lblMove3 = lstBuffer.List(2)
            frmBattle.Turn = 0
        End If
        If frmBattle.TurnA = 1 Then
            frmInternetListen.sckListen.SendData "STE-" & lstBench.ItemData(lstBench.ListIndex) & "," & GetHealth(lstBench.ItemData(lstBench.ListIndex))
            frmBattle.Pokemon1 = lstBench.ItemData(lstBench.ListIndex)
            frmBattle.HP1 = GetHealth(lstBench.ItemData(lstBench.ListIndex))
            frmBattle.lblPokemon1.Caption = N2N(frmBattle.Pokemon1)
            frmBattle.imgPokemon1.Picture = frmPokedex.imgList.ListImages.Item(frmBattle.Pokemon1).Picture
            frmBattle.AddHP1 0
            lstBuffer.Clear
            movelist LCase(lstBench.Text), lstBuffer
            frmBattle.lblMove1 = lstBuffer.List(0)
            frmBattle.lblMove2 = lstBuffer.List(1)
            frmBattle.lblMove3 = lstBuffer.List(2)
            frmBattle.Turn = 0
        End If
        frmBattle.SetStatus "Go " & N2N(frmBattle.Pokemon1) & "!"
        Unload Me
    End If
End Sub
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lstBench_Click()
    If Not lstBench.ListIndex = -1 Then
        lblHeaDAT.Caption = GetHealth(lstBench.ItemData(lstBench.ListIndex)) & " / " & N2H(lstBench.ItemData(lstBench.ListIndex))
    End If
End Sub
Private Sub lstBench_DblClick()
    lblSwitch_Click
End Sub
