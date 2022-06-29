VERSION 5.00
Begin VB.Form frmBattle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox b2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   915
      Picture         =   "frmBattle.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   3765
      Width           =   225
   End
   Begin VB.PictureBox b3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   645
      Picture         =   "frmBattle.frx":0CC8
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   3765
      Width           =   225
   End
   Begin VB.PictureBox a2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   315
      Picture         =   "frmBattle.frx":195E
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   3765
      Width           =   225
   End
   Begin VB.PictureBox a3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   45
      Picture         =   "frmBattle.frx":290E
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   3765
      Width           =   225
   End
   Begin VB.PictureBox b1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2625
      Picture         =   "frmBattle.frx":391E
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   0
      Width           =   225
   End
   Begin VB.PictureBox a1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2400
      Picture         =   "frmBattle.frx":45B4
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   0
      Width           =   225
   End
   Begin VB.Timer tmrEnabled 
      Interval        =   1
      Left            =   1590
      Top             =   2850
   End
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   1740
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Left            =   0
      Picture         =   "frmBattle.frx":55C4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Caption         =   " Trainer Battle"
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
      Height          =   225
      Left            =   225
      TabIndex        =   20
      Top             =   0
      Width           =   2190
   End
   Begin VB.Image c4 
      Height          =   90
      Left            =   1200
      Picture         =   "frmBattle.frx":5E8E
      Stretch         =   -1  'True
      Top             =   3825
      Width           =   1065
   End
   Begin VB.Image c2 
      Height          =   90
      Left            =   1200
      Picture         =   "frmBattle.frx":6793
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   1065
   End
   Begin VB.Image c3 
      Height          =   90
      Left            =   1200
      Picture         =   "frmBattle.frx":70CF
      Stretch         =   -1  'True
      Top             =   3765
      Width           =   1065
   End
   Begin VB.Label lblChat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M e s s a g e"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   1995
      TabIndex        =   13
      ToolTipText     =   "Open Chat"
      Top             =   2385
      Width           =   825
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   1170
      Width           =   2730
   End
   Begin VB.Label lblMove3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move3"
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
      Left            =   105
      TabIndex        =   10
      Top             =   3180
      Width           =   1965
   End
   Begin VB.Label lblMove2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move2"
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
      Left            =   105
      TabIndex        =   9
      Top             =   2940
      Width           =   1965
   End
   Begin VB.Label lblMove1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move1"
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
      Left            =   105
      TabIndex        =   8
      Top             =   2715
      Width           =   1965
   End
   Begin VB.Label lblRun 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2250
      TabIndex        =   7
      Top             =   3135
      Width           =   510
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2250
      TabIndex        =   6
      Top             =   2925
      Width           =   510
   End
   Begin VB.Label lblPKMN 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PKMN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2250
      TabIndex        =   5
      Top             =   2715
      Width           =   510
   End
   Begin VB.Shape shpControl 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   825
      Left            =   90
      Top             =   2625
      Width           =   2730
   End
   Begin VB.Shape shpFill2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   60
      Left            =   360
      Top             =   660
      Width           =   885
   End
   Begin VB.Label lblPokemon2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pokemon2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   135
      TabIndex        =   4
      Top             =   375
      Width           =   990
   End
   Begin VB.Label lblHPCap2 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   135
      TabIndex        =   3
      Top             =   615
      Width           =   225
   End
   Begin VB.Line lneArrow2a 
      BorderColor     =   &H00000000&
      X1              =   1215
      X2              =   1290
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line lneArrow1a 
      BorderColor     =   &H00000000&
      X1              =   1215
      X2              =   1350
      Y1              =   765
      Y2              =   795
   End
   Begin VB.Line lneBottom1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   1350
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Line lneLeft 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   90
      X2              =   90
      Y1              =   525
      Y2              =   810
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   345
      Top             =   645
      Width           =   915
   End
   Begin VB.Shape shpFill 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   60
      Left            =   1860
      Top             =   2025
      Width           =   885
   End
   Begin VB.Label lblPokemon1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pokemon1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1605
      TabIndex        =   2
      Top             =   1770
      Width           =   990
   End
   Begin VB.Label lblHPCap 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   1620
      TabIndex        =   1
      Top             =   1995
      Width           =   225
   End
   Begin VB.Label lblHP1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "XX / XX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1695
      TabIndex        =   0
      Top             =   2085
      Width           =   1050
   End
   Begin VB.Line lneArrow1 
      BorderColor     =   &H00000000&
      X1              =   1590
      X2              =   1665
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line lneArrow2 
      BorderColor     =   &H00000000&
      X1              =   1530
      X2              =   1665
      Y1              =   2295
      Y2              =   2265
   End
   Begin VB.Line lneBottom 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1545
      X2              =   2775
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line lneRight 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   2805
      X2              =   2805
      Y1              =   2025
      Y2              =   2310
   End
   Begin VB.Shape shpBordera 
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   1845
      Top             =   2010
      Width           =   915
   End
   Begin VB.Image imgPokemon1 
      Height          =   1125
      Left            =   180
      Stretch         =   -1  'True
      Top             =   1455
      Width           =   1170
   End
   Begin VB.Image imgPokemon2 
      Height          =   765
      Left            =   1905
      Stretch         =   -1  'True
      Top             =   300
      Width           =   810
   End
   Begin VB.Image c1 
      Height          =   225
      Left            =   0
      Picture         =   "frmBattle.frx":7A81
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HP1 As Integer, HP2 As Integer, Pokemon1 As Integer, Pokemon2 As Integer, Move1 As String, Move2 As String, Move3 As String, ForeignID As String, Turn As Integer, TurnA As Integer
Private bOn As Boolean, b1On As Boolean, b2On As Boolean
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_GotFocus()
    c1.Picture = c3.Picture
End Sub
Private Sub Form_LostFocus()
    c1.Picture = c4.Picture
End Sub
Private Sub Form_Load()
    If Not TurnA = 3 Then TurnA = Turn
    If TurnA = 0 Then frmInternetConnect.Tag = "battle"
    lblHP1.Caption = HP1 & " / " & HP1
    lblPokemon1.Caption = N2N(Pokemon1)
    lblPokemon2.Caption = N2N(Pokemon2)
    imgPokemon1.Picture = frmPokedex.imgList.ListImages.Item(Pokemon1).Picture
    imgPokemon2.Picture = frmPokedex.imgList.ListImages.Item(Pokemon2).Picture
    lblMove1.Caption = UCase(Left(Move1, 1)) + Right(Move1, Len(Move1) - 1) + ", " + N2N(Pokemon1) + "!"
    lblMove2.Caption = UCase(Left(Move2, 1)) + Right(Move2, Len(Move2) - 1) + "! You can do it " + N2N(Pokemon1) + "!"
    lblMove3.Caption = UCase(Left(Move3, 1)) + Right(Move3, Len(Move3) - 1) + " now " + N2N(Pokemon1) + "!"
    AddHP1 0
    AddHP2 0
    If N2E(Pokemon1) = "Power" Then
        lblMove3.Caption = UCase(Left(Move3, 1)) & Right(Move3, Len(Move3) - 1)
    ElseIf N2E(Pokemon1) = Poison Then
        lblMove3.Caption = "Poison Powder"
    ElseIf N2E(Pokemon1) = Freeze Then
        lblMove3.Caption = "Ice Beam"
    ElseIf N2E(Pokemon1) = Burn Then
        lblMove3.Caption = "Fire Blast"
    ElseIf N2E(Pokemon1) = Confuse Then
        lblMove3.Caption = "Confuse Ray"
    ElseIf N2E(Pokemon1) = Paralyz Then
        lblMove3.Caption = "Thunder"
    ElseIf N2E(Pokemon1) = Sleep Then
        lblMove3.Caption = "Sing"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload frmChatroom
End Sub
Private Sub lblChat_Click()
    FormNotOnTop Me
    If MsgBox("Reset the Incoming Message alert?" + vbNewLine + "(chat log will be lost)", vbYesNo + vbQuestion, "Reset Alert?") = vbYes Then
        lblChat.ForeColor = &HFF0000
        frmChatroom.txtChat.Text = Empty
    End If
End Sub
Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    c1.Picture = c2.Picture
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    c1.Picture = c3.Picture
End Sub
Private Sub lblItem_Click()
    Me.Hide
    frmUseItem.Show
End Sub
Private Sub lblMinimize_Click()
    Me.WindowState = 1
End Sub
Private Sub lblMove1_Click()
    lblMove1.ForeColor = &HFF&
    lblMove2.ForeColor = &HFF0000
    lblMove3.ForeColor = &HFF0000
End Sub
Private Sub lblMove1_DblClick()
    intMiss = Random(1, 3)
    If intMiss > 2 Then
        If TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "MISS"
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "MISS"
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "MISS"
        End If
        Turn = 0
        SetStatus N2N(frmBattle.Pokemon1) & " missed!"
    Else
        intAttack = Random(1, 6)
        If TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "MIN-" & intAttack
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "MIN-" & intAttack
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "MIN-" & intAttack
        End If
        MinusHP2 intAttack
        Turn = 0
        SetStatus N2N(frmBattle.Pokemon1) & " hit for " & intAttack & " HP!"
    End If
End Sub
Private Sub lblMove2_Click()
    lblMove1.ForeColor = &HFF0000
    lblMove2.ForeColor = &HFF&
    lblMove3.ForeColor = &HFF0000
End Sub
Private Sub lblMove2_DblClick()
    intMiss = Random(1, 3)
    If intMiss > 2 Then
        If TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "MISS"
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "MISS"
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "MISS"
        End If
        Turn = 0
        SetStatus N2N(frmBattle.Pokemon1) & " missed!"
    Else
        intAttack = Random(1, 9)
        If TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "MIN-" & intAttack
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "MIN-" & intAttack
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "MIN-" & intAttack
        End If
        MinusHP2 intAttack
        Turn = 0
        SetStatus N2N(frmBattle.Pokemon1) & " hit for " & intAttack & " HP!"
    End If
End Sub
Private Sub lblMove3_Click()
    lblMove1.ForeColor = &HFF0000
    lblMove2.ForeColor = &HFF0000
    lblMove3.ForeColor = &HFF&
End Sub
Private Sub lblMove3_DblClick()
    intMiss = Random(1, 3)
    If intMiss > 2 Then
        If TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "MISS"
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "MISS"
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "MISS"
        End If
        Turn = 0
        SetStatus N2N(frmBattle.Pokemon1) & " missed!"
    Else
        intAttack = Random(1, 11)
        If TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "MIN-" & intAttack
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "MIN-" & intAttack
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "MIN-" & intAttack
        End If
        MinusHP2 intAttack
        Turn = 0
        SetStatus N2N(frmBattle.Pokemon1) & " hit for " & intAttack & " HP!"
    End If
End Sub
Private Sub lblPKMN_Click()
    Me.Hide
    frmSwitch.Show
End Sub
Private Sub lblRun_Click()
    If GetMoney - 500 < 0 Then
        lblCaption = "Lost $" & GetMoney & " For Losing"
        DeleteMoney GetMoney
        If frmBattle.TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "CAS-" & GetMoney
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "CAS-" & GetMoney
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "CAS-" & GetMoney
        End If
    Else
        lblCaption = "Lost $500 For Losing"
        DeleteMoney 500
        If frmBattle.TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "CAS-500"
        ElseIf TurnA = 1 Then
            frmInternetListen.sckListen.SendData "CAS-500"
        ElseIf TurnA = 2 Then
            frmInternetConnect.sckConnect.SendData "re-" + "CAS-500"
        End If
    End If
    frmLose.Show
    Me.Hide
End Sub
Public Function MinusHP1(num)
    HP1% = HP1% - num
    If Not HP1% <= 0 Then
        perc1% = 885 / N2H(Pokemon1)
        shpFill.Width = perc1% * HP1%
        lblHP1.Caption = HP1% & " / " & N2H(Pokemon1)
        SetHealth Pokemon1, HP1%
    Else
        shpFill.Width = 0
        lblHP1.Caption = 0 & " / " & N2H(Pokemon1)
        HP1% = 0
        SetHealth Pokemon1, HP1%
        If LosePlayer = True Then
            Me.Hide
            frmLose.Show
        Else
            Me.Hide
            frmSwitch.Tag = "No"
            frmSwitch.Show
            If TurnA = 0 Then
                frmInternetConnect.sckConnect.SendData "SLOSE"
            ElseIf TurnA = 1 Then
                frmInternetListen.sckListen.SendData "SLOSE"
            ElseIf TurnA = 2 Then
                frmInternetConnect.sckConnect.SendData "re-" + "SLOSE"
            End If
        End If
    End If
End Function
Public Function AddHP1(num)
    HP1% = HP1% + num
    If Not HP1% > N2H(Pokemon1) Then
        perc1% = 885 / N2H(Pokemon1)
        shpFill.Width = perc1% * HP1%
        lblHP1.Caption = HP1% & " / " & N2H(Pokemon1)
    Else
        HP1% = N2H(Pokemon1)
        perc1% = 885 / N2H(Pokemon1)
        shpFill.Width = perc1% * HP1%
        lblHP1.Caption = HP1% & " / " & N2H(Pokemon1)
    End If
End Function
Public Function MinusHP2(num)
    HP2% = HP2% - num
    If Not HP2% <= 0 Then
        perc1% = 885 / N2H(Pokemon2)
        shpFill2.Width = perc1% * HP2%
    Else
        shpFill2.Width = 0
        HP2% = 0
    End If
End Function
Public Function AddHP2(num)
    HP2% = HP2% - num
    If Not HP2% > N2H(Pokemon2) Then
        perc1% = 885 / N2H(Pokemon2)
        shpFill2.Width = perc1% * HP2%
    Else
        HP2% = N2H(Pokemon2)
        perc1% = 885 / N2H(Pokemon2)
        shpFill2.Width = perc1% * HP2%
    End If
End Function
Function LosePlayer() As Boolean
    lstBuffer.Clear
    LoadBench lstBuffer
    nm$ = 0
    nm1$ = 0
str:
    If nm$ = lstBuffer.ListCount Then
        If nm1$ = lstBuffer.ListCount Then
            LosePlayer = True
        Else
            LosePlayer = False
        End If
        Exit Function
    Else
        If GetHealth(lstBuffer.ItemData(nm$)) = 0 Then
            nm1$ = nm1$ + 1
        End If
        nm$ = nm$ + 1
        GoTo str
    End If
End Function
Private Sub tmrEnabled_Timer()
    If Turn = "1" Then
        lblMove1.Visible = True
        lblMove2.Visible = True
        lblMove3.Visible = True
        lblItem.Visible = True
        lblPKMN.Visible = True
        lblRun.Visible = True
    ElseIf Turn = "0" Then
        lblMove1.Visible = False
        lblMove2.Visible = False
        lblMove3.Visible = False
        lblItem.Visible = False
        lblPKMN.Visible = False
        lblRun.Visible = False
    End If
End Sub
Public Function SetStatus(strStatus As String)
    lblStatus.Caption = strStatus
    TimeOut 1
End Function
Private Sub a1_Click()
    bOn = False
    frmChatroom.Show
End Sub
Private Sub a1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bOn Then
        bOn = True
        a1.Picture = a2.Picture
        SetCapture a1.hwnd
    ElseIf X < 0 Or Y < 0 Or X > a1.Width Or Y > a1.Height Then
        bOn = False
        a1.Picture = a3.Picture
        ReleaseCapture
    End If
End Sub
Private Sub b1_Click()
    b1On = False
    Me.WindowState = vbMinimized
    frmChatroom.Hide
End Sub
Private Sub b1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not b1On Then
        b1On = True
        b1.Picture = b2.Picture
        SetCapture b1.hwnd
    ElseIf X < 0 Or Y < 0 Or X > b1.Width Or Y > b1.Height Then
        b1On = False
        b1.Picture = b3.Picture
        ReleaseCapture
    End If
End Sub

