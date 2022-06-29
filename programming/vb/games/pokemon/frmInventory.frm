VERSION 5.00
Begin VB.Form frmInventory 
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
      ItemData        =   "frmInventory.frx":0000
      Left            =   3900
      List            =   "frmInventory.frx":0002
      TabIndex        =   4
      Top             =   645
      Width           =   1125
   End
   Begin VB.ListBox lstItems 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      ItemData        =   "frmInventory.frx":0004
      Left            =   2355
      List            =   "frmInventory.frx":0061
      TabIndex        =   1
      Top             =   150
      Width           =   1545
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
      Left            =   3915
      TabIndex        =   8
      Top             =   1830
      Width           =   1080
   End
   Begin VB.Label lblBench 
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
      Left            =   3900
      TabIndex        =   7
      Top             =   420
      Width           =   1125
   End
   Begin VB.Label lblToss 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Toss"
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
      Left            =   4515
      TabIndex        =   6
      Top             =   2070
      Width           =   495
   End
   Begin VB.Label lblUse 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use"
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
      Left            =   3900
      TabIndex        =   5
      Top             =   2070
      Width           =   495
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2460
      Left            =   0
      Top             =   0
      Width           =   5070
   End
   Begin VB.Label lblOwnDAT 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Left            =   4335
      TabIndex        =   3
      Top             =   165
      Width           =   705
   End
   Begin VB.Label lblOwn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Own:"
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
      Left            =   3885
      TabIndex        =   2
      Top             =   165
      Width           =   465
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
      Picture         =   "frmInventory.frx":011E
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strP As String, strE As String
Private Sub imgBg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub Form_Activate()
    FormOnTop Me
    lstBench.Clear
    LoadBench lstBench
End Sub
Private Sub lblExit_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub lblToss_Click()
    If lstItems.ListIndex = -1 Then
        MsgBoxA Me, "Select an item to toss first!"
    Else
        If lblOwnDAT.Caption = 0 Then
            MsgBoxA Me, "You do not have that item!"
        Else
            DeleteItem lstItems.ListIndex
            lblOwnDAT.Caption = GetItem(lstItems.ListIndex)
        End If
    End If
End Sub
Private Sub lblUse_Click()
    If lstItems.ListIndex = -1 Then
        MsgBoxA Me, "Select an item to use first!"
    ElseIf GetItem(lstItems.ListIndex) = Empty Or GetItem(lstItems.ListIndex) = "0" Then
        MsgBoxA Me, "You do not have any of that item!"
    ElseIf lstBench.ListIndex = -1 Then
        MsgBoxA Me, "Select a Pokémon to use that item on first!"
    ElseIf lstItems.ListIndex <= 5 Then
        MsgBoxA Me, "OAK: " & frmMain.Player & "! This isn't the time to use that!"
    ElseIf lstItems.ListIndex = 6 Then
        strP = lstBench.Text
        If strP = "Charmander" Then
            strE = "Charmeleon"
        ElseIf strP = "Charmeleon" Then
            strE = "Charizard"
        ElseIf strP = "Vulpix" Then
            strE = "Ninetales"
        ElseIf strP = "Growlithe" Then
            strE = "Arcanine"
        ElseIf strP = "Ponyta" Then
            strE = "Rapidash"
        ElseIf strP = "Eevee" Then
            strE = "Flareon"
        Else
            MsgBoxA Me, "This will have no effect."
            Exit Sub
        End If
        strMsg = strP & " evolved into " & strE & "!"
        MsgBoxA Me, strMsg
        SavePokemon Na2N(strE)
        DeletePokemon Na2N(strE)
        DeletePokemon lstBench.ItemData(lstBench.ListIndex)
        DeleteBench lstBench.ItemData(lstBench.ListIndex)
        DeleteItem lstItems.ListIndex
        SaveBench Na2N(LCase(strE))
        lstBench.Clear
        LoadBench lstBench
        lstItems_Click
    ElseIf lstItems.ListIndex = 7 Then
        strP = lstBench.Text
        If strP = "Bulbasaur" Then
            strE = "Ivysaur"
        ElseIf strP = "Ivysaur" Then
            strE = "Venusaur"
        ElseIf strP = "NidoranM" Then
            strE = "Nidorino"
        ElseIf strP = "NidoranF" Then
            strE = "Nidorina"
        ElseIf strP = "Nidorino" Then
            strE = "Nidoking"
        ElseIf strP = "Nidorina" Then
            strE = "Nidoqueen"
        ElseIf strP = "Oddish" Then
            strE = "Gloom"
        ElseIf strP = "Gloom" Then
            strE = "Vileplume"
        ElseIf strP = "Bellsprout" Then
            strE = "Weepinbell"
        ElseIf strP = "Weepinbell" Then
            strE = "Victrebell"
        Else
            MsgBoxA Me, "This will have no effect."
            Exit Sub
        End If
        strMsg = strP & " evolved into " & strE & "!"
        MsgBoxA Me, strMsg
        SavePokemon Na2N(strE)
        DeletePokemon Na2N(strE)
        DeletePokemon lstBench.ItemData(lstBench.ListIndex)
        DeleteBench lstBench.ItemData(lstBench.ListIndex)
        DeleteItem lstItems.ListIndex
        SaveBench Na2N(strE)
        lstBench.Clear
        LoadBench lstBench
        lstItems_Click
    ElseIf lstItems.ListIndex = 8 Then
        strP = lstBench.Text
        If strP = "Pikachu" Then
            strE = "Raichu"
        ElseIf strP = "Magemite" Then
            strE = "Magneton"
         ElseIf strP = "Voltorb" Then
            strE = "Electrode"
        ElseIf strP = "Eevee" Then
            strE = "Jolteon"
        Else
            MsgBoxA Me, "This will have no effect."
            Exit Sub
        End If
        strMsg = strP & " evolved into " & strE & "!"
        MsgBoxA Me, strMsg
        SavePokemon Na2N(strE)
        DeletePokemon Na2N(strE)
        DeletePokemon lstBench.ItemData(lstBench.ListIndex)
        DeleteBench lstBench.ItemData(lstBench.ListIndex)
        DeleteItem lstItems.ListIndex
        SaveBench Na2N(strE)
        lstBench.Clear
        LoadBench lstBench
        lstItems_Click
    ElseIf lstItems.ListIndex = 9 Then
        strP = lstBench.Text
        If strP = "Squirtle" Then
            strE = "Wartortle"
        ElseIf strP = "Wartortle" Then
            strE = "Blastoise"
        ElseIf strP = "Slowpoke" Then
            strE = "Slowbro"
        ElseIf strP = "Seel" Then
            strE = "Dewgong"
        ElseIf strP = "Shellder" Then
            strE = "Cloyster"
        ElseIf strP = "Krabby" Then
            strE = "Kingler"
        ElseIf strP = "Horsea" Then
            strE = "Seadra"
        ElseIf strP = "Goldeen" Then
            strE = "Seaking"
        ElseIf strP = "Staryu" Then
            strE = "Starmie"
        ElseIf strP = "Eevee" Then
            strE = "Vaporeon"
        ElseIf strP = "Omanyte" Then
            strE = "Omastar"
        ElseIf strP = "Kabuto" Then
            strE = "Kabutops"
        ElseIf strP = "Dratini" Then
            strE = "Dragonair"
        ElseIf strP = "Dragonair" Then
            strE = "Dragonite"
        Else
            MsgBoxA Me, "This will have no effect."
            Exit Sub
        End If
        strMsg = strP & " evolved into " & strE & "!"
        MsgBoxA Me, strMsg
        SavePokemon Na2N(strE)
        DeletePokemon Na2N(strE)
        DeletePokemon lstBench.ItemData(lstBench.ListIndex)
        DeleteBench lstBench.ItemData(lstBench.ListIndex)
        DeleteItem lstItems.ListIndex
        SaveBench Na2N(strE)
        lstBench.Clear
        LoadBench lstBench
        lstItems_Click
    ElseIf lstItems.ListIndex = 10 Then
        If GetHealth(Na2N(LCase(lstBench.Text))) = 0 Or GetHealth(Na2N(LCase(lstBench.Text))) = N2H(Na2N(LCase(lstBench.Text))) Then
            MsgBoxA Me, "This will have no effect."
        Else
            If GetHealth(Na2N(LCase(lstBench.Text))) + 10 > N2H(Na2N(LCase(lstBench.Text))) Then
                SetHealth Na2N(LCase(lstBench.Text)), N2H(Na2N(LCase(lstBench.Text)))
            Else
                SetHealth Na2N(LCase(lstBench.Text)), GetHealth(Na2N(LCase(lstBench.Text))) + 10
            End If
            strMsg = lstBench.Text & "'s health restored by 10!"
            MsgBoxA Me, strMsg
            DeleteItem lstItems.ListIndex
            lstBench_Click
            lstItems_Click
        End If
    ElseIf lstItems.ListIndex = 11 Then
        If GetHealth(Na2N(LCase(lstBench.Text))) = 0 Or GetHealth(Na2N(LCase(lstBench.Text))) = N2H(Na2N(LCase(lstBench.Text))) Then
            MsgBoxA Me, "This will have no effect."
        Else
            If GetHealth(Na2N(LCase(lstBench.Text))) + 15 > N2H(Na2N(LCase(lstBench.Text))) Then
                SetHealth Na2N(LCase(lstBench.Text)), N2H(Na2N(LCase(lstBench.Text)))
            Else
                SetHealth Na2N(LCase(lstBench.Text)), GetHealth(Na2N(LCase(lstBench.Text))) + 15
            End If
            strMsg = lstBench.Text & "'s health restored by 15!"
            MsgBoxA Me, strMsg
            DeleteItem lstItems.ListIndex
            lstBench_Click
            lstItems_Click
        End If
    ElseIf lstItems.ListIndex = 12 Then
        If GetHealth(Na2N(LCase(lstBench.Text))) = 0 Or GetHealth(Na2N(LCase(lstBench.Text))) = N2H(Na2N(LCase(lstBench.Text))) Then
            MsgBoxA Me, "This will have no effect."
        Else
            If GetHealth(Na2N(LCase(lstBench.Text))) + 20 > N2H(Na2N(LCase(lstBench.Text))) Then
                SetHealth Na2N(LCase(lstBench.Text)), N2H(Na2N(LCase(lstBench.Text)))
            Else
                SetHealth Na2N(LCase(lstBench.Text)), GetHealth(Na2N(LCase(lstBench.Text))) + 20
            End If
            strMsg = lstBench.Text & "'s health restored by 20!"
            MsgBoxA Me, strMsg
            DeleteItem lstItems.ListIndex
            lstBench_Click
            lstItems_Click
        End If
    ElseIf lstItems.ListIndex = 13 Then
        If GetHealth(Na2N(LCase(lstBench.Text))) = 0 Or GetHealth(Na2N(LCase(lstBench.Text))) = N2H(Na2N(LCase(lstBench.Text))) Then
            MsgBoxA Me, "This will have no effect."
        Else
            SetHealth Na2N(LCase(lstBench.Text)), N2H(Na2N(LCase(lstBench.Text)))
            strMsg = lstBench.Text & "'s health restored to the max!"
            MsgBoxA Me, strMsg
            DeleteItem lstItems.ListIndex
            lstBench_Click
            lstItems_Click
        End If
    ElseIf lstItems.ListIndex = 14 Then
        If Not GetHealth(Na2N(LCase(lstBench.Text))) = 0 Or GetHealth(Na2N(LCase(lstBench.Text))) = N2H(Na2N(LCase(lstBench.Text))) Then
            MsgBoxA Me, "This will have no effect."
        Else
            SetHealth Na2N(LCase(lstBench.Text)), Fix(N2H(Na2N(LCase(lstBench.Text))) / 2)
            strMsg = lstBench.Text & " has been revived!"
            MsgBoxA Me, strMsg
            DeleteItem lstItems.ListIndex
            lstBench_Click
            lstItems_Click
        End If
    End If
End Sub
Private Sub lstBench_Click()
    If lstBench.ListIndex = -1 Then
        lblHeaDAT.Caption = "N/A"
    Else
        lblHeaDAT.Caption = GetFromINI("6" & lstBench.ItemData(lstBench.ListIndex), "6" & lstBench.ItemData(lstBench.ListIndex) & ".1", PathA & "\" & LCase(TrimSpaces(frmMain.Player))) & " / " & N2H(lstBench.ItemData(lstBench.ListIndex))
    End If
End Sub
Private Sub lstItems_Click()
    If GetItem(lstItems.ListIndex) = "" Then
        lblOwnDAT.Caption = "0"
    Else
        lblOwnDAT.Caption = GetItem(lstItems.ListIndex)
    End If
End Sub
