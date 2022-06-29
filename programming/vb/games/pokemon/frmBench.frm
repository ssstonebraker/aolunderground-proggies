VERSION 5.00
Begin VB.Form frmBench 
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
      ItemData        =   "frmBench.frx":0000
      Left            =   3720
      List            =   "frmBench.frx":0002
      TabIndex        =   2
      Top             =   375
      Width           =   1290
   End
   Begin VB.ListBox lstPokemon 
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
      Height          =   1980
      ItemData        =   "frmBench.frx":0004
      Left            =   2370
      List            =   "frmBench.frx":0006
      TabIndex        =   1
      Top             =   375
      Width           =   1290
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2460
      Left            =   0
      Top             =   0
      Width           =   5070
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
      Left            =   3720
      TabIndex        =   6
      Top             =   1860
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1590
      Width           =   1290
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
      Left            =   3750
      TabIndex        =   4
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label lblComputer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Computer:"
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
      Left            =   2370
      TabIndex        =   3
      Top             =   120
      Width           =   1290
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
      Picture         =   "frmBench.frx":0008
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmBench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub imgBg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub Form_Activate()
    FormOnTop Me
    lstPokemon.Clear
    lstBench.Clear
    LoadPokemon lstPokemon
    LoadBench lstBench
End Sub
Private Sub lblExit_Click()
    Unload Me
End Sub
Private Sub lstBench_Click()
    If lstBench.ListIndex = -1 Then
        lblHeaDAT.Caption = "N/A"
    Else
        lblHeaDAT.Caption = GetFromINI("6" & lstBench.ItemData(lstBench.ListIndex), "6" & lstBench.ItemData(lstBench.ListIndex) & ".1", PathA & "\" & LCase(TrimSpaces(frmMain.Player))) & " / " & N2H(lstBench.ItemData(lstBench.ListIndex))
    End If
End Sub
Private Sub lstBench_DblClick()
    SavePokemon lstBench.ItemData(lstBench.ListIndex)
    DeleteBench lstBench.ItemData(lstBench.ListIndex)
    lstPokemon.Clear
    lstBench.Clear
    LoadPokemon lstPokemon
    LoadBench lstBench
End Sub
Private Sub lstPokemon_Click()
    If lstPokemon.ListIndex = -1 Then
        lblHeaDAT.Caption = "N/A"
    Else
        lblHeaDAT.Caption = GetFromINI("6" & lstPokemon.ItemData(lstPokemon.ListIndex), "6" & lstPokemon.ItemData(lstPokemon.ListIndex) & ".1", PathA & "\" & LCase(TrimSpaces(frmMain.Player))) & " / " & N2H(lstPokemon.ItemData(lstPokemon.ListIndex))
    End If
End Sub
Private Sub lstPokemon_DblClick()
    If lstBench.ListCount = 6 Then
        MsgBoxA Me, "Bench is full! Remove a Pokémon from your bench to add a different Pokémon you have."
    Else
        DeletePokemon lstPokemon.ItemData(lstPokemon.ListIndex)
        SaveBench lstPokemon.ItemData(lstPokemon.ListIndex)
        lstPokemon.Clear
        lstBench.Clear
        LoadPokemon lstPokemon
        LoadBench lstBench
    End If
End Sub
