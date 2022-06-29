VERSION 5.00
Begin VB.Form frmTravel 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label imgAsh 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cursor"
      DragIcon        =   "frmTravel.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   30
      TabIndex        =   49
      Top             =   3660
      Width           =   645
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   30
      TabIndex        =   45
      Top             =   30
      Width           =   5250
   End
   Begin VB.Shape shpTop 
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   15
      Top             =   15
      Width           =   5280
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
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4710
      TabIndex        =   48
      Top             =   270
      Width           =   585
   End
   Begin VB.Label lblTravel 
      BackStyle       =   0  'Transparent
      Caption         =   "Travel"
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
      Left            =   4020
      TabIndex        =   47
      Top             =   270
      Width           =   585
   End
   Begin VB.Shape shypBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   3915
      Left            =   0
      Top             =   0
      Width           =   5325
   End
   Begin VB.Label Route18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   1710
      TabIndex        =   46
      Top             =   2640
      Width           =   690
   End
   Begin VB.Label PowerPlant 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4335
      TabIndex        =   44
      Top             =   1080
      Width           =   225
   End
   Begin VB.Label Route1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   990
      TabIndex        =   43
      Top             =   2490
      Width           =   225
   End
   Begin VB.Label Route2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   990
      TabIndex        =   42
      Top             =   1860
      Width           =   225
   End
   Begin VB.Label ViridianForest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   930
      TabIndex        =   41
      Top             =   1185
      Width           =   345
   End
   Begin VB.Label DiglettsCave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1275
      TabIndex        =   40
      Top             =   1185
      Width           =   180
   End
   Begin VB.Label PewterCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   990
      TabIndex        =   39
      Top             =   825
      Width           =   555
   End
   Begin VB.Label Route3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1545
      TabIndex        =   38
      Top             =   780
      Width           =   285
   End
   Begin VB.Label MtMoon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1815
      TabIndex        =   37
      Top             =   525
      Width           =   795
   End
   Begin VB.Label Route4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2220
      TabIndex        =   36
      Top             =   780
      Width           =   510
   End
   Begin VB.Label SeaCottage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   3510
      TabIndex        =   35
      Top             =   360
      Width           =   225
   End
   Begin VB.Label Route25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3225
      TabIndex        =   34
      Top             =   510
      Width           =   330
   End
   Begin VB.Label Route24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2985
      TabIndex        =   33
      Top             =   570
      Width           =   225
   End
   Begin VB.Label CeruleanCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2730
      TabIndex        =   32
      Top             =   765
      Width           =   645
   End
   Begin VB.Label Route5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3000
      TabIndex        =   31
      Top             =   1035
      Width           =   225
   End
   Begin VB.Label Route6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2970
      TabIndex        =   30
      Top             =   1785
      Width           =   225
   End
   Begin VB.Label VermilionCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2895
      TabIndex        =   29
      Top             =   1980
      Width           =   555
   End
   Begin VB.Label SSAnne 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2625
      TabIndex        =   28
      Top             =   2295
      Width           =   345
   End
   Begin VB.Label RockTunnel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3960
      TabIndex        =   27
      Top             =   750
      Width           =   360
   End
   Begin VB.Label Route10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   4110
      TabIndex        =   26
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label LavenderTown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4305
      TabIndex        =   25
      Top             =   1455
      Width           =   165
   End
   Begin VB.Label PokemonTower 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4185
      TabIndex        =   24
      Top             =   1290
      Width           =   120
   End
   Begin VB.Label Route8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3525
      TabIndex        =   23
      Top             =   1470
      Width           =   660
   End
   Begin VB.Label Route7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2550
      TabIndex        =   22
      Top             =   1455
      Width           =   225
   End
   Begin VB.Label CeladonCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1950
      TabIndex        =   21
      Top             =   1245
      Width           =   600
   End
   Begin VB.Label SaffronCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   2775
      TabIndex        =   20
      Top             =   1215
      Width           =   750
   End
   Begin VB.Label Route11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   19
      Top             =   2055
      Width           =   345
   End
   Begin VB.Label Route12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4305
      TabIndex        =   18
      Top             =   1935
      Width           =   150
   End
   Begin VB.Label Route13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3990
      TabIndex        =   17
      Top             =   2355
      Width           =   465
   End
   Begin VB.Label Route14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3765
      TabIndex        =   16
      Top             =   2565
      Width           =   225
   End
   Begin VB.Label Route15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3240
      TabIndex        =   15
      Top             =   3075
      Width           =   660
   End
   Begin VB.Label Route16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   1755
      TabIndex        =   14
      Top             =   1455
      Width           =   195
   End
   Begin VB.Label Route17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   1590
      TabIndex        =   13
      Top             =   1500
      Width           =   165
   End
   Begin VB.Label FushciaCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2400
      TabIndex        =   12
      Top             =   3075
      Width           =   840
   End
   Begin VB.Label SafariZone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   2790
      Width           =   840
   End
   Begin VB.Label SeaRoute19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2325
      TabIndex        =   10
      Top             =   3420
      Width           =   195
   End
   Begin VB.Label SeaRoute20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1755
      TabIndex        =   9
      Top             =   3420
      Width           =   195
   End
   Begin VB.Label SeafoamIslands 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1950
      TabIndex        =   8
      Top             =   3405
      Width           =   375
   End
   Begin VB.Label CinnabarIsland 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1095
      TabIndex        =   7
      Top             =   3255
      Width           =   660
   End
   Begin VB.Label SeaRoute21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   1080
      TabIndex        =   6
      Top             =   3150
      Width           =   225
   End
   Begin VB.Label Route22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   300
      TabIndex        =   5
      Top             =   1485
      Width           =   525
   End
   Begin VB.Label ViridianCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   750
      TabIndex        =   4
      Top             =   1995
      Width           =   540
   End
   Begin VB.Label PalletTown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   870
      TabIndex        =   3
      Top             =   2670
      Width           =   450
   End
   Begin VB.Label Route23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   375
      TabIndex        =   2
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label VictoryRoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   315
      TabIndex        =   1
      Top             =   915
      Width           =   150
   End
   Begin VB.Label IndigoPlateau 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   720
      Width           =   225
   End
   Begin VB.Image imgMap 
      Height          =   3945
      Left            =   30
      Picture         =   "frmTravel.frx":030A
      Top             =   90
      Width           =   5250
   End
End
Attribute VB_Name = "frmTravel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CeladonCity_Click()
    lblPlace.Caption = "Celadon City"
End Sub
Private Sub CeladonCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Celadon City"
End Sub
Private Sub CeruleanCity_Click()
    lblPlace.Caption = "Cerulean City"
End Sub
Private Sub CeruleanCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Cerulean City"
End Sub
Private Sub CinnabarIsland_Click()
    lblPlace.Caption = "Cinnabar Island"
End Sub
Private Sub CinnabarIsland_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Cinnabar Island"
End Sub
Private Sub DiglettsCave_Click()
    lblPlace.Caption = "Diglett's Cave"
End Sub
Private Sub DiglettsCave_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Diglett's Cave"
End Sub
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.lblLocDAT.Caption = GetLocation
    frmMain.Show
End Sub
Private Sub FushciaCity_Click()
    lblPlace.Caption = "Fushcia City"
End Sub
Private Sub FushciaCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Fushcia City"
End Sub
Private Sub imgAsh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgAsh.Drag
End Sub
Private Sub IndigoPlateau_Click()
    lblPlace.Caption = "Indigo Plateau"
End Sub
Private Sub IndigoPlateau_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Indigo Plateau"
End Sub
Private Sub LavenderTown_Click()
    lblPlace.Caption = "Lavender Town"
End Sub
Private Sub LavenderTown_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Lavender Town"
End Sub
Private Sub lblCancel_Click()
    Unload Me
End Sub
Private Sub lblPlace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblTravel_Click()
    If lblPlace.Caption = frmMain.lblLocDAT.Caption Then
        strBuffer$ = "You are already in " + frmMain.lblLocDAT.Caption + "!"
        MsgBoxA Me, strBuffer$
    Else
        var1$ = lblPlace.Caption
        lblPlace.Caption = "Preparing To Travel..."
        Me.Refresh
        TimeOut 2
        lblPlace.Caption = "Traveling..."
        Me.Refresh
        TimeOut 5
        lblPlace.Caption = "Destination Reached!"
        Me.Refresh
        TimeOut 1
        SetLocation var1$
        Unload Me
    End If
End Sub
Private Sub MtMoon_Click()
    lblPlace.Caption = "Mt. Moon"
End Sub
Private Sub MtMoon_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Mt. Moon"
End Sub
Private Sub PalletTown_Click()
    lblPlace.Caption = "Pallet Town"
End Sub
Private Sub PalletTown_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Pallet Town"
End Sub
Private Sub PewterCity_Click()
    lblPlace.Caption = "Pewter City"
End Sub
Private Sub PewterCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Pewter City"
End Sub
Private Sub PokemonTower_Click()
    lblPlace.Caption = "Pokémon Tower"
End Sub
Private Sub PokemonTower_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Pokémon Tower"
End Sub
Private Sub PowerPlant_Click()
    lblPlace.Caption = "Power Plant"
End Sub
Private Sub PowerPlant_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Power Plant"
End Sub
Private Sub RockTunnel_Click()
    lblPlace.Caption = "Rock Tunnel"
End Sub
Private Sub RockTunnel_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Rock Tunnel"
End Sub
Private Sub Route1_Click()
    lblPlace.Caption = "Route 1"
End Sub
Private Sub Route1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 1"
End Sub
Private Sub Route10_Click()
    lblPlace.Caption = "Route 10"
End Sub
Private Sub Route10_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 10"
End Sub
Private Sub Route11_Click()
    lblPlace.Caption = "Route 11"
End Sub
Private Sub Route11_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 11"
End Sub
Private Sub Route12_Click()
    lblPlace.Caption = "Route 12"
End Sub
Private Sub Route12_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 12"
End Sub
Private Sub Route13_Click()
    lblPlace.Caption = "Route 13"
End Sub
Private Sub Route13_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 13"
End Sub
Private Sub Route14_Click()
    lblPlace.Caption = "Route 14"
End Sub
Private Sub Route14_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 14"
End Sub
Private Sub Route15_Click()
    lblPlace.Caption = "Route 15"
End Sub
Private Sub Route15_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 15"
End Sub
Private Sub Route16_Click()
    lblPlace.Caption = "Route 16"
End Sub
Private Sub Route16_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 16"
End Sub
Private Sub Route17_Click()
    lblPlace.Caption = "Route 17"
End Sub
Private Sub Route17_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 17"
End Sub
Private Sub Route18_Click()
    lblPlace.Caption = "Route 18"
End Sub
Private Sub Route18_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 18"
End Sub
Private Sub Route2_Click()
    lblPlace.Caption = "Route 2"
End Sub
Private Sub Route2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 2"
End Sub
Private Sub Route22_Click()
    lblPlace.Caption = "Route 22"
End Sub
Private Sub Route22_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 22"
End Sub
Private Sub Route23_Click()
    lblPlace.Caption = "Route 23"
End Sub
Private Sub Route23_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 23"
End Sub
Private Sub Route24_Click()
    lblPlace.Caption = "Route 24"
End Sub
Private Sub Route24_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 24"
End Sub
Private Sub Route25_Click()
    lblPlace.Caption = "Route 25"
End Sub
Private Sub Route25_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 25"
End Sub
Private Sub Route3_Click()
    lblPlace.Caption = "Route 3"
End Sub
Private Sub Route3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 3"
End Sub
Private Sub Route4_Click()
    lblPlace.Caption = "Route 4"
End Sub
Private Sub Route4_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 4"
End Sub
Private Sub Route5_Click()
    lblPlace.Caption = "Route 5"
End Sub
Private Sub Route5_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 5"
End Sub
Private Sub Route6_Click()
    lblPlace.Caption = "Route 6"
End Sub
Private Sub Route6_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 6"
End Sub
Private Sub Route7_Click()
    lblPlace.Caption = "Route 7"
End Sub
Private Sub Route7_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 7"
End Sub
Private Sub Route8_Click()
    lblPlace.Caption = "Route 8"
End Sub
Private Sub Route8_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Route 8"
End Sub
Private Sub SafariZone_Click()
    lblPlace.Caption = "Safari Zone"
End Sub
Private Sub SafariZone_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Safari Zone"
End Sub
Private Sub SaffronCity_Click()
    lblPlace.Caption = "Saffron City"
End Sub
Private Sub SaffronCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Saffron City"
End Sub
Private Sub SeaCottage_Click()
    lblPlace.Caption = "Sea Cottage"
End Sub
Private Sub SeaCottage_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Sea Cottage"
End Sub
Private Sub SeafoamIslands_Click()
    lblPlace.Caption = "Seafoam Islands"
End Sub
Private Sub SeafoamIslands_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Seafoam Islands"
End Sub
Private Sub SeaRoute19_Click()
    lblPlace.Caption = "Sea Route 19"
End Sub
Private Sub SeaRoute19_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Sea Route 19"
End Sub
Private Sub SeaRoute20_Click()
    lblPlace.Caption = "Sea Route 20"
End Sub
Private Sub SeaRoute20_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Sea Route 20"
End Sub
Private Sub SeaRoute21_Click()
    lblPlace.Caption = "Sea Route 21"
End Sub
Private Sub SeaRoute21_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Sea Route 21"
End Sub
Private Sub SSAnne_Click()
    lblPlace.Caption = "S.S. Anne"
End Sub
Private Sub SSAnne_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "S.S. Anne"
End Sub
Private Sub VermilionCity_Click()
    lblPlace.Caption = "Vermilion City"
End Sub
Private Sub VermilionCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Vermilion City"
End Sub
Private Sub VictoryRoad_Click()
    lblPlace.Caption = "Victory Road"
End Sub
Private Sub VictoryRoad_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Victory Road"
End Sub
Private Sub ViridianCity_Click()
    lblPlace.Caption = "Viridian City"
End Sub
Private Sub ViridianCity_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Viridian City"
End Sub
Private Sub ViridianForest_Click()
    lblPlace.Caption = "Viridian Forest"
End Sub
Private Sub ViridianForest_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lblPlace.Caption = "Viridian Forest"
End Sub
