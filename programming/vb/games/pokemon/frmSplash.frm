VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2460
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   5070
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuffer 
      Height          =   285
      Left            =   225
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ListBox lstPlayers 
      Appearance      =   0  'Flat
      Columns         =   2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2385
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1155
      Width           =   2595
   End
   Begin VB.Label lblDistribution 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "License File Not Loaded"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1710
      Width           =   2100
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2460
      Left            =   0
      Top             =   0
      Width           =   5070
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
      TabIndex        =   9
      Top             =   2070
      Width           =   135
   End
   Begin VB.Label lblNewGame 
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
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
      Left            =   4065
      TabIndex        =   8
      Top             =   930
      Width           =   900
   End
   Begin VB.Label lblSG 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Game:"
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
      Left            =   2385
      TabIndex        =   7
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label lblLicenseDAT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOT LOADED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   3090
      TabIndex        =   5
      Top             =   540
      Width           =   1725
   End
   Begin VB.Label lblLicense 
      BackStyle       =   0  'Transparent
      Caption         =   "License:"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   540
      Width           =   720
   End
   Begin VB.Label lblDatDAT 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   3090
      TabIndex        =   3
      Top             =   345
      Width           =   1725
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   2400
      TabIndex        =   2
      Top             =   330
      Width           =   705
   End
   Begin VB.Label lblVerDAT 
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
      Left            =   3075
      TabIndex        =   1
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   705
   End
   Begin VB.Shape shpBgBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2220
      Left            =   120
      Top             =   120
      Width           =   2220
   End
   Begin VB.Image imgBg 
      Height          =   2160
      Left            =   150
      Picture         =   "frmSplash.frx":0000
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bOn As Boolean
Public strLicense As String, strAuthorize As String, strServer As String, strAbout As String, strCreator As String, strGenerator As String, clrColor As String
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Load()
    If Not GetSetting("Pokémon Adventure", "License", "ByPassLicense") = "AUTH9384" Then
        If FileExists(PathA + "\license") = True Then
            LoadText txtBuffer, PathA + "\license"
        Else
            MsgBox "License is missing or invalid!", vbExclamation, "License Error"
            End
        End If
        If txtBuffer.Text = Empty Then
            MsgBox "License is missing or invalid!", vbExclamation, "License Error"
            End
        Else
            strAuthorize = DecryptA(LineFromString(txtBuffer.Text, 3))
            If Not InStr(strAuthorize, "839-AUTHENTIC-") <> 0 Then
                MsgBox "License is missing or invalid!", vbExclamation, "License Error"
                End
            Else
                strLicense = DecryptA(LineFromString(txtBuffer.Text, 1))
                strServer = DecryptA(LineFromString(txtBuffer.Text, 2))
                strAuthorize = DecryptA(LineFromString(txtBuffer.Text, 3))
                strExpire = DecryptA(LineFromString(txtBuffer.Text, 4))
                strGenerator = DecryptA(LineFromString(txtBuffer.Text, 5))
                strCreator = DecryptA(LineFromString(txtBuffer.Text, 6))
                strAbout = LineFromString(txtBuffer.Text, 7)
                strAbout = ReplaceString(strAbout, "%release%", strLicense)
                strAbout = ReplaceString(strAbout, "%gametitle%", "Pokémon Adventure")
                strAbout = ReplaceString(strAbout, "%authcode%", strAuthorize)
                strAbout = ReplaceString(strAbout, "%creator%", strCreator)
                strAbout = ReplaceString(strAbout, "%generator%", strGenerator)
                clrColor = DecryptA(LineFromString(txtBuffer.Text, 8))
            End If
        End If
        strDate$ = Date
        strDate$ = ReplaceString(strDate$, "/", "")
        If Not strExpire = "Never" Then
            If strExpire <= strDate Then
                MsgBox strLicense & " has expired!", vbExclamation, "License Error"
                End
            End If
        End If
        lblDistribution.Caption = strLicense
        lblDistribution.ForeColor = clrColor
        lblLicenseDAT.Caption = strGenerator
    End If
    LoadGames lstPlayers
    lblVerDAT.Caption = App.Major & "." & App.Minor & "." & App.Revision
    lblDatDAT.Caption = Date
End Sub
Private Sub imgBg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblDistribution_Click()
    MsgBoxA Me, strAbout
End Sub
Private Sub lblExit_Click()
    End
End Sub
Private Sub lblNewGame_Click()
    Me.Enabled = False
    frmNewGame.Show
End Sub
Private Sub lstPlayers_DblClick()
    frmMain.Player = LCase(TrimSpaces(lstPlayers.Text))
    frmMain.Rival = GetFromINI("001", "001.2", PathA & "\" & LCase(TrimSpaces(lstPlayers.Text)))
    frmMain.Location = GetFromINI("001", "001.3", PathA & "\" & LCase(TrimSpaces(lstPlayers.Text)))
    frmMain.Version = GetFromINI("001", "001.4", PathA & "\" & LCase(TrimSpaces(lstPlayers.Text)))
    If GetSetting("Pokémon Adventure", "Introduction", lstPlayers.Text) = "Viewed" Then
        Me.Hide
        frmMain.Show
    Else
        frmIntro.txtIntro.Text = ReplaceString(frmIntro.txtIntro.Text, "-PLAYER-", lstPlayers.Text)
        Me.Hide
        frmIntro.Show
        SaveSetting "Pokémon Adventure", "Introduction", lstPlayers.Text, "Viewed"
    End If
End Sub
Private Sub lstPlayers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu frmPopup.menua
    End If
End Sub
