VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Options 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keno Settings"
   ClientHeight    =   5070
   ClientLeft      =   2565
   ClientTop       =   1785
   ClientWidth     =   6645
   Icon            =   "Options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame statsframe 
      BackColor       =   &H00000000&
      Caption         =   "Keno Stats"
      ForeColor       =   &H000000FF&
      Height          =   2670
      Left            =   1260
      TabIndex        =   18
      Top             =   435
      Visible         =   0   'False
      Width           =   5025
      Begin VB.Label credits_won 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3870
         TabIndex        =   26
         Top             =   1455
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Credits won this game.="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   -30
         TabIndex        =   25
         Top             =   1410
         Width           =   3840
      End
      Begin VB.Label deals_label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3870
         TabIndex        =   24
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deals.="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   -30
         TabIndex        =   23
         Top             =   1995
         Width           =   3840
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Your Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   435
         TabIndex        =   22
         Top             =   315
         Width           =   1980
      End
      Begin VB.Label payrate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   360
         Left            =   3120
         TabIndex        =   21
         Top             =   330
         Width           =   135
      End
      Begin VB.Label bet_total 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3870
         TabIndex        =   20
         Top             =   870
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bet this game.="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   195
         TabIndex        =   19
         Top             =   825
         Width           =   3615
      End
   End
   Begin VB.CommandButton moremoney 
      BackColor       =   &H00008000&
      Caption         =   "Get 100 Credits"
      Height          =   375
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton pay 
      Caption         =   "4 BET"
      Height          =   375
      Index           =   3
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2025
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton pay 
      Caption         =   "3 BET"
      Height          =   375
      Index           =   2
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1545
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton pay 
      Caption         =   "2 BET"
      Height          =   375
      Index           =   1
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1065
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton pay 
      Caption         =   "1 BET"
      Height          =   375
      Index           =   0
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   585
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox paytableimage 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2910
      Left            =   1080
      Picture         =   "Options.frx":000C
      ScaleHeight     =   2850
      ScaleWidth      =   5400
      TabIndex        =   11
      Top             =   285
      Visible         =   0   'False
      Width           =   5460
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   195
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   210
      LargeChange     =   50
      Left            =   1815
      Max             =   1
      Min             =   500
      TabIndex        =   8
      Top             =   3960
      Value           =   1
      Width           =   3975
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label pay_label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keno Paytable"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   -60
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      TabIndex        =   10
      Top             =   4185
      Width           =   5130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   3765
      Width           =   5130
   End
   Begin VB.Menu gamemenu 
      Caption         =   "&Game Menu"
      Begin VB.Menu lastgame 
         Caption         =   "L&ast Game"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu newgame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu loadgame 
         Caption         =   "&Load Game"
      End
      Begin VB.Menu savegame 
         Caption         =   "&Save Game"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu printer 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu automenu 
         Caption         =   "A&uto Play Menu"
         Enabled         =   0   'False
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu sound 
         Caption         =   "&Sound"
         Begin VB.Menu soundon 
            Caption         =   "&On"
            Checked         =   -1  'True
         End
         Begin VB.Menu soundoff 
            Caption         =   "O&ff"
         End
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu showpaytable 
         Caption         =   "Show &Paytable"
      End
      Begin VB.Menu stats 
         Caption         =   "Show S&tats"
         Enabled         =   0   'False
      End
      Begin VB.Menu line10 
         Caption         =   "-"
      End
      Begin VB.Menu fullscreenmode 
         Caption         =   "&Full Screen Mode"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu aboutkeno 
      Caption         =   "&About"
      Begin VB.Menu aboutform 
         Caption         =   "Keno Info"
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Const CCDEVICENAME = 32
    Const CCFORMNAME = 32
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

    Dim DevM As DEVMODE


Public secret As Integer
Public crypt As Integer
Public keypressed As Integer
Public coinsout As Double
Private Sub aboutform_Click()
about.Show
End Sub
Private Sub automenu_Click()
autoplayform.Show
End Sub
Private Sub cmdCancel_Click()
   If board.Visible = True Then
    Me.Hide
    Else
    Unload Me
    End If
    End Sub
Public Sub cmdOK_Click()
If dollars = 0 Then
    MsgBox "Please choose Load Game or New Game from the Game Menu to receive credits.", vbOKOnly, "Keno Settings!"
    Exit Sub
    End If
    delaytime = HScroll1.Value
    board.Show
    cmdOK.Enabled = False
    Me.Hide
    moremoney.Visible = False
    If giveme = 1 Then
    boxes_checked = 0
        getgame = 3
        cmdOK.Enabled = True
        giveme = 0
        Unload board
        board.Show
        End If
End Sub
Private Sub exit_Click()
If MsgBox("Are You Sure You Want To Quit?", vbYesNo, "Vegas Video Keno!") = vbYes Then
 If board.Visible = True Then
 'save sound on or off variable in registry
  SaveSetting App.Title, "Options", "sound", soundonoff
  'save delaytime variable value in registry
  SaveSetting App.Title, "options", "delay", delaytime
  Call savechecks
  SaveSetting App.Title, "options", "previousgame", getgame
SaveSetting App.Title, "options", "lastcredits", dollars
SaveSetting App.Title, "options", "coinsout", hopperempty
SaveSetting App.Title, "options", "totalbet", bettotal
SaveSetting App.Title, "options", "hopper2", hopper2
SaveSetting App.Title, "options", "totaldeals", totaldeals
SaveSetting App.Title, "options", "fullmode", fullmode
End If

If fullmode = 1 Then
Call ChangeRes(normalwidth, normalheight)
End If
   End
 Else
 Exit Sub
 End If
End Sub

Private Sub Form_Activate()
If dollars > 20 Then
automenu.Enabled = True
Else
automenu.Enabled = False
End If
If hopperempty > 0 And bettotal > 0 Or previousgame = 1 Then
stats.Enabled = True
Else
stats.Enabled = False
End If
If board.Visible = True Or stats.Enabled = True Then
printer.Enabled = True
Else
printer.Enabled = False
End If
If board.Visible = True Then
    fullscreenmode.Enabled = True
    End If
    
    If board.Visible = False And fullmode = 0 Then
        fullscreenmode.Enabled = False
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call secrets(KeyCode)
If keypressed = 524 Then
MsgBox "You Found A Secret", vbOKOnly, "VVKeno Secrets"
dollars = dollars + 100000
board.credits_label.Caption = dollars
board.credits_label.Refresh
boxes_checked = 0
        getgame = 3
        cmdOK.Enabled = True
        Unload board
        board.Show
End If
End Sub
Public Function secrets(keyed As Integer)
keypressed = keypressed + keyed
End Function
Public Sub form_load()
If dollars < 20 Then
automenu.Enabled = False
End If

Me.Picture = LoadResPicture("optionback", bitmap)
' Registry previous game
lastcredits = GetSetting(App.Title, "options", "lastcredits", 0)
previousgame = GetSetting(App.Title, "options", "previousgame", 0)
If previousgame = 0 Then
    lastgame.Enabled = False
    End If
Call initboard
cmdOK.Enabled = False
End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If MsgBox("Are You Sure You Want To Quit?", vbYesNo, "Vegas Video Keno!") = vbYes Then
  If board.Visible = True Then
  'save sound on or off variable in registry
  SaveSetting App.Title, "Options", "sound", soundonoff
  'save delaytime variable value in registry
  SaveSetting App.Title, "options", "delay", delaytime
  Call savechecks
  SaveSetting App.Title, "options", "previousgame", getgame
   SaveSetting App.Title, "options", "lastcredits", dollars
   SaveSetting App.Title, "options", "coinsout", hopperempty
   SaveSetting App.Title, "options", "totalbet", bettotal
   SaveSetting App.Title, "options", "hopper2", hopper2
   SaveSetting App.Title, "options", "totaldeals", totaldeals
   SaveSetting App.Title, "options", "fullmode", fullmode
   End If
   
If fullmode = 1 Then
Call ChangeRes(normalwidth, normalheight)
End If
   End
 Else
 cancel = True
  End If
End Sub

Private Sub fullscreenmode_Click()
If fullscreenmode.Checked = False Then
        fullmode = 1
        Call ChangeRes(640, 480)
        fullscreenmode.Checked = True
        Me.Hide
        Me.Show
        cmdOK.Enabled = True
    Else
        fullmode = 0
        Call ChangeRes(normalwidth, normalheight)
        fullscreenmode.Checked = False
       If board.Visible = True Then
        board.Hide
        board.Show
        End If
        Me.Hide
        Me.Show
        cmdOK.Enabled = True
    End If

End Sub

Private Sub hScroll1_Change()
    If HScroll1.Value >= 90 Then
    Label1.Caption = "I'm waiting for a drink speed!"
    Label3.Caption = "Super Turtle Slow"
ElseIf HScroll1.Value >= 65 And HScroll1.Value < 90 Then
    Label1.Caption = "Feeling pretty good just had a drink Speed!"
    Label3.Caption = "Slow-Medium"
ElseIf HScroll1.Value >= 30 And HScroll1.Value < 65 Then
    Label1.Caption = "I Just got a free roll of nickels speed!"
    Label3.Caption = "Medium"
    ElseIf HScroll1.Value >= 10 And HScroll1.Value < 30 Then
    Label1.Caption = "I Just hit for 600 credits. Bring it on!"
    Label3.Caption = "Medium-Fast"
    ElseIf HScroll1.Value <= 10 Then
    Label1.Caption = "Hurry UP!,  We're late for our flight!"
    Label3.Caption = "I think the above says it all!"
    End If
Label1.Refresh
delaytime = HScroll1.Value
cmdOK.Enabled = True
End Sub

Private Sub lastgame_Click()
Unload board
getgame = 2
boxes_checked = 0
amountbet = 0
Call getchecks
Me.Hide
board.Show
cmdOK.Enabled = True
End Sub
Private Sub loadgame_Click()
secret = 12589
crypt = 258
Dim filename As String
On Error GoTo ErrHandler
dialog.Filter = "Text (*.ksg)|*.ksg"
dialog.InitDir = ""
dialog.ShowOpen
filename = dialog.filename
Open filename For Input As #1
Input #1, dollars, key, key2, c1_checked, c2_checked, c3_checked, c4_checked, c5_checked, c6_checked, c7_checked, c8_checked, _
 c9_checked, c10_checked, c11_checked, c12_checked, c13_checked, c14_checked, c15_checked, c16_checked, c17_checked, c18_checked, _
 c19_checked, c20_checked, c21_checked, c22_checked, c23_checked, c24_checked, c25_checked, c26_checked, c27_checked, c28_checked, c29_checked, c30_checked, c31_checked, _
 c32_checked, c33_checked, c34_checked, c35_checked, c36_checked, c37_checked, c38_checked, c39_checked, c40_checked, c41_checked, c42_checked, c43_checked, c44_checked, c45_checked, _
 c46_checked, c47_checked, c48_checked, c49_checked, c50_checked, c51_checked, c52_checked, c53_checked, c54_checked, c55_checked, _
 c56_checked, c57_checked, c58_checked, c59_checked, c60_checked, c61_checked, c62_checked, c63_checked, c64_checked, c65_checked, c66_checked, c67_checked, c68_checked, _
 c69_checked, c70_checked, c71_checked, c72_checked, c73_checked, c74_checked, c75_checked, c76_checked, c77_checked, c78_checked, c79_checked, c80_checked, _
 bettotal, hopperempty, hopper2, totaldeals
 Close #1
 
If secret = (key / dollars) / crypt And crypt = (key2 - dollars) / key Then
    getgame = 3
    Unload board
    amountbet = 0
    boxes_checked = 0
    cmdOK.Enabled = True
    board.Show
    Me.Hide
    Exit Sub
    Else
   MsgBox "This is not a valid saved game", vbOKOnly
    dollars = 0
    Exit Sub
    End If
ErrHandler:
Exit Sub
cmdOK.Enabled = True
     End Sub
Private Sub moremoney_Click()
dollars = dollars + 100
cmdOK.Enabled = True
giveme = 1
If dollars >= 200 Then
moremoney.Visible = False
Exit Sub
End If

End Sub
Private Sub newgame_Click()
Unload board
getgame = 1
Me.Hide
newgamedollars.Show
amountbet = 0
bettotal = 0
hopperempty = 0
hopper2 = 0
totaldeals = 0
End Sub

Private Sub printer_Click()
print2form.Show

End Sub

Private Sub savegame_Click()
cmdOK.Enabled = True
secret = 12589
crypt = 258
key = (dollars * secret) * crypt
key2 = (key * crypt) + dollars
Dim filename As String
On Error GoTo ErrHandler
dialog.Filter = "Text (*.ksg)|*.ksg"
dialog.InitDir = ""
dialog.Flags = cdlOFNOverwritePrompt
dialog.ShowSave
filename = dialog.filename
Open filename For Output As #1
Write #1, dollars, key, key2, c1_checked, c2_checked, c3_checked, c4_checked, c5_checked, c6_checked, c7_checked, c8_checked, _
 c9_checked, c10_checked, c11_checked, c12_checked, c13_checked, c14_checked, c15_checked, c16_checked, c17_checked, c18_checked, _
 c19_checked, c20_checked, c21_checked, c22_checked, c23_checked, c24_checked, c25_checked, c26_checked, c27_checked, c28_checked, c29_checked, c30_checked, c31_checked, _
 c32_checked, c33_checked, c34_checked, c35_checked, c36_checked, c37_checked, c38_checked, c39_checked, c40_checked, c41_checked, c42_checked, c43_checked, c44_checked, c45_checked, _
 c46_checked, c47_checked, c48_checked, c49_checked, c50_checked, c51_checked, c52_checked, c53_checked, c54_checked, c55_checked, _
 c56_checked, c57_checked, c58_checked, c59_checked, c60_checked, c61_checked, c62_checked, c63_checked, c64_checked, c65_checked, c66_checked, c67_checked, c68_checked, _
 c69_checked, c70_checked, c71_checked, c72_checked, c73_checked, c74_checked, c75_checked, c76_checked, c77_checked, c78_checked, c79_checked, c80_checked, _
 bettotal, hopperempty, hopper2, totaldeals
Close #1
ErrHandler:
Exit Sub
End Sub

Private Function paytablesound()
If soundonoff = 0 Then
Dim x%
 soundName$ = "deal.wav" ' The file to play
 wFlags% = SND_ASYNC Or SND_NODEFAULT
 x% = sndPlaySound(soundName$, uFlags%)
 Else
 Exit Function
 End If
 
End Function
Private Sub pay_click(index As Integer)
Select Case index
    Case 0
    Call paytablesound
         'paytableimage.Picture = LoadPicture
         pay_label.Caption = "Pay with 1 Bet."
        paytableimage.Picture = LoadResPicture("pay1", bitmap)
    Case 1
    Call paytablesound
       '  paytableimage.Picture = LoadPicture
   pay_label.Caption = "Pay with 2 Bet."
        paytableimage.Picture = LoadResPicture("pay2", bitmap)
    Case 2
    Call paytablesound
       'paytableimage.Picture = LoadPicture
    pay_label.Caption = "Pay with 3 Bet."
        paytableimage.Picture = LoadResPicture("pay3", bitmap)
    Case 3
    Call paytablesound
        'paytableimage.Picture = LoadPicture
    pay_label.Caption = "Pay with 4 Bet."
        paytableimage.Picture = LoadResPicture("pay4", bitmap)
    End Select
End Sub

Private Sub showpaytable_Click()
If stats.Checked = True Then
statsframe.Visible = False
stats.Checked = False
End If
If showpaytable.Checked = True Then
Call hidedisplaytable
showpaytable.Checked = False
Else
Call displaypaytable
showpaytable.Checked = True
End If
cmdOK.Enabled = True
End Sub
Private Function hidedisplaytable()
paytableimage.Visible = False
pay(0).Visible = False
pay(1).Visible = False
pay(2).Visible = False
pay(3).Visible = False
pay_label.Visible = False
End Function

Private Function displaypaytable()
statsframe.Visible = False
paytableimage.Visible = True
pay(0).Visible = True
pay(1).Visible = True
pay(2).Visible = True
pay(3).Visible = True
pay_label.Visible = True

End Function

Private Sub soundoff_Click()
soundonoff = 1
soundoff.Checked = True
soundon.Checked = False
cmdOK.Enabled = True
keypressed = 0
End Sub
Private Sub soundon_Click()
soundonoff = 0
soundoff.Checked = False
soundon.Checked = True
cmdOK.Enabled = True
End Sub
Private Function initboard()
' Registry sound info calls
 soundonoff = GetSetting(App.Title, "Options", "sound", 0)
        If soundonoff = 0 Then
            soundon.Checked = True
            soundoff.Checked = False
            soundonoff = 0
        ElseIf soundonoff = 1 Then
            soundon.Checked = False
            soundoff.Checked = True
            soundonoff = 1
            End If
'Registry Delay Calls
delaytime = GetSetting(App.Title, "options", "delay", 0)
    If delaytime = 0 Then
        delaytime = 50
        End If
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
HScroll1.Value = delaytime
HScroll1.Refresh
 If HScroll1.Value >= 90 Then
    Label1.Caption = "I'm waiting for a drink speed!"
    Label3.Caption = "Super Turtle Slow"
ElseIf HScroll1.Value >= 65 And HScroll1.Value < 90 Then
    Label1.Caption = "Feeling pretty good just had a drink Speed!"
    Label3.Caption = "Slow-Medium"
ElseIf HScroll1.Value >= 30 And HScroll1.Value < 65 Then
    Label1.Caption = "I Just got a free roll of nickels speed!"
    Label3.Caption = "Medium"
    ElseIf HScroll1.Value >= 10 And HScroll1.Value < 30 Then
    Label1.Caption = "I Just hit for 600 credits. Bring it on!"
    Label3.Caption = "Medium-Fast"
    ElseIf HScroll1.Value <= 10 Then
    Label1.Caption = "Hurry UP!,  We're late for our flight!"
    Label3.Caption = "I think the above says it all!"
    End If
Label1.Refresh
End Function
Private Function savechecks()
SaveSetting App.Title, "Options", "c1", c1_checked
SaveSetting App.Title, "Options", "c2", c2_checked
SaveSetting App.Title, "Options", "c3", c3_checked
SaveSetting App.Title, "Options", "c4", c4_checked
SaveSetting App.Title, "Options", "c5", c5_checked
SaveSetting App.Title, "Options", "c6", c6_checked
SaveSetting App.Title, "Options", "c7", c7_checked
SaveSetting App.Title, "Options", "c8", c8_checked
SaveSetting App.Title, "Options", "c9", c9_checked
SaveSetting App.Title, "Options", "c10", c10_checked
SaveSetting App.Title, "Options", "c11", c11_checked
SaveSetting App.Title, "Options", "c12", c12_checked
SaveSetting App.Title, "Options", "c13", c13_checked
SaveSetting App.Title, "Options", "c14", c14_checked
SaveSetting App.Title, "Options", "c15", c15_checked
SaveSetting App.Title, "Options", "c16", c16_checked
SaveSetting App.Title, "Options", "c17", c17_checked
SaveSetting App.Title, "Options", "c18", c18_checked
SaveSetting App.Title, "Options", "c19", c19_checked
SaveSetting App.Title, "Options", "c20", c20_checked
SaveSetting App.Title, "Options", "c21", c21_checked
SaveSetting App.Title, "Options", "c22", c22_checked
SaveSetting App.Title, "Options", "c23", c23_checked
SaveSetting App.Title, "Options", "c24", c24_checked
SaveSetting App.Title, "Options", "c25", c25_checked
SaveSetting App.Title, "Options", "c26", c26_checked
SaveSetting App.Title, "Options", "c27", c27_checked
SaveSetting App.Title, "Options", "c28", c28_checked
SaveSetting App.Title, "Options", "c29", c29_checked
SaveSetting App.Title, "Options", "c30", c30_checked
SaveSetting App.Title, "Options", "c31", c31_checked
SaveSetting App.Title, "Options", "c32", c32_checked
SaveSetting App.Title, "Options", "c33", c33_checked
SaveSetting App.Title, "Options", "c34", c34_checked
SaveSetting App.Title, "Options", "c35", c35_checked
SaveSetting App.Title, "Options", "c36", c36_checked
SaveSetting App.Title, "Options", "c37", c37_checked
SaveSetting App.Title, "Options", "c38", c38_checked
SaveSetting App.Title, "Options", "c39", c39_checked
SaveSetting App.Title, "Options", "c40", c40_checked
SaveSetting App.Title, "Options", "c41", c41_checked
SaveSetting App.Title, "Options", "c42", c42_checked
SaveSetting App.Title, "Options", "c43", c43_checked
SaveSetting App.Title, "Options", "c44", c44_checked
SaveSetting App.Title, "Options", "c45", c45_checked
SaveSetting App.Title, "Options", "c46", c46_checked
 SaveSetting App.Title, "Options", "c47", c47_checked
 SaveSetting App.Title, "Options", "c48", c48_checked
 SaveSetting App.Title, "Options", "c49", c49_checked
 SaveSetting App.Title, "Options", "c50", c50_checked
SaveSetting App.Title, "Options", "c51", c51_checked
 SaveSetting App.Title, "Options", "c52", c52_checked
  SaveSetting App.Title, "Options", "c53", c53_checked
 SaveSetting App.Title, "Options", "c54", c54_checked
SaveSetting App.Title, "Options", "c55", c55_checked
  SaveSetting App.Title, "Options", "c56", c56_checked
 SaveSetting App.Title, "Options", "c57", c57_checked
SaveSetting App.Title, "Options", "c58", c58_checked
 SaveSetting App.Title, "Options", "c59", c59_checked
SaveSetting App.Title, "Options", "c60", c60_checked
SaveSetting App.Title, "Options", "c61", c61_checked
SaveSetting App.Title, "Options", "c62", c62_checked
SaveSetting App.Title, "Options", "c63", c63_checked
 SaveSetting App.Title, "Options", "c64", c64_checked
  SaveSetting App.Title, "Options", "c65", c65_checked
SaveSetting App.Title, "Options", "c66", c66_checked
SaveSetting App.Title, "Options", "c67", c67_checked
SaveSetting App.Title, "Options", "c68", c68_checked
SaveSetting App.Title, "Options", "c69", c69_checked
SaveSetting App.Title, "Options", "c70", c70_checked
SaveSetting App.Title, "Options", "c71", c71_checked
SaveSetting App.Title, "Options", "c72", c72_checked
SaveSetting App.Title, "Options", "c73", c73_checked
SaveSetting App.Title, "Options", "c74", c74_checked
SaveSetting App.Title, "Options", "c75", c75_checked
SaveSetting App.Title, "Options", "c76", c76_checked
SaveSetting App.Title, "Options", "c77", c77_checked
SaveSetting App.Title, "Options", "c78", c78_checked
SaveSetting App.Title, "Options", "c79", c79_checked
SaveSetting App.Title, "Options", "c80", c80_checked
End Function
Public Function getchecks()
bettotal = GetSetting(App.Title, "options", "totalbet", 0)
hopperempty = GetSetting(App.Title, "options", "coinsout", 0)
hopper2 = GetSetting(App.Title, "options", "hopper2", 0)
totaldeals = GetSetting(App.Title, "options", "totaldeals", 0)
c1_checked = GetSetting(App.Title, "Options", "c1", 0)
c2_checked = GetSetting(App.Title, "Options", "c2", 0)
c3_checked = GetSetting(App.Title, "Options", "c3", 0)
c4_checked = GetSetting(App.Title, "Options", "c4", 0)
c5_checked = GetSetting(App.Title, "Options", "c5", 0)
c6_checked = GetSetting(App.Title, "Options", "c6", 0)
c7_checked = GetSetting(App.Title, "Options", "c7", 0)
c8_checked = GetSetting(App.Title, "Options", "c8", 0)
c9_checked = GetSetting(App.Title, "Options", "c9", 0)
c10_checked = GetSetting(App.Title, "Options", "c10", 0)
c11_checked = GetSetting(App.Title, "Options", "c11", 0)
c12_checked = GetSetting(App.Title, "Options", "c12", 0)
c13_checked = GetSetting(App.Title, "Options", "c13", 0)
c14_checked = GetSetting(App.Title, "Options", "c14", 0)
c15_checked = GetSetting(App.Title, "Options", "c15", 0)
c16_checked = GetSetting(App.Title, "Options", "c16", 0)
c17_checked = GetSetting(App.Title, "Options", "c17", 0)
c18_checked = GetSetting(App.Title, "Options", "c18", 0)
c19_checked = GetSetting(App.Title, "Options", "c19", 0)
c20_checked = GetSetting(App.Title, "Options", "c20", 0)
c21_checked = GetSetting(App.Title, "Options", "c21", 0)
c22_checked = GetSetting(App.Title, "Options", "c22", 0)
c23_checked = GetSetting(App.Title, "Options", "c23", 0)
c24_checked = GetSetting(App.Title, "Options", "c24", 0)
c25_checked = GetSetting(App.Title, "Options", "c25", 0)
c26_checked = GetSetting(App.Title, "Options", "c26", 0)
c27_checked = GetSetting(App.Title, "Options", "c27", 0)
c28_checked = GetSetting(App.Title, "Options", "c28", 0)
c29_checked = GetSetting(App.Title, "Options", "c29", 0)
c30_checked = GetSetting(App.Title, "Options", "c30", 0)
c31_checked = GetSetting(App.Title, "Options", "c31", 0)
c32_checked = GetSetting(App.Title, "Options", "c32", 0)
c33_checked = GetSetting(App.Title, "Options", "c33", 0)
c34_checked = GetSetting(App.Title, "Options", "c34", 0)
c35_checked = GetSetting(App.Title, "Options", "c35", 0)
c36_checked = GetSetting(App.Title, "Options", "c36", 0)
c37_checked = GetSetting(App.Title, "Options", "c37", 0)
c38_checked = GetSetting(App.Title, "Options", "c38", 0)
c39_checked = GetSetting(App.Title, "Options", "c39", 0)
c40_checked = GetSetting(App.Title, "Options", "c40", 0)
c41_checked = GetSetting(App.Title, "Options", "c41", 0)
c42_checked = GetSetting(App.Title, "Options", "c42", 0)
c43_checked = GetSetting(App.Title, "Options", "c43", 0)
c44_checked = GetSetting(App.Title, "Options", "c44", 0)
c45_checked = GetSetting(App.Title, "Options", "c45", 0)
c46_checked = GetSetting(App.Title, "Options", "c46", 0)
c47_checked = GetSetting(App.Title, "Options", "c47", 0)
c48_checked = GetSetting(App.Title, "Options", "c48", 0)
c49_checked = GetSetting(App.Title, "Options", "c49", 0)
c50_checked = GetSetting(App.Title, "Options", "c50", 0)
c51_checked = GetSetting(App.Title, "Options", "c51", 0)
c52_checked = GetSetting(App.Title, "Options", "c52", 0)
c53_checked = GetSetting(App.Title, "Options", "c53", 0)
c54_checked = GetSetting(App.Title, "Options", "c54", 0)
c55_checked = GetSetting(App.Title, "Options", "c55", 0)
c56_checked = GetSetting(App.Title, "Options", "c56", 0)
c57_checked = GetSetting(App.Title, "Options", "c57", 0)
c58_checked = GetSetting(App.Title, "Options", "c58", 0)
c59_checked = GetSetting(App.Title, "Options", "c59", 0)
c60_checked = GetSetting(App.Title, "Options", "c60", 0)
c61_checked = GetSetting(App.Title, "Options", "c61", 0)
c62_checked = GetSetting(App.Title, "Options", "c62", 0)
c63_checked = GetSetting(App.Title, "Options", "c63", 0)
c64_checked = GetSetting(App.Title, "Options", "c64", 0)
c65_checked = GetSetting(App.Title, "Options", "c65", 0)
c66_checked = GetSetting(App.Title, "Options", "c66", 0)
c67_checked = GetSetting(App.Title, "Options", "c67", 0)
c68_checked = GetSetting(App.Title, "Options", "c68", 0)
c69_checked = GetSetting(App.Title, "Options", "c69", 0)
c70_checked = GetSetting(App.Title, "Options", "c70", 0)
c71_checked = GetSetting(App.Title, "Options", "c71", 0)
c72_checked = GetSetting(App.Title, "Options", "c72", 0)
c73_checked = GetSetting(App.Title, "Options", "c73", 0)
c74_checked = GetSetting(App.Title, "Options", "c74", 0)
c75_checked = GetSetting(App.Title, "Options", "c75", 0)
c76_checked = GetSetting(App.Title, "Options", "c76", 0)
c77_checked = GetSetting(App.Title, "Options", "c77", 0)
c78_checked = GetSetting(App.Title, "Options", "c78", 0)
c79_checked = GetSetting(App.Title, "Options", "c79", 0)
c80_checked = GetSetting(App.Title, "Options", "c80", 0)
End Function
Public Function returnrate()
If hopperempty = 0 Or bettotal = 0 Then
MsgBox "Return percentage not available yet.", vbApplicationModal
payrate.Caption = "N/A"
Exit Function
End If

Dim payreturn As Double

payreturn = hopperempty / bettotal
percentagerate = Format(payreturn, "percent")
payrate.Caption = percentagerate
End Function
Private Sub stats_Click()
If previousgame = 0 Then
MsgBox "Stats Not Available Yet" & Chr(13) & "Please Select New, Load, or Last Game", vbOKOnly
Exit Sub
End If
If showpaytable.Checked = True Then
Call hidedisplaytable
showpaytable.Checked = False
stats.Checked = True
statsframe.Visible = True
bet_total.Caption = Format(bettotal, "#,###,###")
credits_won.Caption = Format(hopperempty, "#,###,###")
deals_label.Caption = Format(totaldeals, "#,###,###")
Call returnrate
ElseIf stats.Checked = False Then
stats.Checked = True
statsframe.Visible = True
bet_total.Caption = Format(bettotal, "#,###,###")
credits_won.Caption = Format(hopperempty, "#,###,###")
deals_label.Caption = Format(totaldeals, "#,###,###")
Call returnrate
ElseIf stats.Checked = True Then
statsframe.Visible = False
stats.Checked = False
End If
cmdOK.Enabled = True

End Sub
Sub ChangeRes(iWidth As Single, iHeight As Single)

    Dim a As Boolean
    Dim i&
    i = 0

    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)

        Dim b&
        DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        DevM.dmPelsWidth = iWidth
        DevM.dmPelsHeight = iHeight
        b = ChangeDisplaySettings(DevM, 0)
End Sub


