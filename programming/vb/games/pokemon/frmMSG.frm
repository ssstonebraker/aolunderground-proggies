VERSION 5.00
Begin VB.Form frmMSG 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   1665
      Left            =   0
      Top             =   0
      Width           =   3990
   End
   Begin VB.Label lblDrag 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   15
      Width           =   4005
   End
   Begin VB.Image c3 
      Height          =   90
      Left            =   1005
      Picture         =   "frmMSG.frx":0000
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   1065
   End
   Begin VB.Image c2 
      Height          =   90
      Left            =   1005
      Picture         =   "frmMSG.frx":09B2
      Stretch         =   -1  'True
      Top             =   2595
      Width           =   1065
   End
   Begin VB.Image c4 
      Height          =   90
      Left            =   1005
      Picture         =   "frmMSG.frx":12EE
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Image c1 
      Height          =   225
      Left            =   0
      Picture         =   "frmMSG.frx":1BF3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4110
   End
   Begin VB.Label lblOkay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label lblMSG 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   3900
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frmA As String
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_GotFocus()
    c1.Picture = c3.Picture
End Sub
Private Sub Form_LostFocus()
    c1.Picture = c4.Picture
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If frmA = "frmBattle" Then
        frmBattle.Enabled = True
    End If
    If frmA = "frmBench" Then
        frmBench.Enabled = True
    End If
    If frmA = "frmCableChat" Then
        frmCableChat.Enabled = True
    End If
    If frmA = "frmCableClub" Then
        frmCableClub.Enabled = True
    End If
    If frmA = "frmChatroom" Then
        frmChatroom.Enabled = True
    End If
    If frmA = "frmChoose" Then
        frmChoose.Enabled = True
    End If
    If frmA = "frmInternetConnect" Then
        frmInternetConnect.Enabled = True
    End If
    If frmA = "frmInternetListen" Then
        frmInternetListen.Enabled = True
    End If
    If frmA = "frmIntro" Then
        frmIntro.Enabled = True
    End If
    If frmA = "frmInventory" Then
        frmInventory.Enabled = True
    End If
    If frmA = "frmLose" Then
        frmLose.Enabled = True
    End If
    If frmA = "frmMain" Then
        frmMain.Enabled = True
    End If
    If frmA = "frmNewGame" Then
        frmNewGame.Enabled = True
    End If
    If frmA = "frmPokedex" Then
        frmPokedex.Enabled = True
    End If
    If frmA = "frmPokeMart" Then
        frmPokeMart.Enabled = True
    End If
    If frmA = "frmPokemonCenter" Then
        frmPokemonCenter.Enabled = True
    End If
    If frmA = "frmPopup" Then
        frmPopup.Enabled = True
    End If
    If frmA = "frmSelect" Then
        frmSelect.Enabled = True
    End If
    If frmA = "frmSendIM" Then
        frmSendIM.Enabled = True
    End If
    If frmA = "frmSplash" Then
        frmSplash.Enabled = True
    End If
    If frmA = "frmStartup" Then
        frmStartup.Enabled = True
    End If
    If frmA = "frmSwitch" Then
        frmSwitch.Enabled = True
    End If
    If frmA = "frmTrade" Then
        frmTrade.Enabled = True
    End If
    If frmA = "frmTravel" Then
        frmTravel.Enabled = True
    End If
    If frmA = "frmUseItem" Then
        frmUseItem.Enabled = True
    End If
    If frmA = "frmWin" Then
        frmWin.Enabled = True
    End If
End Sub
Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    c1.Picture = c2.Picture
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    c1.Picture = c3.Picture
End Sub
Private Sub lblOkay_Click()
    Unload Me
End Sub
Sub Activate(frm As Form)
    frm.Enabled = False
    frmA = frm.name
End Sub
