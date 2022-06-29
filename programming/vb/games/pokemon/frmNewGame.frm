VERSION 5.00
Begin VB.Form frmNewGame 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   1515
   ClientLeft      =   7455
   ClientTop       =   1275
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBuffer 
      Height          =   255
      Left            =   2820
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtRival 
      Height          =   270
      Left            =   1635
      MaxLength       =   10
      TabIndex        =   4
      Top             =   750
      Width           =   1905
   End
   Begin VB.TextBox txtPlayer 
      Height          =   270
      Left            =   1635
      MaxLength       =   10
      TabIndex        =   3
      Top             =   420
      Width           =   1905
   End
   Begin VB.Label lblB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
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
      Left            =   3105
      TabIndex        =   9
      Top             =   1050
      Width           =   435
   End
   Begin VB.Label lblR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
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
      Left            =   2700
      TabIndex        =   8
      Top             =   1050
      Width           =   360
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   1515
      Left            =   0
      Top             =   0
      Width           =   3645
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
      Left            =   765
      TabIndex        =   6
      Top             =   1215
      Width           =   600
   End
   Begin VB.Label lblCreate 
      BackStyle       =   0  'Transparent
      Caption         =   "Create"
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
      Left            =   45
      TabIndex        =   5
      Top             =   1215
      Width           =   600
   End
   Begin VB.Label lblRival 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rival Name:"
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
      Left            =   375
      TabIndex        =   2
      Top             =   765
      Width           =   1155
   End
   Begin VB.Label lblPlayer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Player Name:"
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
      Left            =   375
      TabIndex        =   1
      Top             =   450
      Width           =   1155
   End
   Begin VB.Label lblNewGame 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Game"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   225
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1530
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmSplash.Enabled = True
End Sub
Private Sub lblB_Click()
    lblB.ForeColor = &HFFFFFF
    lblR.ForeColor = &HFF&
End Sub
Private Sub lblCancel_Click()
    Unload Me
End Sub
Private Sub lblCreate_Click()
    If txtPlayer.Text = "" Or txtRival.Text = "" Or lblR.ForeColor = &HFF& And lblB.ForeColor = &HFF0000 Then
        MsgBoxA Me, "Player, Rival Name, or Version Missing!"
    Else
        lstBuffer.Clear
        LoadGames lstBuffer
        lstBuffer.AddItem txtPlayer.Text
        SaveGames lstBuffer
        If lblB.ForeColor = &HFFFFFF Then
            CreateGame txtPlayer.Text, txtRival.Text, "Blue"
        Else
            CreateGame txtPlayer.Text, txtRival.Text, "Red"
        End If
        frmSplash.lstPlayers.Clear
        LoadGames frmSplash.lstPlayers
        Unload Me
    End If
End Sub
Private Sub lblNewGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblR_Click()
    lblB.ForeColor = &HFF0000
    lblR.ForeColor = &HFFFFFF
End Sub
