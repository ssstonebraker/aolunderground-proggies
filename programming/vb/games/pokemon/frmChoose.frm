VERSION 5.00
Begin VB.Form frmChoose 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   1395
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
      ItemData        =   "frmChoose.frx":0000
      Left            =   75
      List            =   "frmChoose.frx":0002
      TabIndex        =   0
      Top             =   330
      Width           =   1290
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label lblUse 
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
      Left            =   75
      TabIndex        =   3
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label lblCancel 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   2
      Top             =   1560
      Width           =   570
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
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1290
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Load()
    lstBench.Clear
    LoadBench lstBench
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmUseItem.Show
End Sub
Private Sub lblCancel_Click()
    Unload Me
End Sub
Private Sub lblUse_Click()
    If Not GetHealth(lstBench.ItemData(lstBench.ListIndex)) = 0 Then
        MsgBoxA Me, "This will have no effect."
    Else
        SetHealth lstBench.ItemData(lstBench.ListIndex), Fix(N2H(lstBench.ItemData(lstBench.ListIndex)) / 1.25)
        frmBattle.SetStatus N2N(frmBattle.Pokemon1) & " has been revived!"
        DeleteItem lstItems.ListIndex + 12
        If frmBattle.TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "ITM-Revive on " & N2N(lstBench.ItemData(lstBench.ListIndex))
        Else
            frmInternetListen.sckListen.SendData "ITM-Revive on " & N2N(lstBench.ItemData(lstBench.ListIndex))
        End If
        frmBattle.HP1 = GetHealth(lstBench.ItemData(lstBench.ListIndex))
        frmBattle.AddHP1 0
        frmBattle.Turn = 0
        Unload Me
        Unload frmUseItem
    End If
End Sub
