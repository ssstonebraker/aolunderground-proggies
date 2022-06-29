VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBench2HP 
      Height          =   255
      Left            =   3675
      TabIndex        =   6
      Top             =   1665
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ListBox lstBench2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   2550
      TabIndex        =   1
      Top             =   90
      Width           =   1395
   End
   Begin VB.ListBox lstBench1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1395
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Cancel"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2055
      TabIndex        =   5
      Top             =   1515
      Width           =   1395
   End
   Begin VB.Label lblTrade 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Trade"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   645
      TabIndex        =   4
      Top             =   1515
      Width           =   1395
   End
   Begin VB.Label lblHPDat 
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
      Height          =   225
      Left            =   1485
      TabIndex        =   3
      Top             =   810
      Width           =   1050
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1485
      TabIndex        =   2
      Top             =   585
      Width           =   1050
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Load()
    lstBench1.Clear
    LoadBench lstBench1
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblTrade_Click()
    If lstBench1.ListIndex = -1 Or lstBench2.ListIndex = -1 Then
        MsgBoxA Me, "Please select the Pokémon you wish to trade with!"
    Else
        Me.Hide
        MsgBox "Press OK To Offer This Trade."
        Me.Show
        If frmInternetConnect.Visible = True Then
            frmInternetConnect.sckConnect.SendData "datOFFER-" & lstBench1.ItemData(lstBench1.ListIndex) & "|" & lstBench2.ItemData(lstBench2.ListIndex)
        Else
            frmInternetListen.sckListen.SendData "datOFFER-" & lstBench1.ItemData(lstBench1.ListIndex) & "|" & lstBench2.ItemData(lstBench2.ListIndex)
        End If
        lblTrade.Visible = False
        lblCancel.Visible = False
    End If
End Sub
Private Sub lstBench1_Click()
    If lstBench1.ListIndex = -1 Then
        lblHPDat.Caption = "N/A"
    Else
        lblHPDat.Caption = GetFromINI("6" & lstBench1.ItemData(lstBench1.ListIndex), "6" & lstBench1.ItemData(lstBench1.ListIndex) & ".1", PathA & "\" & LCase(TrimSpaces(frmMain.Player))) & " / " & N2H(lstBench1.ItemData(lstBench1.ListIndex))
    End If
End Sub
Private Sub lstBench2_Click()
    If lstBench2.ListIndex = -1 Then
        lblHPDat.Caption = "N/A"
    Else
        lblHPDat.Caption = lstBench2HP.List(lstBench2.ListIndex) & " / " & N2H(lstBench2.ItemData(lstBench2.ListIndex))
    End If
End Sub
