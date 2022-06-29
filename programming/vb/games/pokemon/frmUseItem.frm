VERSION 5.00
Begin VB.Form frmUseItem 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   LinkTopic       =   "frmUseItem"
   ScaleHeight     =   2775
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmUseItem.frx":0000
      Left            =   120
      List            =   "frmUseItem.frx":0024
      TabIndex        =   0
      Top             =   300
      Width           =   1545
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   1785
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
      Left            =   555
      TabIndex        =   6
      Top             =   2490
      Width           =   420
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1485
      TabIndex        =   5
      Top             =   45
      Width           =   165
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Item"
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
      Left            =   60
      TabIndex        =   4
      Top             =   45
      Width           =   1290
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
      Left            =   975
      TabIndex        =   3
      Top             =   2490
      Width           =   690
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
      Left            =   120
      TabIndex        =   2
      Top             =   2490
      Width           =   375
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
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1080
   End
End
Attribute VB_Name = "frmUseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmBattle.Show
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblExit_Click()
    Unload Me
End Sub
Private Sub lblUse_Click()
    If lstItems.ListIndex = -1 Then
        MsgBoxA Me, "Select an item to use first!"
    ElseIf GetItem(lstItems.ListIndex + 12) = Empty Or GetItem(lstItems.ListIndex + 12) = "0" Then
        MsgBoxA Me, "You do not have any of that item!"
    ElseIf lstItems.ListIndex = 0 Then
        If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = 0 Or GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = N2H(frmBattle.Pokemon1) Then
            MsgBoxA Me, "This will have no effect."
        Else
            If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) + 10 > N2H(Na2N(LCase(N2N(frmBattle.Pokemon1)))) Then
                SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), N2H(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            Else
                SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) + 10
            End If
            frmBattle.HP1 = GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            frmBattle.AddHP1 0
            frmBattle.SetStatus N2N(frmBattle.Pokemon1) & "'s HP restored by 10!"
            DeleteItem lstItems.ListIndex + 12
            If frmBattle.TurnA = 0 Then
                frmInternetConnect.sckConnect.SendData "ITM-Potion"
            Else
                frmInternetListen.sckListen.SendData "ITM-Potion"
            End If
            frmBattle.Turn = 0
            Unload Me
        End If
    ElseIf lstItems.ListIndex = 1 Then
        If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = 0 Or GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = N2H(frmBattle.Pokemon1) Then
            MsgBoxA Me, "This will have no effect."
        Else
            If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) + 15 > N2H(Na2N(LCase(N2N(frmBattle.Pokemon1)))) Then
                SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), N2H(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            Else
                SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) + 15
            End If
            frmBattle.HP1 = GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            frmBattle.AddHP1 0
            frmBattle.SetStatus N2N(frmBattle.Pokemon1) & "'s HP restored by 15!"
            DeleteItem lstItems.ListIndex + 12
            If frmBattle.TurnA = 0 Then
                frmInternetConnect.sckConnect.SendData "ITM-Super Potion"
            Else
                frmInternetListen.sckListen.SendData "ITM-Super Potion"
            End If
            frmBattle.Turn = 0
            Unload Me
        End If
    ElseIf lstItems.ListIndex = 2 Then
        If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = 0 Or GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = N2H(frmBattle.Pokemon1) Then
            MsgBoxA Me, "This will have no effect."
        Else
            If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) + 20 > N2H(Na2N(LCase(N2N(frmBattle.Pokemon1)))) Then
                SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), N2H(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            Else
                SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) + 20
            End If
            frmBattle.HP1 = GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            frmBattle.AddHP1 0
            frmBattle.SetStatus N2N(frmBattle.Pokemon1) & "'s HP restored by 20!"
            DeleteItem lstItems.ListIndex + 12
            If frmBattle.TurnA = 0 Then
                frmInternetConnect.sckConnect.SendData "ITM-Hyper Potion"
            Else
                frmInternetListen.sckListen.SendData "ITM-Hyper Potion"
            End If
            frmBattle.Turn = 0
            Unload Me
        End If
    ElseIf lstItems.ListIndex = 3 Then
        If GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = 0 Or GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1)))) = N2H(frmBattle.Pokemon1) Then
            MsgBoxA Me, "This will have no effect."
        Else
            SetHealth Na2N(LCase(N2N(frmBattle.Pokemon1))), N2H(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            frmBattle.SetStatus N2N(frmBattle.Pokemon1) & "'s HP at max!"
            DeleteItem lstItems.ListIndex + 12
            If frmBattle.TurnA = 0 Then
                frmInternetConnect.sckConnect.SendData "ITM-Max Potion"
            Else
                frmInternetListen.sckListen.SendData "ITM-Max Potion"
            End If
            frmBattle.HP1 = GetHealth(Na2N(LCase(N2N(frmBattle.Pokemon1))))
            frmBattle.AddHP1 0
            frmBattle.Turn = 0
            Unload Me
        End If
    ElseIf lstItems.ListIndex = 5 Then
        frmChoose.Show
    End If
End Sub
Private Sub lstItems_Click()
    If GetItem(lstItems.ListIndex + 12) = "" Then
        lblOwnDAT.Caption = "0"
    Else
        lblOwnDAT.Caption = GetItem(lstItems.ListIndex + 12)
    End If
End Sub
