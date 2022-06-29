VERSION 5.00
Begin VB.Form frmPokeMart 
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
      Height          =   1785
      ItemData        =   "frmPokeMart.frx":0000
      Left            =   2400
      List            =   "frmPokeMart.frx":005D
      TabIndex        =   3
      Top             =   555
      Width           =   1620
   End
   Begin VB.Label lblSell 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sell"
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
      Left            =   4110
      TabIndex        =   9
      Top             =   1845
      Width           =   810
   End
   Begin VB.Label lblBuy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase"
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
      Left            =   4110
      TabIndex        =   8
      Top             =   2085
      Width           =   810
   End
   Begin VB.Label lblOwnDAT 
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
      Left            =   4035
      TabIndex        =   7
      Top             =   1245
      Width           =   990
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
      Left            =   4245
      TabIndex        =   6
      Top             =   1020
      Width           =   465
   End
   Begin VB.Label lblCosDAT 
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
      Left            =   4020
      TabIndex        =   5
      Top             =   810
      Width           =   990
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cost:"
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
      Left            =   4260
      TabIndex        =   4
      Top             =   585
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2460
      Left            =   0
      Top             =   0
      Width           =   5070
   End
   Begin VB.Label lblMonDAT 
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
      Left            =   3630
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label lblMoney 
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
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
      Left            =   2415
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
   Begin VB.Shape shpBgBorder 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   2220
      Left            =   120
      Top             =   120
      Width           =   2220
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
   Begin VB.Image imgBg 
      Height          =   2160
      Left            =   150
      Picture         =   "frmPokeMart.frx":011A
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmPokeMart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgBg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub Form_Activate()
    FormOnTop Me
    If GetMoney = "" Then
        lblMonDAT.Caption = "$0"
    Else
        lblMonDAT.Caption = "$" & GetMoney & ".00"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub lblBuy_Click()
    If lstItems.ListIndex = -1 Then
        MsgBoxA Me, "Select an item to purchase first!"
    Else
        If GetMoney - lstItems.ItemData(lstItems.ListIndex) < 0 Then
            MsgBoxA Me, "Insufficient funds to purchase this item!"
        Else
            DeleteMoney lstItems.ItemData(lstItems.ListIndex)
            SaveItem lstItems.ListIndex
            If GetMoney = "" Then
                lblMonDAT.Caption = "$0"
            Else
                lblMonDAT.Caption = "$" & GetMoney & ".00"
            End If
            If GetItem(lstItems.ListIndex) = "" Then
                lblOwnDAT.Caption = "0"
            Else
                lblOwnDAT.Caption = GetItem(lstItems.ListIndex)
            End If
        End If
    End If
End Sub
Private Sub lblExit_Click()
    Unload Me
End Sub
Private Sub lblSell_Click()
    If lstItems.ListIndex = -1 Then
        MsgBoxA Me, "Select an item to sell first!"
    Else
        If lblOwnDAT.Caption = "0" Then
            MsgBoxA Me, "You do not have that item!"
        Else
            SaveMoney Int(lblCosDAT.Caption) / 2
            DeleteItem lstItems.ListIndex
            If GetMoney = "" Then
                lblMonDAT.Caption = "$0"
            Else
                lblMonDAT.Caption = "$" & GetMoney & ".00"
            End If
            If GetItem(lstItems.ListIndex) = "" Then
                lblOwnDAT.Caption = "0"
            Else
                lblOwnDAT.Caption = GetItem(lstItems.ListIndex)
            End If
        End If
    End If
End Sub
Private Sub lstItems_Click()
    lblCosDAT.Caption = "$" & lstItems.ItemData(lstItems.ListIndex)
    If GetItem(lstItems.ListIndex) = "" Then
        lblOwnDAT.Caption = "0"
    Else
        lblOwnDAT.Caption = GetItem(lstItems.ListIndex)
    End If
End Sub
