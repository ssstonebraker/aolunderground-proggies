VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Connection Mode"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2055
      TabIndex        =   5
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Frame fraSep1 
      Height          =   60
      Left            =   -105
      TabIndex        =   4
      Top             =   1935
      Width           =   3300
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   1080
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "This means you wait for your opponent in Connect Mode to connect to you."
      Height          =   810
      Left            =   1260
      TabIndex        =   3
      Top             =   1005
      Width           =   1800
   End
   Begin VB.Label lblConnect 
      Caption         =   "This means you enter a hostname to connect to. The other player has to be in Listen Mode."
      Height          =   810
      Left            =   1260
      TabIndex        =   1
      Top             =   90
      Width           =   1800
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdConnect_Click()
    frmInternetConnect.intAction = 1
    frmInternetConnect.Show
    Unload Me
End Sub
Private Sub cmdListen_Click()
    frmInternetListen.intAction = 1
    frmInternetListen.Show
    Unload Me
End Sub
Private Sub Form_Activate()
    FormOnTop Me
End Sub
