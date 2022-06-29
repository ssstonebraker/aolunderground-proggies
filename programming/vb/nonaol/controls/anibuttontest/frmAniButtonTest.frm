VERSION 5.00
Object = "{EAAF0D04-8757-11D1-B5BB-8033ED902553}#1.0#0"; "XAniButton.ocx"
Begin VB.Form frmAniButtonTest 
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin AniButton.XAniButton XAniButton1 
      Height          =   330
      Left            =   1905
      TabIndex        =   1
      Top             =   315
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   582
      AnimationStyle  =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1005
      Width           =   3780
   End
End
Attribute VB_Name = "frmAniButtonTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    XAniButton1.AnimationStyle = Combo1.ListIndex
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Always Animate"
    Combo1.AddItem "Animate ONLY when Mouse Over"
    Combo1.AddItem "Animate ONLY when Mouse Clicked"
    Combo1.AddItem "Stop Animate when Mouse Over"
    Combo1.AddItem "Stop Animate when Mouse Clicked"
    Combo1.ListIndex = XAniButton1.AnimationStyle
End Sub

