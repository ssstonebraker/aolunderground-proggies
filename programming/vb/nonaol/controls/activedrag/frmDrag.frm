VERSION 5.00
Object = "{D59AB42B-CC72-11D3-B3C7-444553540000}#1.0#0"; "drag.ocx"
Begin VB.Form frmDrag 
   Caption         =   "Active Drag Example"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin ActiveDragControl.ActiveDrag ActiveDrag1 
      Left            =   3120
      Top             =   1470
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblInfo 
      Caption         =   "Example to Active Drag (by sonic). Click on form and drag. Form will snap to edges."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   360
      TabIndex        =   0
      Top             =   75
      Width           =   2970
   End
End
Attribute VB_Name = "frmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveDrag1_MoveX(LeftX As Long)
    Left = LeftX
End Sub

Private Sub ActiveDrag1_MoveY(TopY As Long)
    Top = TopY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveDrag1.SnapRange = 195
    ActiveDrag1.DragNow Width, Height, Left, Top, Button, X, Y
End Sub

