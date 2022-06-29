VERSION 5.00
Object = "{695A91C7-0CD7-11D1-BF9E-00AA0059999E}#2.0#0"; "Box.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3672
   ClientLeft      =   1236
   ClientTop       =   1704
   ClientWidth     =   5592
   LinkTopic       =   "Form1"
   ScaleHeight     =   3672
   ScaleWidth      =   5592
   Begin Meter.StatusMeter StatusMeter1 
      Height          =   3312
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5232
      _ExtentX        =   9229
      _ExtentY        =   5842
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3312
      Left            =   5280
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   1
      Width           =   252
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      Left            =   0
      Max             =   100
      TabIndex        =   0
      Top             =   3360
      Width           =   5232
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change()
StatusMeter1.Value = HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()
StatusMeter1.Value = HScroll1.Value
End Sub


Private Sub VScroll1_Change()
StatusMeter1.Depth = VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()
StatusMeter1.Depth = VScroll1.Value
End Sub


