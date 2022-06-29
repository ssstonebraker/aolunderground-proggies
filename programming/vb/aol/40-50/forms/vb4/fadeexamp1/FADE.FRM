VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Scroll Bar fade By: N KillaZ"
   ClientHeight    =   2175
   ClientLeft      =   2235
   ClientTop       =   3720
   ClientWidth     =   6690
   Height          =   2580
   Left            =   2175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   6690
   Top             =   3375
   Width           =   6810
   Begin VB.HScrollBar Blue2 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.HScrollBar Green2 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.HScrollBar Red2 
      Height          =   255
      Left            =   240
      Max             =   255
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "what to fade"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.HScrollBar blue1 
      Height          =   255
      Left            =   3480
      Max             =   255
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.HScrollBar green1 
      Height          =   255
      Left            =   3480
      Max             =   255
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.HScrollBar red1 
      Height          =   255
      Left            =   3480
      Max             =   255
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Color2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label color1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   4800
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub blue1_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub Blue2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub


Private Sub Command1_Click()
 FaDe$ = "" & FadeByColor2(color1.BackColor, Color2.BackColor, Text1.Text, False)
SendChat "" & ("" + FaDe$ + "")
End Sub


Private Sub green1_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub


Private Sub HScroll1_Change()

End Sub

Private Sub Green2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub red1_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub


Private Sub Red2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub


