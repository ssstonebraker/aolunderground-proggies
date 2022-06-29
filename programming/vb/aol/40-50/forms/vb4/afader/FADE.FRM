VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Scroll Bar fade V2 By: N KillaZ"
   ClientHeight    =   1425
   ClientLeft      =   2235
   ClientTop       =   3720
   ClientWidth     =   6690
   Height          =   1830
   Left            =   2175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   6690
   Top             =   3375
   Width           =   6810
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send WaVy"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.HScrollBar Blue2 
      Height          =   135
      Left            =   0
      Max             =   255
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.HScrollBar Green2 
      Height          =   135
      Left            =   0
      Max             =   255
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.HScrollBar Red2 
      Height          =   135
      Left            =   0
      Max             =   255
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "what to fade"
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.HScrollBar blue1 
      Height          =   135
      Left            =   2400
      Max             =   255
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.HScrollBar green1 
      Height          =   135
      Left            =   2400
      Max             =   255
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.HScrollBar red1 
      Height          =   135
      Left            =   2400
      Max             =   255
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Code"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Color2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label color1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   1095
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
FadedText$ = FadeByColor2(color1.BackColor, Color2.BackColor, Text1.Text, False)
Text2 = FadedText$
End Sub


Private Sub Command2_Click()
 FaDe$ = "" & FadeByColor2(color1.BackColor, Color2.BackColor, Text1.Text, True)
SendChat "" & ("" + FaDe$ + "")
FadedText$ = FadeByColor2(color1.BackColor, Color2.BackColor, Text1.Text, True)
Text1 = FadedText$
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


