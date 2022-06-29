VERSION 5.00
Begin VB.Form frmFadeEx 
   Caption         =   "Fader Example"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox FadeMetxt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "Text to Fade"
      ToolTipText     =   "Text to Fade"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton CopyButt 
      Caption         =   "Copy Fade HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox HTMLtxt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Fade HTML"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.HScrollBar Blue2 
      Height          =   135
      Left            =   1440
      Max             =   255
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.HScrollBar Green2 
      Height          =   135
      Left            =   1440
      Max             =   255
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar Red2 
      Height          =   135
      Left            =   1440
      Max             =   255
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.HScrollBar Blue1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.HScrollBar Green1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar Red1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Color2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Color1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFadeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Blue1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)

End Sub

Private Sub Blue1_Scroll()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)

End Sub


Private Sub Blue2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)

End Sub

Private Sub Blue2_Scroll()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)

End Sub


Private Sub CopyButt_Click()
Clipboard.Clear
DoEvents
CopyMe$ = FadeByColor2(Color1.BackColor, Color2.BackColor, FadeMetxt.Text, False)
Clipboard.SetText CopyMe$
End Sub



Private Sub Form_Load()
MsgBox "This example requires MONKEFADE.bas"

End Sub

Private Sub Green1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)

End Sub

Private Sub Green1_Scroll()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)

End Sub


Private Sub Green2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)

End Sub

Private Sub Green2_Scroll()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)

End Sub


Private Sub Red1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)

End Sub

Private Sub Red1_Scroll()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)

End Sub


Private Sub Red2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)

End Sub


Private Sub Red2_Scroll()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)

End Sub


