VERSION 4.00
Begin VB.Form Form1 
   ClientHeight    =   2595
   ClientLeft      =   2820
   ClientTop       =   3765
   ClientWidth     =   6690
   Height          =   3000
   Left            =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   6690
   Top             =   3420
   Width           =   6810
   Begin VB.HScrollBar blue6 
      Height          =   135
      Left            =   4080
      Max             =   225
      TabIndex        =   27
      Top             =   2280
      Width           =   1095
   End
   Begin VB.HScrollBar green6 
      Height          =   135
      Left            =   4080
      Max             =   225
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.HScrollBar red6 
      Height          =   135
      Left            =   4080
      Max             =   225
      TabIndex        =   25
      Top             =   1800
      Width           =   1095
   End
   Begin VB.HScrollBar blue5 
      Height          =   135
      Left            =   2040
      Max             =   225
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.HScrollBar green5 
      Height          =   135
      Left            =   2040
      Max             =   225
      TabIndex        =   20
      Top             =   2040
      Width           =   1215
   End
   Begin VB.HScrollBar red5 
      Height          =   135
      Left            =   2040
      Max             =   225
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
   End
   Begin VB.HScrollBar blue4 
      Height          =   135
      Left            =   120
      Max             =   225
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.HScrollBar green4 
      Height          =   135
      Left            =   120
      Max             =   225
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.HScrollBar red4 
      Height          =   135
      Left            =   120
      Max             =   225
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.HScrollBar blue3 
      Height          =   135
      Left            =   4080
      Max             =   225
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.HScrollBar green3 
      Height          =   135
      Left            =   4080
      Max             =   225
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.HScrollBar red3 
      Height          =   135
      Left            =   4080
      Max             =   255
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "wavy on/off"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Text            =   "fasle"
      Top             =   360
      Width           =   495
   End
   Begin VB.HScrollBar Blue2 
      Height          =   135
      Left            =   2040
      Max             =   255
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.HScrollBar Green2 
      Height          =   135
      Left            =   2040
      Max             =   255
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.HScrollBar Red2 
      Height          =   135
      Left            =   2040
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
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.HScrollBar green1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.HScrollBar red1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label color6 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5280
      TabIndex        =   24
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label color5 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   3360
      TabIndex        =   23
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label color4 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1440
      TabIndex        =   22
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label color3 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5280
      TabIndex        =   13
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Color2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label color1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   495
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


Private Sub blue3_Change()
color3.BackColor = RGB(red3.Value, green3.Value, blue3.Value)
End Sub

Private Sub blue4_Change()
color4.BackColor = RGB(red4.Value, green4.Value, blue4.Value)
End Sub

Private Sub blue5_Change()
color5.BackColor = RGB(red5.Value, green5.Value, blue5.Value)
End Sub

Private Sub blue6_Change()
color6.BackColor = RGB(red6.Value, green6.Value, blue6.Value)
End Sub

Private Sub Command1_Click()
 FaDe$ = "" & FadeByColor6(color1.BackColor, Color2.BackColor, color3.BackColor, color4.BackColor, color5.BackColor, color6.BackColor, Text1.Text, text2.Text)
SendChat "" & ("" + FaDe$ + "")
End Sub


Private Sub Command2_Click()
If text2.Text = "true" Then
text2.Text = "false"
Else
text2.Text = "true"
End If
End Sub

Private Sub Command3_Click()
text2.Text = "false"
End Sub

Private Sub green1_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub


Private Sub HScroll1_Change()
color3.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub Green2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub green3_Change()
color3.BackColor = RGB(red3.Value, green3.Value, blue3.Value)
End Sub

Private Sub HScroll2_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)

End Sub

Private Sub HScroll3_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub HScroll4_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub HScroll5_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub HScroll6_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub HScroll7_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub HScroll8_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub HScroll9_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub

Private Sub green4_Change()
color4.BackColor = RGB(red4.Value, green4.Value, blue4.Value)
End Sub

Private Sub green5_Change()
color5.BackColor = RGB(red5.Value, green5.Value, blue5.Value)
End Sub

Private Sub green6_Change()
color6.BackColor = RGB(red6.Value, green6.Value, blue6.Value)
End Sub

Private Sub red1_Change()
color1.BackColor = RGB(red1.Value, green1.Value, blue1.Value)
End Sub


Private Sub Red2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub


Private Sub red3_Change()
color3.BackColor = RGB(red3.Value, green3.Value, blue3.Value)
End Sub


Private Sub red4_Change()
color4.BackColor = RGB(red4.Value, green4.Value, blue4.Value)
End Sub


Private Sub red5_Change()
color5.BackColor = RGB(red5.Value, green5.Value, blue5.Value)
End Sub


Private Sub red6_Change()
color5.BackColor = RGB(red5.Value, green5.Value, blue5.Value)
End Sub


