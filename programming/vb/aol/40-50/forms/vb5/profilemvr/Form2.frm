VERSION 5.00
Begin VB.Form Color 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.HScrollBar Blue2 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.HScrollBar Green2 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.HScrollBar Red2 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.HScrollBar Blue1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.HScrollBar Green1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.HScrollBar Red1 
      Height          =   135
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Color2 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox Color1 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Font Color"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "BackGround Color"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change()

End Sub

Private Sub HScroll2_Change()

End Sub

Private Sub HScroll3_Change()

End Sub

Private Sub HScroll4_Change()

End Sub

Private Sub HScroll5_Change()

End Sub

Private Sub HScroll6_Change()

End Sub

Private Sub Blue1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub Blue2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub Command1_Click()
Dim RedA As String, BlueA As String, GreenA As String, BGColor1 As String
Dim RedB As String, BlueB As String, GreenB As String, FontColor1 As String
Dim ColorA As String, ColorB As String, BGColor As String, FontColor As String
Main.Text1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text2.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text3.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text4.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text5.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text6.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text7.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text8.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
Main.Text1.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text2.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text3.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text4.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text5.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text6.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text7.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
Main.Text8.ForeColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
BGColor1$ = RGBtoHEX(Main.Text1.BackColor)
FontColor1$ = RGBtoHEX(Main.Text1.ForeColor)
RedA$ = Right(BGColor1$, 2)
BlueA$ = Left(BGColor1$, 2)
GreenA$ = Mid(BGColor1$, 3, 2)
BGColor$ = RedA$ + GreenA$ + BlueA$
RedB$ = Right(FontColor1$, 2)
BlueB$ = Left(FontColor1$, 2)
GreenB$ = Mid(FontColor1$, 3, 2)
FontColor$ = RedB$ + GreenB$ + BlueB$
Main.Text10 = "<<u>Body BGColor=#" & BGColor$ & "><<u>Font Color=#" & FontColor$ & ">"
Color.Hide
Main.Show
End Sub

Private Sub Command2_Click()
Color.Hide
Main.Show
End Sub

Private Sub Form_Load()
FormOnTop Me
End Sub

Private Sub Green1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub Green2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub

Private Sub Red1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub Red2_Change()
Color2.BackColor = RGB(Red2.Value, Green2.Value, Blue2.Value)
End Sub
