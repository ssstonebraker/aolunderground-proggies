VERSION 4.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Glaze 98 Fader v4.o By PUntEr AnD (X)"
   ClientHeight    =   4710
   ClientLeft      =   1950
   ClientTop       =   1395
   ClientWidth     =   6540
   Height          =   5115
   Icon            =   "Fade3color.frx":0000
   Left            =   1890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Top             =   1050
   Width           =   6660
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      ScaleHeight     =   315
      ScaleWidth      =   4275
      TabIndex        =   20
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3600
      Width           =   5775
   End
   Begin VB.HScrollBar blue1 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar green1 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.HScrollBar red1 
      Height          =   255
      Left            =   120
      Max             =   255
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   "Glaze 98 Rulz!!!"
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fade It"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.HScrollBar red2 
      Height          =   255
      Left            =   3360
      Max             =   255
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.HScrollBar blue2 
      Height          =   255
      Left            =   3360
      Max             =   255
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.HScrollBar green2 
      Height          =   255
      Left            =   3360
      Max             =   255
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.HScrollBar blue3 
      Height          =   255
      Left            =   1680
      Max             =   255
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.HScrollBar green3 
      Height          =   255
      Left            =   1680
      Max             =   255
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.HScrollBar red3 
      Height          =   255
      Left            =   1680
      Max             =   255
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fade It Wavy"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label color1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1560
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label color2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   4680
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Color3 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   3240
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Color"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "End"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Blue1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub blue2_Change()
Color2.BackColor = RGB(red2.Value, green2.Value, blue2.Value)

End Sub


Private Sub blue3_Change()
Color3.BackColor = RGB(red3.Value, green3.Value, blue3.Value)
End Sub

Private Sub Command1_Click()
FadedText$ = FadeByColor3(Color1.BackColor, Color2.BackColor, Color3.BackColor, text1.text, False)
text2 = FadedText$
Call FadePreview(Picture1, text2.text)
End Sub


Private Sub Command2_Click()

End Sub


Private Sub Command3_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command4_Click()

End Sub


Private Sub Command5_Click()
FadedText$ = FadeByColor3(Color1.BackColor, Color2.BackColor, Color3.BackColor, text1.text, True)
text2 = FadedText$
Call FadePreview(Picture1, text2)
End Sub


Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub Form_Paint()
FadeFormBlue Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
Form3.Show
End Sub


Private Sub Green1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub

Private Sub green2_Change()
Color2.BackColor = RGB(red2.Value, green2.Value, blue2.Value)

End Sub

Private Sub green3_Change()
Color3.BackColor = RGB(red3.Value, green3.Value, blue3.Value)
End Sub

Private Sub Red1_Change()
Color1.BackColor = RGB(Red1.Value, Green1.Value, Blue1.Value)
End Sub


Private Sub red2_Change()
Color2.BackColor = RGB(red2.Value, green2.Value, blue2.Value)
End Sub


Private Sub red3_Change()
Color3.BackColor = RGB(red3.Value, green3.Value, blue3.Value)

End Sub


