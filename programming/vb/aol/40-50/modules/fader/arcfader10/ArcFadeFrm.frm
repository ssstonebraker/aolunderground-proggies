VERSION 5.00
Begin VB.Form ArcFadeFrm 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Knight Rider"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hexagon Pulse_Flash"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Square Flash"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Circle Pulse"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Text Flash Fore_Back"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Label Pulse Fore & Back"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Label Flash"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Testing!!!                Testing!!!           Testing!!!           Testing!!!"
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Form Pulse"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   2880
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   2640
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   2400
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   2160
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   1920
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   1680
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   1440
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   1200
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   3480
      Shape           =   5  'Rounded Square
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1800
      Top             =   1440
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Testing!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "ArcFadeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do

Call PulseFadeBack_Red_Yellow(ArcFadeFrm, 0.00001)
Loop
End Sub




Private Sub Command2_Click()
Do
Call FlashFadeBack_Black(Text1, 0.1)
Call FlashFadeFore_Yellow(Text1, 0.1)
Loop
End Sub

Private Sub Command3_Click()
Do
Call FlashFadeBack_Purple(Label1, 0.1)
Loop
End Sub

Private Sub Command4_Click()
Do
Call PulseFadeBack_Black_Yellow(Label1, 0.0001)
Call PulseFadeFore_Black_Yellow(Label1, 0.0001)
Loop
End Sub

Private Sub Command5_Click()
Do
Call PulseFadeBack_Red_Yellow(Shape1, 0.0001)
Loop
End Sub

Private Sub Command6_Click()
Do
Call FlashFadeBack_Red(Shape2, 0.0001)
Loop
End Sub

Private Sub Command7_Click()
Do
Call PulseFadeBack_White_Black(Shape3, 0.0001)
For i = 0 To 3
Call FlashFadeBack_Red(Shape3, 0.0001)
Next i
Loop
End Sub

Private Sub Command8_Click()
Do
Call FlashFadeBack_Red(Shape5, 0.0001)
Call FlashFadeBack_Red(Shape6, 0.0001)
Call FlashFadeBack_Red(Shape7, 0.0001)
Call FlashFadeBack_Red(Shape8, 0.0001)
Call FlashFadeBack_Red(Shape9, 0.0001)
Call FlashFadeBack_Red(Shape10, 0.0001)
Call FlashFadeBack_Red(Shape11, 0.0001)
Call FlashFadeBack_Red(Shape10, 0.0001)
Call FlashFadeBack_Red(Shape9, 0.0001)
Call FlashFadeBack_Red(Shape8, 0.0001)
Call FlashFadeBack_Red(Shape7, 0.0001)
Call FlashFadeBack_Red(Shape6, 0.0001)
Call FlashFadeBack_Red(Shape5, 0.0001)
Call FlashFadeBack_Red(Shape4, 0.0001)
Loop
End Sub

