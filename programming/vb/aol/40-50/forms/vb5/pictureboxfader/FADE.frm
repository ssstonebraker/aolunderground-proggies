VERSION 5.00
Begin VB.Form Fade 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Fading Example - By: EccO (xeccox@mailcity.com)"
   ClientHeight    =   1335
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1335
   ScaleWidth      =   7260
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   6915
      TabIndex        =   1
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Click to fade."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   100
      Width           =   6975
   End
End
Attribute VB_Name = "Fade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Color Fading Example - By: EccO
'E-Mail: xeccox@mailcity.com

Private Sub Picture1_Click()

On Error Resume Next
Dim FadeW As Integer
Dim Loo As Integer

Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double

'Change numbers to change the color.
'It's in RGB value.

'Starting color
FirstColor(1) = 255 'HScroll1.Value
FirstColor(2) = 0 'HScroll2.Value
FirstColor(3) = 0 'HScroll3.Value
'Ending color
SecondColor(1) = 0 'HScroll4.Value
SecondColor(2) = 0 'HScroll5.Value
SecondColor(3) = 255 'HScroll6.Value

SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100
FadeW = Picture1.Width / 100

For Loo = 0 To 99
Picture1.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo

End Sub

