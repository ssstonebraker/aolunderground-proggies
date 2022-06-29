VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   1455
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "enter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      MaskColor       =   &H0000C000&
      MousePointer    =   10  'Up Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Text            =   "password?"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2520
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "enter your password for auto sign on"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   2520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   2520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line7 
      X1              =   2520
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   120
      Y1              =   1320
      Y2              =   120
   End
   Begin VB.Line Line10 
      X1              =   2520
      X2              =   2520
      Y1              =   120
      Y2              =   1320
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Timer1.Enabled = True
Form3.Hide
End Sub

Private Sub Form_Load()
FormAbove Me
End Sub
