VERSION 5.00
Object = "{69057B54-4F0D-11D2-A11D-549F06C10000}#1.0#0"; "DiceRoller.ocx"
Begin VB.Form frmDice 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dice Active X Control 1.0"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll 'em"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin DiceRoller.Dice Dice3 
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      DiceColor       =   2
   End
   Begin DiceRoller.Dice Dice2 
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      DiceColor       =   1
   End
   Begin DiceRoller.Dice Dice1 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   630
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   585
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   585
   End
End
Attribute VB_Name = "frmDice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRoll_Click()
    Dice1.RollEm
    Dice2.RollEm
    Dice3.RollEm
    lbl1.Caption = Dice1.DiceValue
    lbl2.Caption = Dice2.DiceValue
    lbl3.Caption = Dice3.DiceValue
    Text1.Text = Dice1.DiceValue + Dice2.DiceValue + Dice3.DiceValue
End Sub

Private Sub Form_Load()
    lbl1.Caption = Dice1.DiceValue
    lbl2.Caption = Dice2.DiceValue
    lbl3.Caption = Dice3.DiceValue
End Sub
