VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help?"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK?"
      Height          =   735
      Left            =   3360
      TabIndex        =   10
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "Ship"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Hit"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Miss"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Water/Undetermined"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Help is needed?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim streng As String
    streng = "This is a clone of an old favourite, the Battleship" + vbNewLine + _
        vbNewLine + _
        "First of all you'll have to choose opposition. Ranging from" + _
        " mate(stupid) to admiral(clever old bastard). And which player" + _
        " that should start." + vbNewLine + _
        vbNewLine + _
        "Thereafter is just a matter of starting a new game. You'll of" + _
        " course shoot in the grid to the left. And only the blue squares." + _
        " When one of you has hit sixteen squares then we've got a" + _
        " winner." + vbNewLine + vbNewLine + _
        "What the colors mean:"
    Label2.Caption = streng
End Sub
