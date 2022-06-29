VERSION 5.00
Begin VB.Form frmInstructBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Playing Instructions"
   ClientHeight    =   6030
   ClientLeft      =   3585
   ClientTop       =   3030
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2190
      TabIndex        =   0
      Top             =   5280
      Width           =   1245
   End
   Begin VB.TextBox txtInstruct 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmInstructBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    
    Dim CRLF As String
    CRLF = Chr$(13) & Chr$(10)
    
    Dim Msg As String
    Msg = CRLF & "Press F2 To Start a New Game." & CRLF & CRLF
    Msg = Msg & "The object of the game is to mark all squares that contain mines. "
    Msg = Msg & "This can be done by opening squares at random initially and then deducing if the surrounding squares could have potential mines. "
    Msg = Msg & "If you open a square that contains a mine, you lose and the game ends. "
    Msg = Msg & "If you open a square that displays a number, it indicates that there are so many mines in the 8 squares surrounding it. "
    Msg = Msg & "To mark a mine, click on a square with the right mouse button. "
    Msg = Msg & "To unmark the mine, clicking again on a marked square with the right mouse button, will display a ?, and yet again will diplay the original square. "
    Msg = Msg & "This is helpful, if you are unsure about a particular square and wish to come back to it later. "
    Msg = Msg & "You could then click the right mouse button twice on that square. This will display a ? on it. "
    Msg = Msg & "To open a square, click on a square with the left mouse button."
    
    txtInstruct = Msg
    
End Sub

