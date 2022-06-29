VERSION 5.00
Begin VB.Form frmCustomDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom MineField"
   ClientHeight    =   2925
   ClientLeft      =   4305
   ClientTop       =   2640
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtMines 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtColumns 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtRows 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdEscape 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblMine 
      Caption         =   "&MINES:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblColumn 
      Caption         =   "&COLUMNS:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblRows 
      Caption         =   "&ROWS:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmCustomDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' flag so that no action is taken when ESC key
' is pressed to close Dialog Box
Public mblnEscape As Boolean
Private Sub cmdEscape_Click()
    ' Unload Modal Dialog if ESC pressed as
    ' values from dialog can be discarded.
    mblnEscape = True
    Unload Me
End Sub
Private Sub cmdOK_Click()
    ' Dont unload as yet, so that dialog
    ' values can still be accessed
    Me.Hide
End Sub
Private Sub Form_Load()
    mblnEscape = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mblnEscape = True
End Sub
Private Sub txtColumns_GotFocus()
    ' highlight the text in the textbox
    txtColumns.SelStart = 0
    txtColumns.SelLength = Len(txtColumns)
End Sub
Private Sub txtMines_GotFocus()
    ' highlight the text in the textbox
    txtMines.SelStart = 0
    txtMines.SelLength = Len(txtMines)
End Sub
Private Sub txtRows_GotFocus()
    ' highlight the text in the textbox
    txtRows.SelStart = 0
    txtRows.SelLength = Len(txtRows)
End Sub
