VERSION 5.00
Begin VB.Form frmAboutBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About WinMine"
   ClientHeight    =   2085
   ClientLeft      =   5685
   ClientTop       =   3030
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2085
   ScaleWidth      =   3585
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   525
      Left            =   1200
      TabIndex        =   0
      Top             =   1290
      Width           =   1245
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   -15
      TabIndex        =   1
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "frmAboutBox"
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
    Msg = CRLF & CRLF & "Author: NAME HERE" & CRLF
    Msg = Msg & "Copyright (c) 1999 Whatever."
    
    lblAbout.Caption = Msg

End Sub

