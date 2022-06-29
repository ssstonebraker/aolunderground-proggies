VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Loading..."
   ClientHeight    =   2010
   ClientLeft      =   1095
   ClientTop       =   1185
   ClientWidth     =   3105
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash.frx":030A
   ScaleHeight     =   2010
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading... Tetris"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "By: Dolan && Hydro"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Moon Tetris"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2280
      Picture         =   "Splash.frx":14C5E
      Top             =   1200
      Width           =   480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   3915
      X2              =   3915
      Y1              =   1080
      Y2              =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

End Sub

Private Sub Form_Activate()

'-------------------------------------------------------
'Read in the options and load the form various forms to
'prevent delays during run time.  Then hide the form or
'enable cmdOk depending on what HideSplash is equal to.
'-------------------------------------------------------
DoEvents
Label3.Caption = "Loading... Tetris"
ReadINIFile
Label3.Caption = "Almost Ready..."
Load frmPics
Load frmHighScore
Load frmAbout
Load frmInstruct
Load frmVBtris
If PlaySounds Then PlayWAV ONE_LINE
DoEvents
Label3.Caption = "Get Ready To Play..."
If HideSplash Then
    frmSplash.Hide
    frmVBtris.Show
Else
    Label4.Enabled = True
End If

End Sub

Private Sub Form_Load()
Label4.ForeColor = QBColor(15)
'Position the form
frmSplash.Left = (Screen.Width - frmSplash.Width) / 2
frmSplash.Top = (Screen.Height - frmSplash.Height) / 2
'Center the OK command button on the form
Label4.Left = (frmSplash.Width - Label4.Width) / 2

DoEvents
frmSplash.Show
DoEvents
StayOnTop Me
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label4_Click()
'Hide the form
frmSplash.Hide
'Show frmVBtris
frmVBtris.Show

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = QBColor(12)
End Sub
