VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   3555
   ClientLeft      =   1095
   ClientTop       =   1185
   ClientWidth     =   3225
   ControlBox      =   0   'False
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Options.frx":030A
   ScaleHeight     =   3555
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   915
   End
   Begin VB.CheckBox chkSkipIntro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Skip Intro Window"
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
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   1965
   End
   Begin VB.CheckBox chkPlaySounds 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Play Sounds"
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
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1965
   End
   Begin VB.CheckBox chkFillLines 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Fill Lines at Start"
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
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1965
   End
   Begin VB.TextBox txtStartingLevel 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
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
      Left            =   1320
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2280
      X2              =   2295
      Y1              =   1560
      Y2              =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2280
      X2              =   2295
      Y1              =   960
      Y2              =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Level (0 - 9):"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
'-------------------------------------------------------
'Hide frmOptions and record the changes made to the
'game options
'-------------------------------------------------------
StartingLevel = Val(txtStartingLevel)
If chkFillLines.Value = 1 Then
    FillLines = True
Else
    FillLines = False
End If
If chkPlaySounds.Value = 1 Then
    PlaySounds = True
Else
    PlaySounds = False
End If
If chkSkipIntro.Value = 1 Then
    HideSplash = True
Else
    HideSplash = False
End If
WriteINIFile
frmOptions.Hide

End Sub


Private Sub Command1_Click()
'-------------------------------------------------------
'Hide frmOptions without recording the changes made to
'the game options
'-------------------------------------------------------
frmOptions.Hide

End Sub


Private Sub Form_Load()
frmOptions.Left = ((frmVBtris.Width - frmOptions.Width) / 2) + frmVBtris.Left
frmOptions.Top = ((frmVBtris.Height - frmOptions.Height) / 2) + frmVBtris.Top
StayOnTop Me
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label2_Click()
StartingLevel = Val(txtStartingLevel)
If chkFillLines.Value = 1 Then
    FillLines = True
Else
    FillLines = False
End If
If chkPlaySounds.Value = 1 Then
    PlaySounds = True
Else
    PlaySounds = False
End If
If chkSkipIntro.Value = 1 Then
    HideSplash = True
Else
    HideSplash = False
End If
WriteINIFile
frmOptions.Hide
End Sub

Private Sub Label3_Click()
frmOptions.Hide
End Sub

Private Sub txtStartingLevel_LostFocus()
'-------------------------------------------------------
'Ensure that an acceptable value has been entered
'-------------------------------------------------------
If Not (IsNumeric(txtStartingLevel)) Then
    MsgBox "You must enter a number between 0 and 9!", , "Full Moon Tetris"
    txtStartingLevel = StartingLevel
ElseIf (txtStartingLevel < 0) Or (txtStartingLevel > 9) Then
    MsgBox "You must enter a number between 0 and 9!", , "Full Moon Tetris"
    txtStartingLevel = StartingLevel
End If

End Sub


