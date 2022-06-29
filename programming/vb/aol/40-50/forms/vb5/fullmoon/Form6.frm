VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form17"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   1290
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Text            =   "Text Here"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1800
      Width           =   11295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " _"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Macro Font 1"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Command1_Click()
TBhWnd% = Text1.hwnd
StartLine% = SendMessage(TBhWnd%, EM_LINEFROMCHAR, Text1.SelStart, 0)
LineCount% = SendMessage(TBhWnd%, EM_GETLINECOUNT, 0, 0)
FontLineCount% = SendMessage(Text2.hwnd, EM_GETLINECOUNT, 0, 0)
For X% = StartLine% To StartLine% + (FontLineCount% - 1)
   If LineCount% - 1 < X% Then Text1.Text = Text1.Text + Chr$(13) + Chr$(10)
   theline$ = Space(10000)
   LineLength% = SendMEssageByString(TBhWnd%, EM_GETLINE, X%, theline$)
   theline$ = Left(theline$, LineLength%)
   TransferLine$ = Space(10000)
   LineLength% = SendMEssageByString(Text2.hwnd, EM_GETLINE, X% - StartLine%, TransferLine$)
   TransferLine$ = Left(TransferLine$, LineLength%)
   SelStart% = SendMessage(TBhWnd%, EM_LINEINDEX, X%, ByVal 0)
   SelEnd% = SendMessage(TBhWnd%, EM_LINEINDEX, X% + 1, ByVal 0)
   If SelEnd% = -1 Then SelEnd% = Len(Text1.Text) + 2
   Text1.SelStart = SelStart%
   Text1.SelLength = (SelEnd% - 2) - SelStart%
   theline$ = theline$ + TransferLine$
   Text1.SelText = theline$
Next X%
Text1.SelStart = (SendMessage(Text1.hwnd, EM_LINEINDEX, StartLine% + 1, ByVal 0) - 2)
End Sub


Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label2_Click()
Form6.WindowState = 1
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Form4.Text1 = Text1.Text
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Then
Text2.Text = ""
Text1.Text = ""
Text3.Text = ""
Form4.Text1 = ""
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "A" Then
Text2.Text = "''  '/¯/   ''      " + Chr$(13) + Chr$(10) + "  /÷/\¯\ ''      " + Chr$(13) + Chr$(10) + "'/÷/ ¯\÷\'      " + Chr$(13) + Chr$(10) + ".|_|    '|_|      "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "B" Then
Text2.Text = "|¯|\¯\           " + Chr$(13) + Chr$(10) + "|÷|/_/           " + Chr$(13) + Chr$(10) + "|÷|\¯\           " + Chr$(13) + Chr$(10) + "|_|/_/           "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "C" Then
Text2.Text = " /¯/\¯\        " + Chr$(13) + Chr$(10) + "'|÷|   ¯        " + Chr$(13) + Chr$(10) + "'|÷|   _        " + Chr$(13) + Chr$(10) + " \_\/_/        "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "D" Then
Text2.Text = "|¯|\¯\''          " + Chr$(13) + Chr$(10) + "|÷| |÷|.          " + Chr$(13) + Chr$(10) + "|÷| |÷|.          " + Chr$(13) + Chr$(10) + "|_|/_/.''         "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "E" Then
Text2.Text = "|¯|¯¯|           " + Chr$(13) + Chr$(10) + "|÷|¯|.'''          " + Chr$(13) + Chr$(10) + "|÷|_|'¸.          " + Chr$(13) + Chr$(10) + "|_|__|           "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "F" Then
Text2.Text = "|¯|¯¯|           " + Chr$(13) + Chr$(10) + "|÷|¯¯|           " + Chr$(13) + Chr$(10) + "|÷|¯¯'           " + Chr$(13) + Chr$(10) + "|_|'               "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "G" Then
Text2.Text = " /¯/\¯\          " + Chr$(13) + Chr$(10) + "'|÷|   ¯          " + Chr$(13) + Chr$(10) + "'|÷| \¯\          " + Chr$(13) + Chr$(10) + " \_\/_/          "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "H" Then
Text2.Text = "|¯|_'|¯|         " + Chr$(13) + Chr$(10) + "|÷|_'|÷|         " + Chr$(13) + Chr$(10) + "|÷|¯'|÷|         " + Chr$(13) + Chr$(10) + "|_| ''|_|''        "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "I" Then
Text2.Text = "|¯|              " + Chr$(13) + Chr$(10) + "|÷|              " + Chr$(13) + Chr$(10) + "|÷|              " + Chr$(13) + Chr$(10) + "|_|              "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "J" Then
Text2.Text = "    ''|¯|        " + Chr$(13) + Chr$(10) + "    ''|÷|        " + Chr$(13) + Chr$(10) + "\¯\ '|÷|         " + Chr$(13) + Chr$(10) + "' \_\|_|         "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "K" Then
Text2.Text = "|¯|' /¯/         " + Chr$(13) + Chr$(10) + "|÷|/_/ '         " + Chr$(13) + Chr$(10) + "|÷|\¯\ '         " + Chr$(13) + Chr$(10) + "|_| '\_\         "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "L" Then
Text2.Text = "|¯|'               " + Chr$(13) + Chr$(10) + "|÷|'               " + Chr$(13) + Chr$(10) + "|÷|__'           " + Chr$(13) + Chr$(10) + "|_|__|           "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "M" Then
Text2.Text = "|¯|\/|¯|         " + Chr$(13) + Chr$(10) + "|÷|\/|÷|         " + Chr$(13) + Chr$(10) + "|÷|  |÷|         " + Chr$(13) + Chr$(10) + "|_|  |_|         "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "N" Then
Text2.Text = "|¯|\|¯|          " + Chr$(13) + Chr$(10) + "|÷|\|÷|          " + Chr$(13) + Chr$(10) + "|÷| |÷|          " + Chr$(13) + Chr$(10) + "|_| |_|          "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "O" Then
Text2.Text = " /¯/\¯\          " + Chr$(13) + Chr$(10) + "'|÷|  |÷|'         " + Chr$(13) + Chr$(10) + "'|÷|  |÷|'         " + Chr$(13) + Chr$(10) + " \_\/_/.         "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "P" Then
Text2.Text = "|¯|\¯\           " + Chr$(13) + Chr$(10) + "|÷|/_/           " + Chr$(13) + Chr$(10) + "|÷|.              " + Chr$(13) + Chr$(10) + "|_|.              "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "Q" Then
Text2.Text = " /¯/\¯\         " + Chr$(13) + Chr$(10) + "'|÷|  |÷|'        " + Chr$(13) + Chr$(10) + "'|÷|  |÷|_ '     " + Chr$(13) + Chr$(10) + " \_\/_/\_\     "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "R" Then
Text2.Text = "|¯|\¯'\          " + Chr$(13) + Chr$(10) + "|÷|/_/'          " + Chr$(13) + Chr$(10) + "|÷|\¯\'          " + Chr$(13) + Chr$(10) + "|_| |_|''.        "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "S" Then
Text2.Text = "/¯/\¯\.          " + Chr$(13) + Chr$(10) + "\_\ |_|..'        " + Chr$(13) + Chr$(10) + "|¯|\¯\'../        " + Chr$(13) + Chr$(10) + "\_\/_/.          "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "T" Then
Text2.Text = "|¯||¯|¯|       " + Chr$(13) + Chr$(10) + " ¯'|÷|¯ ' '    " + Chr$(13) + Chr$(10) + "  ''|÷|  '       " + Chr$(13) + Chr$(10) + "  ''|_|  '       "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "U" Then
Text2.Text = "|¯|  |¯|         " + Chr$(13) + Chr$(10) + "|÷|  |÷|         " + Chr$(13) + Chr$(10) + "|÷|  |÷|         " + Chr$(13) + Chr$(10) + "'\_\/_/         "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "V" Then
Text2.Text = "\¯\      /¯/     " + Chr$(13) + Chr$(10) + " '\÷\   /÷/ ''    " + Chr$(13) + Chr$(10) + "   \÷\/÷/        " + Chr$(13) + Chr$(10) + "    '\_\/''        "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "W" Then
Text2.Text = "\¯\'          /¯/." + Chr$(13) + Chr$(10) + " '\÷\'       /÷/'' " + Chr$(13) + Chr$(10) + "   \÷\/¯'\/÷/.   " + Chr$(13) + Chr$(10) + "    '\_\/\/_/.    "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "X" Then
Text2.Text = "\¯\   /¯/        " + Chr$(13) + Chr$(10) + " '\÷\/÷/ ''       " + Chr$(13) + Chr$(10) + " '/÷/\÷\ ''       " + Chr$(13) + Chr$(10) + "/_/   \_\        "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "Y" Then
Text2.Text = "\¯\./¯/        " + Chr$(13) + Chr$(10) + " '\÷\/ ''        " + Chr$(13) + Chr$(10) + "  '|÷|           " + Chr$(13) + Chr$(10) + "  '|_|           "
Command1_Click
End If
If UCase(Chr(KeyAscii)) = "Z" Then
Text2.Text = "|¯||¯|''.         " + Chr$(13) + Chr$(10) + " ¯/÷/           " + Chr$(13) + Chr$(10) + " '/÷/_''.        " + Chr$(13) + Chr$(10) + "/_/'|_|          "
Command1_Click
End If

End Sub
