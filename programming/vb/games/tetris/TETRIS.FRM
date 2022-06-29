VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toxetris By OxOkl"
   ClientHeight    =   7215
   ClientLeft      =   2745
   ClientTop       =   1065
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tetris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Tetris.frx":0442
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   900
      Left            =   120
      ScaleHeight     =   840
      ScaleWidth      =   3690
      TabIndex        =   10
      Top             =   7395
      Width           =   3750
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   3165
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H000000C0&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2205
      Width           =   1440
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   4815
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   5
      Top             =   525
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   960
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000C0&
      Caption         =   "Show Next Piece"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6120
      Left            =   150
      ScaleHeight     =   404
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   1
      Top             =   780
      Width           =   3600
   End
   Begin VB.PictureBox BMP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Left            =   3975
      ScaleHeight     =   5700
      ScaleWidth      =   8700
      TabIndex        =   0
      Top             =   7380
      Visible         =   0   'False
      Width           =   8760
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5010
      TabIndex        =   9
      Top             =   2805
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   5535
      TabIndex        =   8
      Top             =   2775
      Width           =   210
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   4050
      TabIndex        =   7
      Top             =   2850
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   4395
      TabIndex        =   2
      Top             =   1650
      Width           =   915
   End
   Begin VB.Menu Game 
      Caption         =   "&Game"
      Index           =   1
      Begin VB.Menu Game_About 
         Caption         =   "&About"
         Index           =   2
      End
      Begin VB.Menu Game_Start 
         Caption         =   "&Start"
         Index           =   3
      End
      Begin VB.Menu Game_Skip_Level 
         Caption         =   "Skip &Level"
         Index           =   4
      End
      Begin VB.Menu Game_Pause 
         Caption         =   "&Pause"
         Index           =   5
      End
      Begin VB.Menu Game_Exit 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long 'The common counter
Dim j As Long 'Same
Dim File_Num_1 As Long
Dim Path As String

Private Sub Check2_Click()
Picture1.SetFocus
End Sub

Private Sub Command1_Click()
Game_Is_About_To_End = True
Unload Me
End
End Sub

Private Sub Game_Skip_Level_Click(Index As Integer)
If Game_Is_Started = True Then
   Level = Level + 1
   Label3.Caption = LTrim(Str$(Level))
   Vy_Level = Vy_Level + 0.1
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If Game_Is_Started = True Then
   Right_Move_Requested = False
   Left_Move_Requested = False
   Key_Up = False
   If KeyCode = 38 Then
      If Turn_Permission(Piece_No, x, y, False) = True Then
         x = x + Piece(Piece_No).After_Turn_dx
         y = y + Piece(Piece_No).After_Turn_dy
         Piece_No = Piece(Piece_No).After_Turn_Piece_No
         Create_Permission_Database Piece_No
      End If
   End If
      If KeyCode = 40 Then
      Vy = Vy + 0.05
   End If
   If KeyCode = 39 Then
      Right_Move_Requested = True
      If Vx < 0 Then Vx = Vx + 0.2 Else Vx = Vx + 0.1
   End If
   If KeyCode = 37 Then
      Left_Move_Requested = True
      If Vx > 0 Then Vx = Vx - 0.2 Else Vx = Vx - 0.1
   End If
End If
End Sub

Private Sub Form_Load()
Form1.Picture = LoadPicture(VB.App.Path & "\form.bmp")
BMP.Picture = LoadPicture(VB.App.Path & "\pieces.bmp")
Picture3.Picture = LoadPicture(VB.App.Path & "\GameOver.bmp")
'Real randomize
Randomize Timer
'PictureBox scalemodes already setted to VBPixels = 3
BMP.Width = 560
BMP.Height = 400
Picture1.Width = 244
Picture1.Height = 414
Picture2.Width = 110
Picture2.Height = 64
File_Num_1 = FreeFile
'Please do not change the file names etc...
Open VB.App.Path + "\Pieces.dat" For Input As 1
For i = 1 To 19
'Read all the information from the data file
'For all the 19 pieces
   Input #File_Num_1, Piece(i).OnBmp_x, Piece(i).OnBmp_y
   Input #File_Num_1, Piece(i).MaskOnBmp_x, Piece(i).MaskOnBmp_y
   Input #File_Num_1, Piece(i).Width, Piece(i).Height
   Input #File_Num_1, Piece(i).Creating_Piece_No
   Input #File_Num_1, Piece(i).After_Turn_Piece_No
   Input #File_Num_1, Piece(i).Number_Of_Positions_To_Check_Left
   For j = 1 To Piece(i).Number_Of_Positions_To_Check_Left
      Input #File_Num_1, Piece(i).Check_These_Positions_Left_x(j)
      Input #File_Num_1, Piece(i).Check_These_Positions_Left_y(j)
   Next j
   Input #File_Num_1, Piece(i).Number_Of_Positions_To_Check_Right
   For j = 1 To Piece(i).Number_Of_Positions_To_Check_Right
      Input #File_Num_1, Piece(i).Check_These_Positions_Right_x(j)
      Input #File_Num_1, Piece(i).Check_These_Positions_Right_y(j)
   Next j
   Input #File_Num_1, Piece(i).Number_Of_Positions_To_Check_Down
   For j = 1 To Piece(i).Number_Of_Positions_To_Check_Down
      Input #File_Num_1, Piece(i).Check_These_Positions_Down_x(j)
      Input #File_Num_1, Piece(i).Check_These_Positions_Down_y(j)
   Next j
   Input #File_Num_1, Piece(i).Number_Of_Squares
   For j = 1 To Piece(i).Number_Of_Squares
      Input #File_Num_1, Piece(i).Piece_Is_This_x(j)
      Input #File_Num_1, Piece(i).Piece_Is_This_y(j)
   Next j
   Piece(i).After_Turn_dx = (Piece(i).Width - Piece(i).Height) / 2
   Piece(i).After_Turn_dy = (Piece(i).Height - Piece(i).Width) / 2
   Piece(i).Creating_x = (Picture1.Width - Piece(Piece(i).Creating_Piece_No).Width) / 2
   Piece(i).Next_Piece_Pos_x = (Picture2.Width - Piece(i).Width) / 2
   Piece(i).Next_Piece_Pos_y = (Picture2.Height - Piece(i).Height) / 2
Next i
Close
Check1.Enabled = False
Check2.Enabled = False
Check1.Value = vbChecked
Game_Is_Started = False
End Sub

Private Sub Form_Terminate()
Unload Me
End
End Sub

Private Sub Game_About_Click(Index As Integer)
MsgBox "Toxetris... A design of Tetris by Oguz Ozgul 1999", vbOKOnly, "Toxetris About"
End Sub

Private Sub Game_Exit_Click(Index As Integer)
Game_Is_Started = False
Game_Is_About_To_End = True
Unload Me
End
End Sub

Private Sub Game_Pause_Click(Index As Integer)
If Game_Is_Started = True Then
   If Check2.Value = vbUnchecked Then
      Check2.Value = vbChecked
   Else
      Check2.Value = vbUnchecked
   End If
End If
Picture1.SetFocus
End Sub

Private Sub Game_Start_Click(Index As Integer)
If Game_Is_Started = True Then Exit Sub
Score = 0
Label3.Caption = "1"
Picture1.Cls
Picture1_Resize
Picture2.Cls
Label2.Caption = "00:00"
Label1.Caption = "0000000"
Level = 1
Vy_Level = 0.2
Game_Is_Started = True
Check1.Enabled = True
Check2.Enabled = True
For j = 1 To 18
   Position_Empty(j, 0) = False
Next j
For j = 1 To 18
   Position_Empty(j, 11) = False
Next j
For i = 0 To 11
   Position_Empty(18, i) = False
Next i
For i = 1 To 17
   For j = 1 To 10
      Position_Empty(i, j) = True
   Next j
Next i
'Here is the main routine for the game
'It controls the game by calling subs each time
'And the DoEvents lets you to do other things.
'The_Main_Routine:
'End of the main routine
'   DoEvents
'GoTo The_Main_Routine
y = 1
Piece_No = Get_A_Piece
Piece_No = Piece(Piece_No).Creating_Piece_No
x = Piece(Piece_No).Creating_x
Next_Piece_No = Get_A_Piece
Next_Piece_No = Piece(Next_Piece_No).Creating_Piece_No
Picture2.Cls
BitBlt Picture2.hDC, Piece(Next_Piece_No).Next_Piece_Pos_x, Piece(Next_Piece_No).Next_Piece_Pos_y, Piece(Next_Piece_No).Width, Piece(Next_Piece_No).Height, BMP.hDC, Piece(Next_Piece_No).OnBmp_x, Piece(Next_Piece_No).OnBmp_y, SRCCOPY
Vx = 0
Vy = Vy_Level
Create_Permission_Database Piece_No
Piece_Stopped = False
mainRoutine:
   DoEvents
If Check2.Value = vbUnchecked Then
   If Key_Up Then Vx = Vx * 32 / 33
   If Vy > 0 And Key_Up Then Vy = Vy * 119 / 120
   If Abs(Vx) < 0.04 Then Vx = 0
   If Vy < Vy_Level + 0.04 And Abs(Vy) > 0 Then Vy = Vy_Level
   Check_The_Permissions Piece_No, x, y, Vx, Vy
   y = y + Vy
   x = x + Vx
   If x < 0 Then x = 0: Vx = 0
   If x + Piece(Piece_No).Width > 240 Then x = 240 - Piece(Piece_No).Width: Vx = 0
   Draw_The_Piece Prv_Piece_No, Piece_No, Int(x), Int(y), Int(Prv_x), Int(Prv_y)
   Prv_Piece_No = Piece_No
   Prv_x = x
   Prv_y = y
   Prv_Vertical_Stop_Status = Vertical_Stop_Status
   Prv_Left_Stop_Status = Left_Stop_Status
   Prv_Right_Stop_Status = Right_Stop_Status
   If Piece_Stopped = True Then
      Score = Score + 100 * Level
      Copied_To_BMP = False
      Check_Rows_After_Stop
      If Score > (Level ^ 2) * 5000 Then Vy_Level = Vy_Level + 0.1: Level = Level + 1: Label3.Caption = LTrim(Str$(Level))
      Label1.Caption = Right$("0000000" + LTrim(Str$(Score)), 7)
      Piece_No = Piece(Next_Piece_No).Creating_Piece_No
      Next_Piece_No = Get_A_Piece
      Next_Piece_No = Piece(Next_Piece_No).Creating_Piece_No
      Picture2.Cls
      If Check1.Value = vbChecked Then
         BitBlt Picture2.hDC, Piece(Next_Piece_No).Next_Piece_Pos_x, Piece(Next_Piece_No).Next_Piece_Pos_y, Piece(Next_Piece_No).Width, Piece(Next_Piece_No).Height, BMP.hDC, Piece(Next_Piece_No).OnBmp_x, Piece(Next_Piece_No).OnBmp_y, SRCCOPY
      End If
      Piece_Stopped = False
      x = Piece(Piece_No).Creating_x
      y = 1
      Vy = Vy_Level
      Vx = 0
      If Turn_Permission(Piece_No, x, y, True) = False Then
         BitBlt Picture1.hDC, 0, 156, 240, 48, Picture3.hDC, 0, 0, SRCCOPY
         Picture1.Refresh
         For i = 0 To 240 Step 2
            For j = 0 To 408 Step 2
               Picture1.PSet (i, j), 0
            Next j, i
         Game_Is_Started = False
         Exit Sub
      End If
      Create_Permission_Database Piece_No
   End If
End If
GoTo mainRoutine
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
Key_Up = True
Right_Move_Requested = False
Left_Move_Requested = False
End Sub

Private Sub Picture1_Resize()
For i = 408 To 0 Step -24
   Picture1.Line (0, i)-(240, i), RGB(64, 64, 64)
Next i
For i = 240 To 0 Step -24
   Picture1.Line (i, 0)-(i, 408), RGB(64, 64, 64)
Next i
Picture1.Refresh
For i = 408 To 0 Step -24
   For j = 240 To 0 Step -24
      Picture1.PSet (j, i), RGB(96, 96, 96)
   Next j
Next i
End Sub

Private Sub Timer1_Timer()
Dim T1_sminute As String
Dim T1_ssecond As String
Dim T1_lsecond As Long
Dim T1_lminute As Long
If Game_Is_Started = True And Check2.Value = vbUnchecked Then
   T1_sminute = Left(Label2.Caption, 2)
   T1_ssecond = Right(Label2.Caption, 2)
   T1_lminute = Val(T1_sminute)
   T1_lsecond = Val(T1_ssecond)
   T1_lsecond = T1_lsecond + 1
   If T1_lsecond = 61 Then
      T1_lminute = T1_lminute + 1
      T1_lsecond = 0
   End If
   T1_ssecond = Right("00" + LTrim(Str$(T1_lsecond)), 2)
   T1_sminute = Right("00" + LTrim(Str$(T1_lminute)), 2)
   Label2.Caption = T1_sminute + ":" + T1_ssecond
End If
End Sub


