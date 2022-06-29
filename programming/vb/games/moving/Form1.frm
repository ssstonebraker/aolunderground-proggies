VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move Everything Around"
   ClientHeight    =   0
   ClientLeft      =   2025
   ClientTop       =   2655
   ClientWidth     =   1800
   Height          =   435
   Left            =   1965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   1800
   Top             =   2280
   Width           =   1920
   Begin VB.Label Roll 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "roll"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   4080
      TabIndex        =   2
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   4620
      Stretch         =   -1  'True
      Top             =   315
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   3405
      Stretch         =   -1  'True
      Top             =   315
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   72
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   1485
      TabIndex        =   1
      Top             =   105
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   48
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1425
      Left            =   135
      TabIndex        =   0
      Top             =   30
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Single
Dim OldY As Single
Dim OldcX As Single
Dim OldcY As Single
Dim NewX As Single
Dim NewY As Single
Private Sub RollDice()
Dim Die(1 To 2) As String
Dim Chance As Integer
Dim i As Integer
Roll.Visible = False
Do Until i > 20 + Chance 'chance will be 0 until first roll
DoEvents
Randomize
Chance = Int((Rnd * 6) + 1) 'returns random between 1 and 6
Die(1) = App.Path & "\" & CStr(Chance) & ".Gif"
'assumes not in root directory
Image1 = LoadPicture(Die(1))
Chance = Int((Rnd * 6) + 1)
Die(2) = App.Path & "\" & CStr(Chance) & ".Gif"
Image2 = LoadPicture(Die(2))
i = i + 1
Loop
Roll.Visible = True
End Sub
Private Function Resize() As Single
Select Case Screen.Width  'resize for resoulution
Case 9600
Resize = 1
Case 12000
Resize = 1.25
Case 15360
Resize = 1.6
Case 19200
Resize = 2
Case Else
Resize = 1
End Select
End Function



Private Sub Form_Load()
Width = Screen.Width * 0.75
Height = Screen.Height * 0.75
Move Width / 6, Height / 6
Label1.FontSize = Label1.FontSize * Resize
Label2.FontSize = Label2.FontSize * Resize
Roll.FontSize = Roll.FontSize * Resize
Image2.Width = Image2.Width * Resize
Image2.Height = Image2.Height * Resize
Image1.Width = Image1.Width * Resize
Image1.Height = Image1.Height * Resize
Roll.Width = Roll.Width * Resize
Roll.Height = Roll.Height * Resize
Label1.Move 0, 0
Label2.Move Width - Label2.Width, ScaleHeight - Label2.Height
Image1.Move Width / 2 - Image1.Width - Roll.Width / 2, Height / 3
Image2.Move Width / 2 + Roll.Width / 2, Height / 3
Roll.Move Width / 2 - Roll.Width / 2, Height / 3 + Roll.Height / 2
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldcX = Image1.Left 'record old location
OldcY = Image1.Top
OldX = X 'capture X and Y for mousemove
OldY = Y
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim NewX As Single
'Dim NewY As Single
If Button = 1 Then
NewX = X - OldX 'make control follow mouse
NewY = Y - OldY
Roll.Move Roll.Left + NewX, Roll.Top + NewY
Image1.Move Image1.Left + NewX, Image1.Top + NewY
Image2.Move Image2.Left + NewX, Image2.Top + NewY
End If
End Sub




Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldcX = Image2.Left 'record old location
OldcY = Image2.Top
OldX = X 'capture X and Y for mousemove
OldY = Y
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
NewX = X - OldX 'make control follow mouse
NewY = Y - OldY
Roll.Move Roll.Left + NewX, Roll.Top + NewY
Image1.Move Image1.Left + NewX, Image1.Top + NewY
Image2.Move Image2.Left + NewX, Image2.Top + NewY
End If
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldcX = Label1.Left 'record old location
OldcY = Label1.Top
OldX = X 'capture X and Y for mousemove
OldY = Y
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
NewX = X - OldX 'make control follow mouse
NewY = Y - OldY
Label1.Move Label1.Left + NewX, Label1.Top + NewY
End If
End Sub


Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Left < Label2.Left - Label1.Width Or _
Label1.Left > Label2.Left + Label1.Width Or _
Label1.Top > Label2.Top + Label1.Height Or _
Label1.Top < Label2.Top - Label1.Height Then
Label1.Move OldcX, OldcY 'note:
'If you try to use the OldX and OldY here the
'value will not update next move
Else
Label1.Move Label2.Left + Width * 0.025, Label2.Top + Height * 0.06
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldcX = Label1.Left 'record old location
OldcY = Label1.Top
OldX = X 'capture X and Y for mousemove
OldY = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
NewX = X - OldX 'make control follow mouse
NewY = Y - OldY
Label2.Move Label2.Left + NewX, Label2.Top + NewY
End If
End Sub


Private Sub Roll_Click()
RollDice
End Sub


Private Sub Roll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldcX = Roll.Left 'record old location
OldcY = Roll.Top
OldX = X 'capture X and Y for mousemove
OldY = Y
End Sub


Private Sub Roll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
NewX = X - OldX 'make control follow mouse
NewY = Y - OldY
Roll.Move Roll.Left + NewX, Roll.Top + NewY
Image1.Move Image1.Left + NewX, Image1.Top + NewY
Image2.Move Image2.Left + NewX, Image2.Top + NewY
End If
End Sub


