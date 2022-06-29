VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Greets Example"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "kool end'n 2"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "kool end'n"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "example 3"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "example 2"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5520
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "example 1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "example"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim TextGrow As Integer
    Timer1.interval = 15
    Label1.Caption = "Greets to"
    For TextGrow = 4 To 42
        Label1.FontSize = TextGrow
        Timer1.Enabled = True
        Do
            DoEvents
            DoEvents
        Loop Until Timer1.Enabled = False
    Next TextGrow
    Timer1.interval = 15
    Label1.Caption = "Sonic"
    For TextGrow = 4 To 42
        Label1.FontSize = TextGrow
        Timer1.Enabled = True
        Do
            DoEvents
            DoEvents
        Loop Until Timer1.Enabled = False
    Next TextGrow
    Timer1.interval = 15
    Label1.Caption = "For"
    For TextGrow = 4 To 42
        Label1.FontSize = TextGrow
        Timer1.Enabled = True
        Do
            DoEvents
            DoEvents
        Loop Until Timer1.Enabled = False
    Next TextGrow
    Timer1.interval = 15
    Label1.Caption = "help'n"
    For TextGrow = 4 To 42
        Label1.FontSize = TextGrow
        Timer1.Enabled = True
        Do
            DoEvents
            DoEvents
        Loop Until Timer1.Enabled = False
    Next TextGrow
End Sub

Private Sub Command2_Click()
    Dim RandColor As Integer, Rand As Integer
    Dim RGBRED As Integer, RGBGREEN As Integer
    Dim RGBBLUE As Integer, PauseNow
    Label1.FontSize = 36
    Label1.Caption = "greets to"
    For RandColor = 1 To 50
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBRED = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBGREEN = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBBLUE = Rand
        Label1.ForeColor = RGB(RGBRED, RGBGREEN, RGBBLUE)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.001)
            DoEvents
        Wend
    Next RandColor
    Label1.Caption = "sonic"
    For RandColor = 1 To 50
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBRED = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBGREEN = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBBLUE = Rand
        Label1.ForeColor = RGB(RGBRED, RGBGREEN, RGBBLUE)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.001)
            DoEvents
        Wend
    Next RandColor
    Label1.Caption = "for"
    For RandColor = 1 To 50
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBRED = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBGREEN = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBBLUE = Rand
        Label1.ForeColor = RGB(RGBRED, RGBGREEN, RGBBLUE)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.001)
            DoEvents
        Wend
    Next RandColor
    Label1.Caption = "help'n"
    For RandColor = 1 To 50
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBRED = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBGREEN = Rand
        Randomize
        Rand = Int((Val(250) * Rnd) + 1)
        RGBBLUE = Rand
        Label1.ForeColor = RGB(RGBRED, RGBGREEN, RGBBLUE)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.001)
            DoEvents
        Wend
    Next RandColor
End Sub

Private Sub Command3_Click()
    Dim InOut As String, MakeInOut As Integer
    Dim PauseNow
    InOut = "greets to"
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, Len(InOut) - MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    InOut = "sonic"
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, Len(InOut) - MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    InOut = "for"
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, Len(InOut) - MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    InOut = "help'n"
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, Len(InOut) - MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    InOut = "sonicz visual basic guide/help"
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    For MakeInOut = 1 To Len(InOut)
        Label1.Caption = Left(InOut, Len(InOut) - MakeInOut)
        PauseNow = Timer
        While Timer - Val(PauseNow) < Val(0.1)
            DoEvents
        Wend
    Next MakeInOut
    
End Sub

Private Sub Command4_Click()
    Dim EveryWhere As Integer, FormLeft As Integer
    Dim FormTop As Integer
    For EveryWhere = 1 To 35
        Randomize
        DoEvents
        FormLeft = Int((Val(Screen.Width - 200) * Rnd) + 1)
        Randomize
        DoEvents
        FormTop = Int((Val(Screen.Width - 200) * Rnd) + 1)
        DoEvents
        Me.Left = FormLeft
        Me.Top = FormTop
        DoEvents
    Next EveryWhere
    End
End Sub

Private Sub Command5_Click()
    Do
        DoEvents
        DoEvents
        Me.Height = Me.Height - 90
        DoEvents
        DoEvents
    Loop Until Me.Height < 500
    Do
        DoEvents
        DoEvents
        Me.Left = Me.Left - 400
        DoEvents
        DoEvents
    Loop Until Me.Left + Me.Width < 0
    End
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub
