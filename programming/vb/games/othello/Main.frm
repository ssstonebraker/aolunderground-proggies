VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Othello 1.0"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6240
      Top             =   5280
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5520
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Config"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Next Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label status 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label swhite 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblBlack 
      Alignment       =   2  'Center
      Caption         =   "Black"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label sblack 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblWhite 
      Alignment       =   2  'Center
      Caption         =   "White"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   36
      Left            =   2040
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   46
      Left            =   3240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      Caption         =   "PUSH START"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   64
      Left            =   4440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   63
      Left            =   3840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   62
      Left            =   3240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   61
      Left            =   2640
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   60
      Left            =   2040
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   59
      Left            =   1440
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   58
      Left            =   840
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   57
      Left            =   240
      Top             =   4920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   56
      Left            =   4440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   55
      Left            =   3840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   54
      Left            =   3240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   53
      Left            =   2640
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   52
      Left            =   2040
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   51
      Left            =   1440
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   50
      Left            =   840
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   49
      Left            =   240
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   48
      Left            =   4440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   47
      Left            =   3840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   45
      Left            =   2640
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   44
      Left            =   2040
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   43
      Left            =   1440
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   42
      Left            =   840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   41
      Left            =   240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   40
      Left            =   4440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   39
      Left            =   3840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   38
      Left            =   3240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   37
      Left            =   2640
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   35
      Left            =   1440
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   34
      Left            =   840
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   33
      Left            =   240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   32
      Left            =   4440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   31
      Left            =   3840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   30
      Left            =   3240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   29
      Left            =   2640
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   28
      Left            =   2040
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   27
      Left            =   1440
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   26
      Left            =   840
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   25
      Left            =   240
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   24
      Left            =   4440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   23
      Left            =   3840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   22
      Left            =   3240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   21
      Left            =   2640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   20
      Left            =   2040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   19
      Left            =   1440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   18
      Left            =   840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   17
      Left            =   240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   16
      Left            =   4440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   15
      Left            =   3840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   14
      Left            =   3240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   13
      Left            =   2640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   12
      Left            =   2040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   11
      Left            =   1440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   10
      Left            =   840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   9
      Left            =   240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   8
      Left            =   4440
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   7
      Left            =   3840
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   6
      Left            =   3240
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   5
      Left            =   2640
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   4
      Left            =   2040
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   1440
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   840
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bjoueur As Boolean
Dim tmppicture(255) As Integer
Dim TmpFindOK As Boolean
Dim TmpFindTmp As Boolean
Dim tmpNextPlayer As Boolean
Dim Start As Boolean
Private Sub Pts()
bscore = 0
wscore = 0
    For X = 1 To 64
        If tmppicture(X) = 1 Then bscore = bscore + 1
        If tmppicture(X) = 2 Then wscore = wscore + 1
    Next
    sblack.Caption = bscore
    swhite.Caption = wscore
End Sub
Private Sub Joute(pos As Integer)
Select Case pos
    Case 9, 17, 25, 33, 41, 49:
        Call FindCenter(pos, -8)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -7)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 1)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 8)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 9)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        If tmpNextPlayer = True Then
            tmpNextPlayer = False
            Bjoueur = Not Bjoueur
            Call Pts
            If Bjoueur = True Then
                lblPlayer.Caption = Form2.txtBlack & " to play"
            Else: lblPlayer.Caption = Form2.txtWhite & " to play"
            End If
        Else: MsgBox "Jeux Impossible"
        End If
    ''=-=-=-=-=-=-=-=-=-=-=-=-=
    Case 16, 24, 32, 40, 48, 56:
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -9)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -8)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -1)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 7)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 8)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        If tmpNextPlayer = True Then
            tmpNextPlayer = False
            Bjoueur = Not Bjoueur
            Call Pts
            If Bjoueur = True Then
                lblPlayer.Caption = Form2.txtBlack & " to play"
            Else: lblPlayer.Caption = Form2.txtWhite & " to play"
            End If
        Else: MsgBox "Jeux Impossible"
        End If
    
    
    
    ''=-=-=-=-=-=-=-=-=-=-=-=-=
    Case Else
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -9)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -8)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -7)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, -1)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 1)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 7)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 8)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        Call FindCenter(pos, 9)
        Call Pitoune(pos)
        TmpFindOK = False
        TmpFindTmp = False
        If tmpNextPlayer = True Then
            Call Pts
            tmpNextPlayer = False
            Bjoueur = Not Bjoueur
            If Bjoueur = True Then
                lblPlayer.Caption = Form2.txtBlack & " to play"
            Else: lblPlayer.Caption = Form2.txtWhite & " to play"
            End If
        Else: MsgBox "Jeux Impossible"
        End If
End Select

End Sub
Private Sub Pitoune(pos)
    If TmpFindOK = True Then
        If Bjoueur = False Then
            Image1(pos).Picture = LoadPicture(App.Path & "\blanc.bmp")
            tmppicture(pos) = 2
        Else
            Image1(pos).Picture = LoadPicture(App.Path & "\noir.bmp")
            tmppicture(pos) = 1
        End If
        tmpNextPlayer = True
    End If
End Sub
Private Sub FindCenter(pos As Integer, Mode As Single)
Dim TmpFind As Integer
TmpFind = pos + Mode
If Bjoueur = True Then
    tmpimage = 2
Else: tmpimage = 1
End If
'7'
If tmppicture(TmpFind) = tmpimage Then
    TmpFindTmp = True
    Call FindCenter(TmpFind, Mode)
Else
    If tmppicture(TmpFind) <> 0 And TmpFindTmp = True Then
        TmpFindOK = True
    End If
End If

    If TmpFindOK = True Then
        If tmpimage = 1 Then
            Image1(TmpFind).Picture = LoadPicture(App.Path & "\blanc.bmp")
            tmppicture(TmpFind) = 2
        Else
            Image1(TmpFind).Picture = LoadPicture(App.Path & "\noir.bmp")
            tmppicture(TmpFind) = 1
        End If
        
    End If




End Sub

Private Sub Command1_Click()
If Start = False Then
    Start = True
    lblPlayer.Caption = Form2.txtWhite & " to play"
    Call SubStart
Else: MsgBox "Game already started!"
End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()

If Start = True Then
    ok = MsgBox("Do you realy want to end the game", vbYesNo, "End!")
    If ok = 6 Then
        Start = False
        
        bscore = 0
        wscore = 0
        
        For X = 1 To 64
        Image1(X).Picture = LoadPicture("")
        Next
        
        For X = 1 To 64
            If tmppicture(X) = 1 Then bscore = bscore + 1
            If tmppicture(X) = 2 Then wscore = wscore + 1
        Next
        If bscore > wscore Then
            lblPlayer.Caption = Form2.txtBlack.Text & " Win!!!!!"
            For X = 1 To bscore
                Image1(X).Picture = LoadPicture(App.Path & "\noir.bmp")
            Next
            For X = bscore + 1 To wscore + bscore
                Image1(X).Picture = LoadPicture(App.Path & "\blanc.bmp")
            Next
        Else
            lblPlayer.Caption = Form2.txtWhite.Text & " Win!!!!!"
            For X = 1 To wscore
                Image1(X).Picture = LoadPicture(App.Path & "\blanc.bmp")
            Next
            For X = wscore + 1 To bscore + wscore
                Image1(X).Picture = LoadPicture(App.Path & "\noir.bmp")
            Next
        End If
        If bscore = wscore Then lblPlayer.Caption = "Draw Game!!!!!"
    End If
End If
End Sub

Private Sub Command4_Click()
ok = MsgBox("It is impossible to make move!", vbYesNo, "Warning!")
If ok = 6 Then
    Bjoueur = Not Bjoueur
    If Bjoueur = True Then
        lblPlayer.Caption = Form2.txtBlack & " to play"
    Else: lblPlayer.Caption = Form2.txtWhite & " to play"
    End If
End If
End Sub
Private Sub SubStart()

TmpFindOK = False
TmpFindTmp = False
tmpNextPlayer = False
For X = 1 To 64
    tmppicture(X) = 0
    Image1(X).Picture = LoadPicture("")
Next
tmppicture(36) = 1
Image1(29).Picture = LoadPicture(App.Path & "\noir.bmp")
tmppicture(37) = 2
Image1(36).Picture = LoadPicture(App.Path & "\noir.bmp")
tmppicture(28) = 2
Image1(28).Picture = LoadPicture(App.Path & "\blanc.bmp")
tmppicture(29) = 1
Image1(37).Picture = LoadPicture(App.Path & "\blanc.bmp")
Bjoueur = False
End Sub
Private Sub Form_Load()
Start = False
Call SubStart
End Sub
Private Sub Loadwinsock()
IIp = 21
Winsock1.Connect "205.237.57.195", IIp
Select Case Winsock1.State
    Case 0: status.Caption = "Closed"
    Case 1: status.Caption = "Open"
    Case 2: status.Caption = "Listening"
    Case 3: status.Caption = "Pending"
    Case 4: status.Caption = "Resolving host"
    Case 5: status.Caption = "Host resolved"
    Case 6: status.Caption = "Connecting"
    Case 7: status.Caption = "Connected"
    Case 8: status.Caption = "Peer Closing"
    Case 9: status.Caption = "Error"
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click(Index As Integer)
    If tmppicture(Index) = 0 And Start = True Then Call Joute(Index)
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

Private Sub Timer1_Timer()
Select Case Winsock1.State
    Case 0: status.Caption = "Closed"
    Case 1: status.Caption = "Open"
    Case 2: status.Caption = "Listening"
    Case 3: status.Caption = "Pending"
    Case 4: status.Caption = "Resolving host"
    Case 5: status.Caption = "Host resolved"
    Case 6: status.Caption = "Connecting"
    Case 7: status.Caption = "Connected"
    Case 8: status.Caption = "Peer Closing"
    Case 9: status.Caption = "Error"
End Select

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim DData As String
Dim Strnick As String
Dim StrName As String
Dim StrEmail As String
Winsock1.GetData DData
'''''' Entree
If Mid(DData, 1, 2) = "*SS" Then Number = Val(Mid(DData, 3, Len(DData)))


End Sub

