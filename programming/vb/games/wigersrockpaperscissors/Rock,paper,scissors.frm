VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   Caption         =   "Wiger"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "End"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Clear"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Start"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame status 
      BackColor       =   &H00FFFF00&
      Caption         =   "Status"
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   1095
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame ichose 
      BackColor       =   &H00FFFF00&
      Caption         =   "What I chose"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "What you chose"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      Begin VB.Label theychose 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Wiger's Rock, Paper, Scissors Game"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'this sets the variable guess as a string...guess
'is set as the variable of your input answer
'i is set as an integer so that the computer will choose
'randomly all the "cases" known as i
    
    Dim Guess As String, i As Integer
    Do
        Guess$ = LCase(InputBox("Choose rock, paper, or scissors...", "Choose"))
        Select Case Guess$
            Case "rock"
                i = 1
            Case "paper"
                i = 2
            Case "scissors"
                i = 3
        End Select
    Loop Until i <> 0
    theychose.Caption = Guess$
    Dim x As Integer
    x = Int(Rnd * 3) + 1
    Select Case x
        Case 1
            Label2.Caption = "I chose rock. "
            Select Case i
                Case 1
                    Label3.Caption = Label3.Caption & "We tied!"
                Case 2
                    Label3.Caption = Label3.Caption & "You won! :("
                Case 3
                    Label3.Caption = Label3.Caption & "I won! :)"
            End Select
        Case 2
            Label2.Caption = "I chose paper. "
            Select Case i
                Case 2
                    Label3.Caption = Label3.Caption & "We tied!"
                Case 3
                   Label3.Caption = Label3.Caption & "You won! :("
                Case 1
                    Label3.Caption = Label3.Caption & "I won! :)"
            End Select
        Case 3
            Label2.Caption = "I chose scissors. "
            Select Case i
                Case 3
                    Label3.Caption = Label3.Caption & "We tied!"
                Case 1
                    Label3.Caption = Label3.Caption & "You won! :("
                Case 2
                    Label3.Caption = Label3.Caption & "I won! :)"
            End Select
    End Select
End Sub

Private Sub Command2_Click()
'this just clears all the labels so the user can restart
'another game
Label3.Caption = ""
Label2.Caption = ""
theychose.Caption = ""
End Sub

Private Sub Form_Load()

End Sub
