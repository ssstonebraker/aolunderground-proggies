VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Justins Calculator"
   ClientHeight    =   1935
   ClientLeft      =   6735
   ClientTop       =   1950
   ClientWidth     =   2130
   Icon            =   "Calculater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2130
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      Height          =   360
      Left            =   1080
      TabIndex        =   17
      Top             =   480
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      Height          =   360
      Index           =   4
      Left            =   1200
      TabIndex        =   16
      Top             =   1560
      Width           =   360
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      Height          =   360
      Left            =   840
      TabIndex        =   15
      Top             =   1560
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   1560
      Width           =   840
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   840
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      Height          =   360
      Index           =   2
      Left            =   1080
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   360
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   360
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   360
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   360
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   360
      Index           =   3
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   720
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   465
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   360
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   360
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   345
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   360
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   360
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "ce/c"
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   360
      Index           =   9
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   360
      Index           =   8
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   360
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   360
      Index           =   7
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Op1, Op2
Dim DecimalFlag As Integer
Dim NumOps As Integer
Dim LastInput
Dim OpFlag
Dim TempReadout

Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")

End Sub

Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "X"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Justins Calculator"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub


