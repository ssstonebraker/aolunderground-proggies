VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zricks High Score Table"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmScore.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Scores"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame fraMasters 
      Caption         =   "Zricks Masters..."
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   8
         Left            =   4800
         TabIndex        =   20
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   7
         Left            =   4800
         TabIndex        =   19
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   18
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   5
         Left            =   4800
         TabIndex        =   17
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   4
         Left            =   4800
         TabIndex        =   16
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   15
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   14
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   13
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   2
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Name)"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' -------------------------------------------------------
' frmScore. This Form is dependent upon the module
' modScore.BAS and the form frmNewScore.
' -------------------------------------------------------

Private Sub cmdClear_Click()
Dim i As Long

    If MsgBox("Are you sure you want to clear the High Score Table?", vbYesNo + vbQuestion + vbDefaultButton2, "High Score Table") = vbYes Then
        For i = 1 To MAX_HISCORES
            reg.DeleteSetting SECTION, ENTRY & i
        Next i
        GetScores
        Form_Load
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long

    If Num_HiScores > 0 Then
        ' Show the High Scores.
        For i = 1 To Num_HiScores
            lblName(i).Caption = Hi(i).pName
            lblScore(i).Caption = Hi(i).lngScore
        Next i
    End If
End Sub
