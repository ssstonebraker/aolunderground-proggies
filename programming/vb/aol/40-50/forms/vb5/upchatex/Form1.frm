VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "sonics upchat example"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   1755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   30
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "important.. click"
      Height          =   240
      Left            =   30
      TabIndex        =   6
      Top             =   870
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "off"
      Height          =   240
      Left            =   1095
      TabIndex        =   5
      Top             =   555
      Width           =   585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "on"
      Height          =   240
      Left            =   1095
      TabIndex        =   4
      Top             =   330
      Width           =   585
   End
   Begin VB.Label lblMin 
      Caption         =   "min left:"
      Height          =   210
      Left            =   75
      TabIndex        =   3
      Top             =   570
      Width           =   990
   End
   Begin VB.Label lblPercent 
      Caption         =   "percent:"
      Height          =   210
      Left            =   75
      TabIndex        =   2
      Top             =   330
      Width           =   990
   End
   Begin VB.Label lblName 
      Caption         =   "file: --"
      Height          =   210
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call Upchat
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Call UnUpchat
    Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
    Call MsgBox("sup, this is sonic! if you use sonicUpload.bas, please give me credit.  i couldn't have made it any easier to make an upchat similar to 'up dat encode', this is like a free ride, so please give me just a little credit!.", vbInformation, "credit")
End Sub

Private Sub Timer1_Timer()
    Dim gStats As SonicUPLOADSTAT
    gStats = getUpload
    lblName.Caption = "file: " + gStats.UL_FILENAME
    lblPercent.Caption = "percent: " + gStats.UL_PERDONE
    lblMin.Caption = "min left: " + gStats.UL_MINLEFT
End Sub
