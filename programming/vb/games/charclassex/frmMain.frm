VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEnemy 
      Height          =   1035
      Left            =   3780
      TabIndex        =   8
      Top             =   450
      Width           =   1605
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fight Enemy"
      Height          =   375
      Left            =   6090
      TabIndex        =   7
      Top             =   1620
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ShowList"
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   390
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   2040
      Width           =   7035
   End
   Begin VB.ListBox lstCurrent 
      Height          =   1230
      Left            =   1920
      TabIndex        =   3
      Top             =   420
      Width           =   1635
   End
   Begin VB.ListBox lstTotalChar 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      ItemData        =   "frmMain.frx":0000
      Left            =   180
      List            =   "frmMain.frx":0028
      TabIndex        =   0
      Top             =   390
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "Enemy to Fight:"
      Height          =   345
      Left            =   3780
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Click on name for Status Of Char or Enemy"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   1680
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Members in party"
      Height          =   255
      Left            =   1950
      TabIndex        =   2
      Top             =   180
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "All Available Members"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'RPG Example
'Programed by Dexter
'Email: Dexter@aol.com
'
'Started on: November 29, 2001
'Version 1 Completed on: December 09, 2001
'
'

Private Sub Command1_Click()
  Form2.Show
End Sub

Private Sub Command2_Click()
  Dim i As Integer
  i = Rnd * 5
  lstEnemy.Clear
  Call lstEnemy.AddItem(enemyList(i).GetName)
End Sub

Private Sub lstCurrent_Click()
  Dim which As Integer
  Dim char As Integer
    which = lstCurrent.ListIndex
    Select Case which
      Case 0
        char = char1
      Case 1
        char = char2
      Case 2
        char = char3
      Case 3
        char = char4
      Case 4
        char = char5
    End Select
    Call PrintInfo(char)
End Sub

Private Sub lstTotalChar_Click()
    Call PrintCharInfo(lstTotalChar.ListIndex)

End Sub

