VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Crystal  Dawn  Get Line Example"
   ClientHeight    =   4350
   ClientLeft      =   2085
   ClientTop       =   1230
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "First Visible Line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   675
      Width           =   7530
      Begin VB.CommandButton cmbGetText 
         Caption         =   "Get Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox txtFirstVisibleLine 
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   4830
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Number of Lines"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5760
      TabIndex        =   5
      Top             =   15
      Width           =   1815
      Begin VB.TextBox txtNumberofLines 
         Height          =   285
         Left            =   405
         TabIndex        =   6
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Multi-Line Text Box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   60
      TabIndex        =   3
      Top             =   1545
      Width           =   7530
      Begin VB.TextBox txtMultiLine 
         Height          =   2280
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   7395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Retrieve Single Line of Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   5655
      Begin VB.TextBox txtLineNumber 
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtLineFromFile 
         Height          =   285
         Left            =   645
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Label lblListenHereZealots 
      Alignment       =   2  'Center
      Caption         =   "Coded By Naïve For Crystal Dawn. Give Credit."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   10
      Top             =   4140
      Width           =   7545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_sLineString As String * 1056
Dim m_lngRet As Long
Dim m_sRetString As String
Private Sub cbLoadFile_Click()
    '
    Dim iHandle As Integer
    Dim iFileString As String
    Dim iLine As String
    '
    iHandle = FreeFile
    '
    Open "c:\vb6ht\chapter01\chp1.txt" For Input As #iHandle
    Do While Not EOF(iHandle)
        Line Input #iHandle, sLine
        iFileString = iFileString + sLine + vbCr + vbLf
    Loop
    Close #iHandle
    txtMultiLine.Text = iFileString
    '
    Call txtLineNumber_Change
    CountLines
    UpDown1.Min = 1
    UpDown1.Max = m_lngRet
    UpDown1.Value = 1
    '
End Sub
Private Function ReadLine(iLine As Integer) As String
    Dim m_intRet
    m_sLineString = Space$(1056)
    '
    m_intRet = SendMessage(txtMultiLine.hwnd, EM_GETLINE, iLine, ByVal m_sLineString)
    ReadLine = m_sLineString
    '
End Function
Private Sub CountLines()
    '
    m_lngRet = SendMessage(txtMultiLine.hwnd, EM_GETLINECOUNT, 0, 0)
    txtNumberofLines.Text = Str$(m_lngRet)
    '
End Sub
Private Sub cmbUpdate_Click()
 

End Sub

Private Sub cmbGetText_Click()
   CountLines
    '
    txtFirstVisibleLine.Text = ReadLine(ReadFirstVisibleLine())
    '
    txtLineFromFile.Text = ReadLine(Val(txtLineNumber.Text) - 1)
End Sub


Private Sub txtLineNumber_Change()
On Error GoTo Filth
    txtLineFromFile.Text = ReadLine(Val(txtLineNumber.Text) - 1)
    Exit Sub
Filth:
' This error is reached when  too large a value is entered
' to prevent an ovrflow
txtLineNumber.Text = 1
End Sub
Private Function ReadFirstVisibleLine() As Integer
    '
    ReadFirstVisibleLine = SendMessage(txtMultiLine.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
    '
End Function

Private Sub UpDown1_Change()
    txtLineNumber.Text = Str$(UpDown1.Value)
End Sub

Private Sub VScroll1_Change()

End Sub


Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub VScroll1_Scroll()

End Sub


Private Sub txtMultiLine_Change()
CountLines
End Sub


