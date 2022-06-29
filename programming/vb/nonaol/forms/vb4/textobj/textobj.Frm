VERSION 4.00
Begin VB.Form frmMain 
   Caption         =   "Text File Object Demo"
   ClientHeight    =   6630
   ClientLeft      =   675
   ClientTop       =   375
   ClientWidth     =   9510
   Height          =   7035
   Left            =   615
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   0
   Top             =   30
   Width           =   9630
   Begin VB.CommandButton btnFind 
      Caption         =   "Find &Next"
      Enabled         =   0   'False
      Height          =   510
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   5985
      Width           =   1365
   End
   Begin VB.CommandButton btnFind 
      Caption         =   "Find &First"
      Enabled         =   0   'False
      Height          =   510
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   5985
      Width           =   1365
   End
   Begin VB.CommandButton btnLoadDirect 
      Caption         =   "Load &Direct"
      Height          =   495
      Left            =   1215
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "&Load"
      Height          =   495
      Left            =   90
      TabIndex        =   1
      Top             =   6030
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         name            =   "Courier New"
         charset         =   0
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   5685
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8910
      Top             =   5940
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private objTextFile As New TextFile

Private Sub btnFind_Click(Index As Integer)

    Dim szMatch As String
    Dim nLineNum    As Integer
    
    Select Case Index
        Case 0 '-- FindFirst
    
            szMatch = InputBox("Enter Search Term")
            If Len(szMatch) = 0 Then
                Exit Sub
            End If
                
            '-- Case-Insensitive Search
            nLineNum = objTextFile.FindFirst(szMatch, False)
            
            If nLineNum Then
                List1.TopIndex = nLineNum - 1
                btnFind(1).Enabled = True
            Else
                btnFind(1).Enabled = False
            End If
        Case 1
            '-- Case-Insensitive Search
            nLineNum = objTextFile.FindNext
            
            If nLineNum Then
                List1.TopIndex = nLineNum - 1
                btnFind(1).Enabled = True
            Else
                btnFind(1).Enabled = False
            End If

    End Select
    
End Sub


Private Sub btnLoadDirect_Click()

    Dim lIndex  As Long
    
    '-- Get a text file name to open
    CommonDialog1.DialogTitle = "Open Text File"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Text (*.txt)"
    CommonDialog1.filename = "*.txt"
    On Error Resume Next
    CommonDialog1.Action = 1
    If Err Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    '-- Load the file
    objTextFile.LoadListBox (CommonDialog1.filename), List1
    
    '-- Error loading?
    If objTextFile.ErrorNum Then
        MsgBox objTextFile.ErrorMsg, vbInformation, "TextFile Object Demo"
        Exit Sub
    End If
    
    btnFind(0).Enabled = False
    btnFind(1).Enabled = False
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub btnOpen_Click()
    
    Dim lIndex  As Long
    
    '-- Get a text file name to open
    CommonDialog1.DialogTitle = "Open Text File"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Text (*.txt)"
    CommonDialog1.filename = "*.txt"
    On Error Resume Next
    CommonDialog1.Action = 1
    If Err Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    '-- Load the file
    objTextFile.Load (CommonDialog1.filename)
    
    '-- Error loading?
    If objTextFile.ErrorNum Then
        MsgBox objTextFile.ErrorMsg, vbInformation, "TextFile Object Demo"
        Exit Sub
    End If
    
    '-- Load the file in the list box
    List1.Clear
    On Error Resume Next
    For lIndex = 1 To objTextFile.Lines
        List1.AddItem objTextFile.Line(lIndex)
        If Err Then Stop
    Next

    btnFind(0).Enabled = True

    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Load()

    '-- Center the form
    Move (Screen.Width \ 2) - (Me.Width \ 2), (Screen.Height \ 2) - (Me.Height \ 2)

End Sub


