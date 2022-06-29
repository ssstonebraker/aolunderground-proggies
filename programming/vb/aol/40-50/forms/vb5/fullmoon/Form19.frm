VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form19 
   BorderStyle     =   0  'None
   Caption         =   "Phish Manager"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form19"
   Picture         =   "Form19.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Load Pw's"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Load Sn's"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phish Manager"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Pw's"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " _"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Sn's"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
List1.AddItem Text1
Text1 = ""
End Sub

Private Sub Form_Load()
StayOnTop Me

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
 Dim l003C As Variant
Dim l0040 As Variant
Dim l0044 As String
Dim l0046 As String

On Error Resume Next
CommonDialog1.FLAGS = &H4& Or &H2& Or &H800&
CommonDialog1.DefaultExt = "MMN"
CommonDialog1.DialogTitle = "Save Names as ..."
CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
CommonDialog1.MaxFileSize = 2000
CommonDialog1.FileName = "*.txt"
CommonDialog1.CancelError = True
CommonDialog1.Action = 2
Select Case Err
    Case 0:
    l003C = List1.ListCount - 1
    For l0040 = 0 To l003C - 1
    l0044$ = List1.List(l0040)
    l0046$ = l0046$ + l0044$ & ","
    Next l0040
    l0046$ = l0046$ + List1.List(l003C)
        Open CommonDialog1.FileName For Output As #1
            Print #1, l0046$
        Close #1
    Case 32755:
        Err = False
    Case Else:
        MsgBox "Unexpected error:" & Str(Err) & Chr$(13) & "Error message: " & Error$(Err), 16, "Unexpected Error!!!"
End Select
End Sub

Private Sub Label2_Click()
Form19.WindowState = 1
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
 Dim l003C As Variant
Dim l0040 As Variant
Dim l0044 As String
Dim l0046 As String

On Error Resume Next
CommonDialog1.FLAGS = &H4& Or &H2& Or &H800&
CommonDialog1.DefaultExt = "MMN"
CommonDialog1.DialogTitle = "Save Names as ..."
CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
CommonDialog1.MaxFileSize = 2000
CommonDialog1.FileName = "*.txt"
CommonDialog1.CancelError = True
CommonDialog1.Action = 2
Select Case Err
    Case 0:
    l003C = List1.ListCount - 1
    For l0040 = 0 To l003C - 1
    l0044$ = List1.List(l0040)
    l0046$ = l0046$ + l0044$ & ","
    Next l0040
    l0046$ = l0046$ + List1.List(l003C)
        Open CommonDialog1.FileName For Output As #1
            Print #1, l0046$
        Close #1
    Case 32755:
        Err = False
    Case Else:
        MsgBox "Unexpected error:" & Str(Err) & Chr$(13) & "Error message: " & Error$(Err), 16, "Unexpected Error!!!"
End Select
End Sub

Private Sub lstList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label6_Click()
 Dim lIndex  As Long
    
    '-- Get a text file name to open
    CommonDialog1.DialogTitle = "Open Text File"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Text (*.txt)"
    CommonDialog1.FileName = "*.txt"
    On Error Resume Next
    CommonDialog1.Action = 1
    If Err Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    '-- Load the file
    objTextFile.Load (CommonDialog1.FileName)
    
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

  

    Screen.MousePointer = vbNormal

End Sub

Private Sub Label7_Click()
 Dim lIndex  As Long
    
    '-- Get a text file name to open
    CommonDialog1.DialogTitle = "Open Text File"
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Text (*.txt)"
    CommonDialog1.FileName = "*.txt"
    On Error Resume Next
    CommonDialog1.Action = 1
    If Err Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    '-- Load the file
    objTextFile.Load (CommonDialog1.FileName)
    
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
