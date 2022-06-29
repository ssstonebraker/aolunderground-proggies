VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Convert VB6 to VB5"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3720
   Icon            =   "HackFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Text            =   "Retained=0"
      Top             =   1515
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB6 to 5 example bY: STuCCo"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Select file:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   99999
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2010
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1005
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ChunkSize = 4096
'4096 seems to work best on my PC and I am
'sure it will work well on yours too.

Private Sub Command2_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "WARNING! You are about to change the "
    Msg = Msg & "contents of a file. Are you sure want to Do this?"
    Msg = Msg & vbCrLf & vbCrLf & "It is HIGHLY RECOMMENDED that you make "
    Msg = Msg & "a backup copy of this file before continuing!"
    Msg = Msg & vbCrLf & vbCrLf & "Do you want to continue?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "WARNING! File Modification is Ready."
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)

    If Response = vbYes Then ' User chose Yes.
        MyString = "Yes"
    Else
        MyString = "No"
    End If

    If MyString = "Yes" Then
           'ChangeFile "c:\test.txt", "hello", "SeeYa"
        ChangeFile Text1.Text, Text2.Text, Text3.Text
    End If

End Sub

Private Function FindInString(StartPos As Integer, StrToSearch As String, _
    StrToFind As String) As Integer

    If Check1.Value = 0 Then
        FindInString = InStr(StartPos, UCase(StrToSearch), UCase(StrToFind))
    Else
        FindInString = InStr(StartPos, StrToSearch, StrToFind)
    End If

End Function

Public Sub ChangeFile(FName$, IDString$, NString$)
      
    Dim PosString, WhereString
    Dim FileNumber, A$, NewString$
    Dim AString As String * ChunkSize
    Dim IsChanged As Boolean
    Dim BlockIsChanged As Boolean
    Dim NumChanges As Integer
    
    IsChanged = False
    BlockIsChanged = False
    On Error GoTo Problems
    FileNumber = FreeFile
    PosString = 1
    WhereString = 0
    AString = Space$(ChunkSize)
    'Make sure strings have same size

    If Len(IDString$) > Len(NString$) Then
        NewString$ = NString$ + Space$(Len(IDString$) - Len(NString$))
    Else
        NewString$ = Left$(NString$, Len(IDString$))
    End If

    Open FName$ For Binary As FileNumber
    NumChanges = 0

    If LOF(FileNumber) < ChunkSize Then
        A$ = Space$(LOF(FileNumber))
        Get #FileNumber, 1, A$
        WhereString = FindInString(1, A$, IDString$)
    Else
        A$ = Space$(ChunkSize)
        Get #FileNumber, 1, A$
        WhereString = FindInString(1, A$, IDString$)
    End If

    Do
        While WhereString <> 0
            tempstring = Left$(A$, WhereString - 1) & NewString$ & Mid$(A$, WhereString + Len(NewString$))
            A$ = tempstring
            NumChanges = NumChanges + 1
            IsChanged = True
            BlockIsChanged = True
            WhereString = FindInString(WhereString + 1, A$, IDString$)
        Wend

        If BlockIsChanged Then
            Put #FileNumber, PosString, A$
            BlockIsChanged = False
        End If
        
        PosString = ChunkSize + PosString - Len(IDString$)
        ' If we're finished, exit the loop

        If EOF(FileNumber) Or PosString > LOF(FileNumber) Then
            Exit Do
        End If

        ' Get the next chunk to scan

        If PosString + ChunkSize > LOF(FileNumber) Then
            A$ = Space$(LOF(FileNumber) - PosString + 1)
            Get #FileNumber, PosString, A$
            WhereString = FindInString(1, A$, IDString$)
        Else
            A$ = Space$(ChunkSize)
            Get #FileNumber, PosString, A$
            WhereString = FindInString(1, A$, IDString$)
        End If

    Loop Until EOF(FileNumber) Or PosString > LOF(FileNumber)

    Beep

    If IsChanged = True Then
        Msg = Chr$(34) & FName$ & Chr$(34) & " has been modified." & vbCrLf
        Msg = Msg & NumChanges & " occurrence(s) of " & Chr$(34) & IDString$ & _
        Chr$(34) & " replaced With " & Chr$(34) & _
        Left$(NString$, Len(IDString$)) & Chr$(34)
        MsgBox FName$ & " has been converted to a VB5 file!", vbInformation, "File converted to VB5"
    Else
        MsgBox "File has not been converted to VB5.", vbInformation, "Could not convert"
    End If

    Close
    Exit Sub
Problems:
    Close
    MsgBox "An Error has occurred." & vbCrLf & Err.Description, _
    vbExclamation, "Error number " & Err.Number
End Sub

Private Sub Command1_Click()

    Dim Ans As String
    Ans = GetOpenFileNameDLG("File to convert *.vbp|*.vbp", "Please Select a file", "", Me.hwnd)
    If Ans <> "" Then
        Text1.Text = Ans
    End If

End Sub

Private Sub Command3_Click()

    Unload Me
End Sub

Private Sub Form_Load()
Label4.Caption = "VB6 to 5 example bY: STuCCo"
    Label2.Caption = "Search for:"
    Label3.Caption = "Change to:"
    Check1.Caption = "Case-sensitive search?"
    Check1.Value = 1
    Command1.Caption = "..."
    Command2.Caption = "Convert"
    Command3.Caption = "E&xit"
    Command3.Cancel = True
    Text1.Text = ""
    End Sub


Private Sub Text1_Change()
'Label1 = ""
'Label1 = "Select file: " & Text1
End Sub
