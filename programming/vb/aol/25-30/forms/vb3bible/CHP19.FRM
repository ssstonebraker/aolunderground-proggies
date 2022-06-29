VERSION 2.00
Begin Form formMain 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "File System Controls Project"
   ClientHeight    =   3945
   ClientLeft      =   750
   ClientTop       =   1440
   ClientWidth     =   6615
   Height          =   4380
   Icon            =   CHP19.FRX:0000
   Left            =   675
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6615
   Top             =   1080
   Width           =   6765
   Begin CommandButton cmndRefresh 
      Caption         =   "&Refresh"
      Height          =   372
      Left            =   5280
      TabIndex        =   13
      Top             =   600
      Width           =   1212
   End
   Begin CheckBox chekAttrib 
      Caption         =   "System"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin CheckBox chekAttrib 
      Caption         =   "Read Only"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin CheckBox chekAttrib 
      Caption         =   "Normal"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin CheckBox chekAttrib 
      Caption         =   "Hidden"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin CheckBox chekAttrib 
      Caption         =   "Archive"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   492
      Left            =   5280
      TabIndex        =   10
      Top             =   3240
      Width           =   1212
   End
   Begin DirListBox Dir1 
      Height          =   3264
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2292
   End
   Begin FileListBox File1 
      Height          =   2760
      Left            =   2640
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin CommandButton cmndOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   1212
   End
   Begin DriveListBox Drive1 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2292
   End
   Begin TextBox textSelected 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   120
      Width           =   2415
   End
   Begin Label lablBytes 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   252
      Left            =   3720
      TabIndex        =   12
      Top             =   3480
      Width           =   1332
   End
   Begin Label lablCount 
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   2640
      TabIndex        =   11
      Top             =   3480
      Width           =   972
   End
End

Sub chekAttrib_Click (Index As Integer)
     Select Case Index
          Case 0' attrib
               File1.Archive = -chekAttrib(0).Value
          Case 1' hidden
               File1.Hidden = -chekAttrib(1).Value
          Case 2' normal
               File1.Normal = -chekAttrib(2).Value
          Case 3' read only
               File1.ReadOnly = -chekAttrib(3).Value
          Case 4' system
               File1.System = -chekAttrib(4).Value
     End Select
     lablCount.Caption = Format$(File1.ListCount, "##,###") + " Files"
End Sub

Sub cmndExit_Click ()
    End                  'End the program
End Sub

Sub cmndOK_Click ()
    oldFilename$ = File1.FileName
    On Error Resume Next
    File1.FileName = textSelected.Text
End Sub

Sub cmndRefresh_Click ()
     File1.Refresh
End Sub

Sub Dir1_Change ()
    File1.Path = Dir1.Path
    ChDir Dir1.Path
End Sub

Sub Drive1_Change ()
    On Error Resume Next
    Dir1.Path = CurDir$(Drive1.Drive)
    If Err > 0 Then
        a$ = "Drive " + Drive1.Drive + " not ready"
        MsgBox a$, 48
        Drive1.Drive = Dir1.Path
    End If
End Sub

Sub File1_Click ()
    textSelected.Text = File1.FileName
End Sub

Sub File1_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
     For i = 0 To File1.ListCount - 1
          If Right$(File1.Path, 1) = "\" Then
             TempPath$ = File1.Path + "\" + File1.Filename
          Else
             TempPath$ = File1.Path + File1.Filename
          End If
          If File1.Selected(i) Then
               bytes = bytes + FileLen(File1.List(i))
          End If
     Next i
     lablBytes.Caption = Format$(bytes, " ###,###,##0") & " Bytes"
End Sub

Sub File1_PathChange ()
    lablCount.Caption = Format$(File1.ListCount, "##,##0") + " Files"
    textSelected.Text = ParseFileName()
    Dir1.Path = File1.Path
    Drive1.Drive = File1.Path
    lablBytes.Caption = ""
End Sub

Sub File1_PatternChange ()
    lablCount.Caption = Format$(File1.ListCount, "##,###") + " Files"
    textSelected.Text = ParseFileName()
End Sub

Sub Form_Load ()
    Drive1.Drive = CurDir$
    Dir1.Path = CurDir$
    File1.Path = CurDir$
    chekAttrib(0).Value = Abs(File1.Archive)
    chekAttrib(1).Value = Abs(File1.Hidden)
    chekAttrib(2).Value = Abs(File1.Normal)
    chekAttrib(3).Value = Abs(File1.ReadOnly)
    chekAttrib(4).Value = Abs(File1.System)
    textSelected.Text = ParseFileName()
End Sub

Function ParseFileName () As String
    Dim tempFile As String
    Dim tempDir As String

    tempDir = File1.Path
    tempFile = File1.Pattern
    If tempFile = "" Then tempFile = "*.*"
    If Right$(tempDir, 1) <> "\" Then tempFile = "\" + tempFile
    ParseFileName = tempDir + tempFile
End Function

Sub textSelected_GotFocus ()
    textSelected.SelStart = 0
    textSelected.SelLength = Len(textSelected.Text)
End Sub

