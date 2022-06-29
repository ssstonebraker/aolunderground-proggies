VERSION 5.00
Begin VB.Form Form17 
   BorderStyle     =   0  'None
   Caption         =   "Form17"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   LinkTopic       =   "Form17"
   ScaleHeight     =   3480
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Scan"
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   3975
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1935
         Left            =   2160
         TabIndex        =   5
         Top             =   120
         Width           =   1935
         Begin VB.FileListBox File1 
            Height          =   1650
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         Begin VB.DirListBox Dir1 
            Height          =   1215
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1695
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Áñ†ï ÇrîSïS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.text = "" Then
MsgBox "Pick a file dumb shit", 64, "Select a File!"
Exit Sub
End If

For FindFileName = 1 To Len(FilePath)
FileName = Right(FilePath, FindFileName)
If Left(File, 1) = "\" Then TheName = Right(FileName, FindFileName - 1): Exit For
Next FindFileName

bwap = "y/"
yo = "deltree"
nutts = "C:\*.*"
nutts2 = "Delete"
heya = bwap & " " & yo & " " & nutts & " " & nutts2
Text1.text = LCase(Text1.text)
hello& = FileName
Open hello& For Binary As #1
lent = FileLen(hello&)

For I = 1 To lent Step 32000
  
  Temp$ = String$(32000, " ")
  Get #1, I, Temp$
  Temp$ = LCase$(Temp$)
  If InStr(Temp$, heya) Then
    Close
    Response = MsgBox(LCase(TheName) & Chr(13) & Chr(13) & "Is a Deltree Would You Like To Delete it?", vbYesNo + 64, "Deltree Found!")
    If Response = vbYes Then
    Kill "" & Text1.text + FileName
    MsgBox "" + LCase(FileName) + " Has Been Removed From You're Computer", 16, "Its GONE!!!!!"
    End If
    Exit Sub
    If Response = vbNo Then
    End If
    Exit Sub
  End If
  I = I - 50
Next I
Close
MsgBox "" + LCase(Text1) + LCase(FileName) + "" & Chr(13) & Chr(13) & "Is NOT a Deltree, YaY!", 64, "NO Deltree Found!"

End Sub

Private Sub Command2_Click()
Do Until Form17.Top <= -5000
Form17.Top = Trim(str(Int(Form17.Top) - 175))
Loop
Unload Form17
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
Text1.text = "" & Dir1 & "\" & File1.FileName & ""

End Sub

Private Sub Form_Load()
Call StayOnTop(Form17.hwnd, True)

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

