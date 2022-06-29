VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Example Deltree Scanner By Evil Cl0wn"
   ClientHeight    =   2460
   ClientLeft      =   1800
   ClientTop       =   2430
   ClientWidth     =   4560
   Height          =   2865
   Left            =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4560
   Top             =   2085
   Width           =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1425
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1380
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "You Have to Select a File to Scan First!", 64, "Select a File!"
Exit Sub
End If

For FindFileName = 1 To Len(FilePath)
filename = Right(FilePath, FindFileName)
If Left(File, 1) = "\" Then TheName = Right(filename, FindFileName - 1): Exit For
Next FindFileName

bwap = "y/"
yo = "deltree"
nutts = "C:\*.*"
nutts2 = "Delete"
heya = bwap & " " & yo & " " & nutts & " " & nutts2
Text1.Text = LCase(Text1.Text)
hello& = filename
Open hello& For Binary As #1
lent = FileLen(hello&)

For i = 1 To lent Step 32000
  
  Temp$ = String$(32000, " ")
  Get #1, i, Temp$
  Temp$ = LCase$(Temp$)
  If InStr(Temp$, heya) Then
    Close
    Response = MsgBox(LCase(TheName) & Chr(13) & Chr(13) & "Is a Deltree Would You Like To Delete it?", vbYesNo + 64, "Deltree Found!")
    If Response = vbYes Then
    Kill "" & Text1.Text + filename
    MsgBox "" + LCase(filename) + " Has Been Removed From You're Computer", 16, "Its GONE!!!!!"
    End If
    Exit Sub
    If Response = vbNo Then
    End If
    Exit Sub
  End If
  i = i - 50
Next i
Close
MsgBox "" + LCase(Text1) + LCase(filename) + "" & Chr(13) & Chr(13) & "Is NOT a Deltree, YaY!", 64, "NO Deltree Found!"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Text1.Text = "" & Dir1 & "\" & File1.filename & ""
End Sub

Private Sub Form_Load()
MsgBox "If All You Did was Change this Form Around and called it your own then you should be a shamed to call your self a programmer this is for learning and not for you to use as your own!,         If You Would Like to Say Sumthin to me Please E-Mail Me at     KoRn50420@juno.com, Piece       -Evil Cl0wn", vbInformation, "U DIDNT MAKE THIS EVIL CL0WN DID =)"
End Sub


