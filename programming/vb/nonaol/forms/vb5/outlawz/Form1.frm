VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OutLawz MP3 Player 1.0 By Mo0NiE"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command10 
         Caption         =   "Exit"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000D&
         Caption         =   "Random"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00FF0000&
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Remove"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Min"
         Height          =   195
         Left            =   3360
         TabIndex        =   12
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Load List"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save List"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear List"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "Rnd Play"
         Height          =   195
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pause"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   195
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000001&
         Height          =   1260
         ItemData        =   "Form1.frx":0442
         Left            =   120
         List            =   "Form1.frx":0444
         TabIndex        =   3
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   4095
      End
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   360
      Width           =   1200
   End
   Begin VB.PictureBox MediaPlayer1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then 'random?
Command4.Caption = "Rnd Play"
Else
Command4.Caption = "Play" 'it wouldn't be random if random was off.
End If

End Sub

Private Sub Command1_Click()
On Error GoTo err
Dim str As String, i As Integer


    For i = 0 To Rnd
    On Error GoTo err
          str = CommonDialog1 _
          & CStr(i%)
    Next i%
    '^^^converts i to a string^^^
    List1.AddItem CommonDialog1 'Adds name to list
    On Error GoTo err
 'the label1 updater
On Error GoTo err
err: 'define err
Exit Sub
End Sub

Private Sub Command2_Click()
 MediaPlayer1.Stop 'stop
End Sub

Private Sub Command3_Click()
If List1.ListCount = 0 Then Exit Sub
If Command3.Caption = "Pause" Then

Command3.Caption = "Resume"
'scroll ("
Else

Command3.Caption = "Pause"
End If
End Sub

Private Sub Command4_Click()
playGo
End Sub

Private Sub Command5_Click()
List1.Clear 'chuck out list
listdoup
End Sub

Private Sub Command6_Click()
Dim i As Variant
CommonDialog1.Filter = "Sonique Playlist|*.PLS"
CommonDialog1.DialogTitle = "Save As Sonique Play List"
CommonDialog1.ShowSave
Kill (CommonDialog1.filename) 'deletes odd file, other wise it would add it onto the bottom of the old file.
For i = 0 To List1.ListCount
    WritePrivateProfileString _
        "playlist", ("File" & i + 1), _
        List1.List(i), CommonDialog1.filename
        'Go though list adding it to the file.
        Next i
    WritePrivateProfileString _
        "playlist", "NumberOfEntries", _
        List1.ListCount, CommonDialog1.filename
        'Sonique uses this to see how many files are in the list.

End Sub

Private Sub Command7_Click()
Dim buf As String * 256, a_line As String, length As Long, numfiles As String, file, fnum As Integer, lines As Integer, boo, i As Variant
CommonDialog1.Filter = "Sonique Playlist (*.PLS)|*.PLS|Winamp Playlist (*.M3U)|*.M3U" 'winamp or sonique file
CommonDialog1.DialogTitle = "Select a List to Load"
CommonDialog1.ShowOpen 'show open
boo = Right(CommonDialog1.filename, 3) 'my way of reading the filetype
On Error GoTo err
If boo = "m3u" Then 'if its a winamp file then...
List1.Clear 'empty list, this stop colisions.
    Dim strFileName As String, strText As String, strFilter As String, strBuffer As String, FileHandle%
        strFileName = CommonDialog1.filename
        FileHandle% = FreeFile
        Open strFileName For Input As #FileHandle%
        Do While Not EOF(FileHandle%) 'If its not the end of the file, keep on going!
            
            Line Input #FileHandle%, strBuffer
            List1.AddItem (strBuffer) 'add item
            strText = strText & strBuffer & vbCrLf
        Loop 'untill EOF, End Of File
        For i = 0 To List1.ListCount
        boo = Right(List1.List(i), 3) 'filetype again
        If boo = "mp3" Then              '
        ElseIf i < List1.ListCount Then  'Makes sure that they are only mp3 files.
        List1.RemoveItem (i)             'if not, chuck it out!
        List1.Refresh                    '
        End If                           '
        Next i                           '
        List1.RemoveItem (0)             '
        Close #FileHandle%               ' Clears the file out of the memory
        
        ElseIf boo = "PLS" Then 'but if its a sonique file
    List1.Clear
    fnum = FreeFile
    On Error GoTo err
    Open CommonDialog1.filename For Input As fnum
    Do While Not EOF(fnum)
    On Error GoTo err
        Line Input #fnum, a_line
        lines = lines + 1
    Loop 'untill EOF
    Close fnum
    On Error GoTo err
    numfiles = GetPrivateProfileString( _
        "playlist", "NumberOfEntries", "", _
        buf, Len(buf), CommonDialog1.filename)
        'read each item in the file
        On Error GoTo err
Do Until List1.ListCount = lines - 3
    file = "File" & List1.ListCount + 1
        length = GetPrivateProfileString( _
        "playlist", file, "", _
        buf, Len(buf), CommonDialog1.filename)
    List1.AddItem Left$(buf, length) ' add item

Loop
End If
listdoup 'the label1 updater
On Error GoTo err
err:
Exit Sub
End Sub

Private Sub Command8_Click()
If Me.Height = 5340 Then 'resize me
Me.Height = 1050
List1.Visible = False 'hide list
Command6.Top = 480
Command6.Left = 1560
Command6.Caption = "Max" 'rename button
Else
Me.Height = 5340 'do it all the other way around
List1.Visible = True
Command6.Top = 4800
Command6.Left = 120
Command6.Caption = "Min"
End If
End Sub

Private Sub Command9_Click()
List1.RemoveItem (List1.ListIndex) 'deletes the selected item

End Sub

Private Sub Form_Load()

    Dim howlong, n As Integer, c As String
    c = Command 'used if user opens a file WITH the program
    n = 1
         
End Sub

Private Sub List1_Click()
 'stops current track
MediaPlayer1.filename = List1.Text 'plays selected file
Scroll (MediaPlayer1.filename)
End Sub

Private Sub Timer1_Timer()
Picture2.Top = Picture2.Top - 1 'move
If Picture2.Top <= -Picture2.Height Then Picture2.Top = Picture1.ScaleHeight 'makes it start over

End Sub
