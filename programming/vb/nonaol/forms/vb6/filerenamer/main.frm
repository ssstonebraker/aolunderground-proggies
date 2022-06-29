VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "spouts mass file renamer"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "H"
      Height          =   285
      Left            =   3960
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "main.frx":0000
      Left            =   1800
      List            =   "main.frx":0031
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   ".extension"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "main.frx":0090
      Left            =   1800
      List            =   "main.frx":00C1
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   ".extension"
      Top             =   1640
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "R"
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   1640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "one file to change"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0%"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo Erhandle
File1.Refresh
If File1.ListCount = 0 Then MsgBox "Error: No file in directory", vbExclamation, "Error"
If Left(Combo1.Text, 1) <> "." Then MsgBox "Error: Leave period at beginning of extension selection", vbExclamation, "Error": Exit Sub
If Len(Combo1.Text) > 5 Then MsgBox "Error: Extension too long", vbExclamation, "Error": Exit Sub
If Len(Combo1.Text) < 3 Then MsgBox "Error: Extension too short", vbExclamation, "Error": Exit Sub
For nk = 0 To File1.ListCount - 1
getit$ = File1.List(nk)
period = InStr(1, "" + Text2.Text + "", ".", 1)
leftperiod = InStr(1, getit$, ".", 1)
leftperiod = leftperiod - 1
peleft$ = Left(getit$, leftperiod)
If Right(File1.Path, 1) = "\" Then
FileCopy File1.Path + getit$, File1.Path + peleft$ + Combo1.Text
Else
FileCopy File1.Path + "\" + getit$, File1.Path + "\" + peleft$ + Combo1.Text
End If
progr = File1.ListCount
progre = 100 / progr
If ProgressBar1.Value + progre <= 100 Then
ProgressBar1.Value = ProgressBar1.Value + progre
Else
ProgressBar1.Value = 100
End If
c = ProgressBar1.Value
a = Len(c)
b = Left(c, 4)
Label2.Caption = b & "%"
Pause 1
Next nk
MsgBox "File changes were successful", vbInformation, "Successful"
ProgressBar1.Value = 0
Label2.Caption = "0%"
Exit Sub
Erhandle:
MsgBox "Error: Unknown error occurred", vbExclamation, "Error"
ProgressBar1.Value = 0
Label2.Caption = "0%"
Exit Sub
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "begin a mass file name change in directory"
End Sub

Private Sub Command2_Click()
On Error GoTo Errhandle
If Not File_IfileExists("" + Text2.Text) Then MsgBox "Error: File does not exist", vbExclamation, "Error": Exit Sub
If Left(Combo2.Text, 1) <> "." Then MsgBox "Error: Leave period at beginning of extension selection", vbExclamation, "Error": Exit Sub
If Len(Combo2.Text) > 5 Then MsgBox "Error: Extension too long", vbExclamation, "Error": Exit Sub
If Len(Combo2.Text) < 3 Then MsgBox "Error: Extension too short", vbExclamation, "Error": Exit Sub
index = File1.ListIndex
getit$ = File1.List(index)
period = InStr(1, "" + Text2.Text + "", ".", 1)
leftperiod = InStr(1, getit$, ".", 1)
ugh = Left(Text2.Text, period)
leftperiod = leftperiod - 1
peleft$ = Left(getit$, leftperiod)
If Right(File1.Path, 1) = "\" Then
FileCopy Text2.Text, File1.Path + peleft$ + Combo2.Text
Else
FileCopy Text2.Text, File1.Path + "\" + peleft$ + Combo2.Text
End If
File1.Refresh
MsgBox "File change successful. Check in " + File1.Path, vbInformation, "Successful"
Exit Sub
Errhandle:
MsgBox "Error: Unknown error occurred", vbExclamation, "Error"
Exit Sub
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "change just one file at a time"
End Sub

Private Sub Command3_Click()
File1.Refresh
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "refresh file list"
End Sub

Private Sub Command4_Click()
MsgBox "I attempted to make this program work smoothly but on every computer there are variations from mine in great ways. This works perfectly for everything on my computer but on a friends it doesnt convert .art files to .jpg's. You must NOT have 2 file anmed the same with different extensions. i did nothing to prevent an error when that happens and you must not go searching for errors that would never occur unless you searched very hard for them. That is pointless and it really does nothing for you. This program was mainly for me for making all 108 of my html files into txt's so i could edit them and then switch them back. this program does NOT delete files that it has changed. it copies them into the same directory with the same file name with a different extension. only because i would not want someone deleting files i wanted changed. have a nice day", vbInformation, "Help"
MsgBox "spoutcast@hotmail.com" + Chr(13) + "http://spoutcast.cjb.net" + Chr(13) + "http://spoutcast.i85.net" + Chr(13) + "http://spoutnet.cjb.net" + Chr(13) + "http://spoutcast.com" + Chr(13) + "blah blah blah, tutorials and stuff. i really suck at vb so dont expect anything big. this goddamn thing was hard to make from scratch. stupid instr and all those rights and lefts.", vbInformation, "About"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Help"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "http://spoutcast.cjb.net"
Label1.BackColor = &H0&
Label1.ForeColor = &HFF&
Pause 0.0000001
Label1.BackColor = &HC0C0C0
Label1.ForeColor = &H0&
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path

End Sub

Private Sub File1_Click()
index = File1.ListIndex
getit$ = File1.List(index)
If Right(File1.Path, 1) = "\" Then
Text2.Text = File1.Path + getit$
Else
Text2.Text = File1.Path + "\" + getit$
End If

End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "http://spoutcast.cjb.net"
Label1.BackColor = &H0&
Label1.ForeColor = &HFF&
Pause 0.0000001
Label1.BackColor = &HC0C0C0
Label1.ForeColor = &H0&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "http://spoutcast.cjb.net"
Label1.BackColor = &H0&
Label1.ForeColor = &HFF&
Pause 0.0000001
Label1.BackColor = &HC0C0C0
Label1.ForeColor = &H0&
End Sub

Private Sub ProgressBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "how long it will take to change all file"
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "filename"
End Sub
