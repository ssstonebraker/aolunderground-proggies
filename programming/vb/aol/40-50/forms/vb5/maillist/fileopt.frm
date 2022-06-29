VERSION 5.00
Begin VB.Form FileOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load/New List"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "fileopt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "New File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add to List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Load List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   2760
      Pattern         =   "*.SNC"
      TabIndex        =   2
      Top             =   435
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FileOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.Locked = False
Command3.Enabled = False
Else
Text1.Locked = True
Command3.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Me.Hide
MailList.Show

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Loc As String
Dim FileName As String
Dim FF As Integer
Dim GetSN As String

Me.Enabled = False

SNCollector.List1.Clear

If Trim$(Text1.Text) = "" Then
MsgBox "Please Enter a File Name!", 64, "Load List"
Me.Enabled = True
Exit Sub
End If

Loc = Dir1 & "\"
If UCase$(Right$(Text1.Text, 4)) <> ".SNC" Then
Text1.Text = Text1.Text + ".SNC"
End If
FileName = Text1.Text
MailList.Label2.Caption = Loc & FileName
FF = FreeFile

Open Loc & FileName For Input As #FF

Do
Input #1, GetSN
If Trim$(GetSN) <> "" Then
MailList.List1.AddItem GetSN
End If
GetSN = ""
If Err Then Exit Do
Loop

Close #FF

FF = FreeFile
Open App.Path & "\" & "MailList.ini" For Output As #FF
Write #1, Loc & FileName
Close #FF

MailList.Label1 = MailList.List1.ListCount
MailList.Show
Unload Me

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim Loc As String
Dim FileName As String
Dim FF As Integer
Dim GetSN As String

Me.Enabled = False

If Trim$(Text1.Text) = "" Then
MsgBox "Please Enter a File Name!", 64, "Add to List"
Me.Enabled = True
Exit Sub
End If

Loc = Dir1 & "\"
FileName = Text1.Text
FF = FreeFile

Open Loc & FileName For Input As #FF

Do
Input #1, GetSN
If Trim$(GetSN) <> "" Then
MailList.List1.AddItem GetSN
End If
GetSN = ""
If Err Then Exit Do
Loop

Close #FF

MailList.Label1 = MailList.List1.ListCount
MailList.Show
Unload Me
End Sub

Private Sub Command7_Click()
Dim Loc As String
Dim FileName As String
Dim FF As Integer
Dim Loo As Integer

Me.Enabled = False

If Trim$(Text1.Text) = "" Then
MsgBox "Please Enter a File Name!", 64, "Save List"
Me.Enabled = True
Exit Sub
End If
If MailList.List1.ListCount = 0 Then
MsgBox "There's Nothing to Save!", 64, "Save List"
Me.Enabled = True
Exit Sub
End If
Loc = Dir1 & "\"
FileName = Text1.Text & ".SNC"
FF = FreeFile

Open Loc & FileName For Output As #FF

For Loo = 0 To (MailList.List1.ListCount - 1)
Print #FF, MailList.List1.List(Loo)
Next Loo

Close #FF

MailList.Show
Unload Me

End Sub

Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1
End Sub

Private Sub File1_Click()
Text1.Text = File1.List(File1.ListIndex)
End Sub

Private Sub Form_Load()

If MailList.Check1.Value = 1 Then
Call StayOnTop(Me)
Else
Call NotOnTop(Me)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MailList.Show
End Sub
