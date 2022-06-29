VERSION 5.00
Begin VB.Form MailList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailing Lister '99 - By: EccO"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "maillist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "&Float Window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Save List"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lo&ad/New List"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "C&opy"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Exit"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "C&lear List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A&dd"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   495
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
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Kill Doubles"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "MailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mailing Lister '99 - By: EccO (xeccox@mailcity.com)

'Welcome to my latiest example of VB5.0 programming.
'This example shows you how to keep peoples e-mail
'saved in a list box.  I do submit alot of examples to KnK
'and so, but I would like a little feedback..

'There is only one problem with my program, you have
'to save the file in a folder, not on desktop, like in programs.

'~EccO

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Call StayOnTop(FileOpt)
Call StayOnTop(Me)
Else
Call NotOnTop(FileOpt)
Call NotOnTop(Me)
End If

End Sub

Private Sub Command1_Click()
Me.Hide
FileOpt.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Loo As Integer
Dim CheckLoo As Integer
Dim GotListItem As String
Dim ItemCheck As String

Me.Enabled = False

For Loo = 0 To (List1.ListCount - 1)
ItemCheck = List1.List(Loo)
If Err Then Exit For
List1.RemoveItem Loo
'For CheckLoo = 0 To (List1.ListCount - 1)
DoEvents
RECheck:
'GotListItem = List1.List(CheckLoo)
GotListItem = List1.List(Loo)
If Err Then Exit For
If ItemCheck = GotListItem Then
List1.RemoveItem Loo
If Err Then Exit For
GoTo RECheck:
End If
'Next CheckLoo
List1.AddItem ItemCheck
Next Loo

'List1.RemoveItem 0
Label1.Caption = List1.ListCount
Me.Enabled = True
End Sub

Private Sub Command3_Click()

If Trim$(Text1.Text) <> "" Then
List1.AddItem Trim$(Text1.Text)
Text1.Text = ""
Label1.Caption = List1.ListCount
Text1.SetFocus
End If

End Sub

Private Sub Command4_Click()
List1.Clear
Label1.Caption = List1.ListCount
End Sub

Private Sub Command5_Click()
Dim FF As Integer

FF = FreeFile

Call Command9_Click
Open App.Path & "\" & "MailList.ini" For Output As #FF
Write #1, Label2.Caption
Close #FF

End
End Sub

Private Sub Command8_Click()
Dim Loo As Integer
Dim CopyList As String

Me.Enabled = False

For Loo = 0 To (List1.ListCount - 1)
CopyList = CopyList & List1.List(Loo) & ", "
Next Loo

Clipboard.SetText CopyList
Me.Enabled = True

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim FileName As String
Dim FF As Integer
Dim Loo As Integer

Me.Enabled = False

If Trim$(Label2.Caption) = "" Then
Exit Sub
End If

If MailList.List1.ListCount = 0 Then
MsgBox "There's Nothing to Save!", 64, "Save List"
Exit Sub
End If

FileName = Label2.Caption
FF = FreeFile

Open FileName For Output As #FF

For Loo = 0 To (MailList.List1.ListCount - 1)
Print #FF, MailList.List1.List(Loo)
Next Loo

Close #FF

Me.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim FF As Integer
Dim TmpStr As String
Dim GetSN As String

FF = FreeFile
Open App.Path & "\" & "MailList.ini" For Input As #FF
Input #1, TmpStr
Label2.Caption = TmpStr
Close #FF

FF = FreeFile
Open Label2.Caption For Input As #FF
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

Load FileOpt
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_DblClick()
On Error Resume Next
Text1.Text = List1.List(List1.ListIndex)
List1.RemoveItem List1.ListIndex
MailList.Label1 = MailList.List1.ListCount
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim$(Text1.Text) <> "" Then
List1.AddItem Trim$(Text1.Text)
Text1.Text = ""
Label1.Caption = List1.ListCount
KeyAscii = 0
End If
End If

End Sub
