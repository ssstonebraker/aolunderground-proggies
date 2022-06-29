VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sequence Linker - spout"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Info"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "http://"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line10 
      X1              =   240
      X2              =   240
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line9 
      X1              =   240
      X2              =   240
      Y1              =   2400
      Y2              =   2520
   End
   Begin VB.Line Line8 
      X1              =   720
      X2              =   240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line7 
      X1              =   2880
      X2              =   2640
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line6 
      X1              =   3600
      X2              =   3600
      Y1              =   960
      Y2              =   1200
   End
   Begin VB.Line Line5 
      X1              =   1080
      X2              =   1200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      X1              =   2160
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   2880
      Y1              =   2400
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   1920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   1680
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://spoutcast.cjb.net"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo stupid
A = 0
b$ = ""
List1.ListIndex = -1
List2.ListIndex = -1
one = Trim(Label2.Caption)
two = Trim(Label3.Caption)
If List1.ListCount = 0 Then MsgBox "Error: No Files To Link", vbCritical, "Error"
    For nk = 0 To List1.ListCount - 1
    c = List2.ListIndex + 1
    pindex = List2.List(c)
    d = List1.ListIndex + 1
    windex = List1.List(d)
    A = A + 1
    F$ = A
    b$ = b$ + "<a href=" + Chr(34) + List1.List(nk) + Chr(34) + ">" + one + F$ + two + "</a>"
    
    Next nk
    Form2.Text1.Text = b$
    Form2.Show
    Exit Sub
stupid:
MsgBox "Error: Unknown", vbExclamation, "Error"
Exit Sub
End Sub

Private Sub Command2_Click()
If Label2.Caption = "[ " Then Label2.Caption = "( ": Label3.Caption = " )": Exit Sub
If Label2.Caption = "( " Then Label2.Caption = "{ ": Label3.Caption = " }": Exit Sub
If Label2.Caption = "{ " Then Label2.Caption = "- ": Label3.Caption = " -": Exit Sub
If Label2.Caption = "- " Then Label2.Caption = ".: ": Label3.Caption = " :.": Exit Sub
If Label2.Caption = ".: " Then Label2.Caption = "< ": Label3.Caption = " >": Exit Sub
If Label2.Caption = "< " Then Label2.Caption = "‹ ": Label3.Caption = " ›": Exit Sub
If Label2.Caption = "‹ " Then Label2.Caption = "[ ": Label3.Caption = " ]": Exit Sub
End Sub

Private Sub Command3_Click()
MsgBox "yeeeeeea right. its for me, hence, you can edit the lists yourself. i made them save as text files, whip out notepad on them.", vbCritical, "shutup"
End Sub

Private Sub Command4_Click()
A$ = InputBox("What would you like to name the file?", "Save")
If A$ = "" Then Exit Sub
If Right(App.Path, 1) = "\" Then
Call SaveListBox(App.Path + A$ + ".txt", List1)
Call SaveListBox(App.Path + A$ + "numb" + ".txt", List2)
Else
Call SaveListBox(App.Path + "\" + A$ + ".txt", List1)
Call SaveListBox(App.Path + "\" + A$ + "numb" + ".txt", List2)
End If
End Sub

Private Sub Command5_Click()
A$ = InputBox("Just smack the file name with no extension in here", "Load")
If A$ = "" Then Exit Sub
If Right(App.Path, 1) = "\" Then
Call Loadlistbox(App.Path + A$ + ".txt", List1)
Call Loadlistbox(App.Path + A$ + "numb" + ".txt", List2)
Else
Call Loadlistbox(App.Path + "\" + A$ + ".txt", List1)
Call Loadlistbox(App.Path + "\" + A$ + "numb" + ".txt", List2)
End If
End Sub

Private Sub Command6_Click()
MsgBox "This is mainly for me. i was so pissed off that i had to link like 60 files doing the [1][2][3] method with many files. and when i get pissed i make programs that help me out. i also havent made a tutorial in a while. i know how to use it, not very user friendly because its for me, so all in all: blah", vbInformation, "Nfo"
End Sub

Private Sub List1_Click()
windex = List1.ListIndex
If List1.List(windex) = "" Then Exit Sub
List2.ListIndex = List1.ListIndex
Text1.Text = List1.List(windex)

End Sub

Private Sub List2_Click()
windex = List2.ListIndex
If List2.List(windex) = "" Then Exit Sub
List1.ListIndex = List2.ListIndex
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = &HD Then
If Text1.Text = "" Then Exit Sub
List1.AddItem Text1.Text
List2.AddItem (List2.ListCount + 1)
Text1.SetFocus
End If
End Sub
