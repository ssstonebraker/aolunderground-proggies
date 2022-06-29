VERSION 5.00
Begin VB.Form PWSD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FoRBiDons Personal File Scanner example"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   4095
      Begin VB.TextBox Text1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   1620
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      Begin VB.FileListBox File1 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2190
         Left            =   120
         Pattern         =   "*.exe"
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Abadi MT Condensed"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "Type string to scan for here."
         Top             =   210
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
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
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Scan"
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
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "PWSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Scan_For(sFile$, ByVal sWhat$)
Dim VariantA As Variant
Dim VariantB As Variant
Dim VariantC As Variant
Dim VariantD As Variant
Dim SingleA As Single
Dim StringA As String
Dim EnterKey As String

On Error Resume Next
Open sFile$ For Binary As #1
    EnterKey$ = Chr$(13) + Chr$(10)
    msg$ = ""
    VariantA = LOF(1)
    VariantB = VariantA
    VariantC = 1

    If VariantB > 32000 Then
        VariantD = 32000
    ElseIf VariantB = 0 Then
        VariantD = 1
    Else
        VariantD = VariantB
    End If

    StringA$ = String$(VariantD, " ")
    Get #1, VariantC, StringA$

    SingleA! = InStr(1, StringA$, sWhat$, 1)

    If SingleA! Then
        Scan_For = 0
    Else
        Scan_For = 1
    End If
Close #1
End Function

Private Sub Command1_Click()
'Thsi is how it works
On Error Resume Next
If File1.filename = "" Then Exit Sub
FPath$ = Dir1.Path
If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ + "\"
Selectedfile$ = FPath$ + File1.filename

SearchString1% = Scan_For(Selectedfile$, Text2.Text): DoEvents

If SearchString1% = 0 Then

    outcome$ = " DOES "
Else
    outcome$ = " does NOT"
End If
MsgBox "The file" & outcome$ & "match", vbCritical, "ForBiDons Person File Scanner"
End Sub

Private Sub Command2_Click()
Call File_Delete(Text1.Text)

End Sub

Private Sub Command3_Click()
MsgBox "This is my Personal File Scanner...what it does: first you type in a string ( a line of code) in the personal file scanner...then you pick the file you want to scan and push scan...it scan the file for the string you told it to.", , "Personal File Scanner"
MsgBox "For example, if you wanted to scan a file for cuss word...you would type a cuss word in and then pick the file that you want to scan then push scan...if the file has any cuss words in it then it will tell you", , "Personal file scanner"
End Sub

Private Sub Command4_Click()
Unload PWSD2
Load PWSD
PWSD.Show
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveErrs
    Dir1.Path = Drive1.Drive
    Text1.Text = Drive1.Drive
Exit Sub
DriveErrs:
    Select Case Err
        Case 68
            MsgBox prompt:="Drive not ready. Please insert disk in drive. Then we can scan The File.", _
            Buttons:=vbExclamation
            Drive1.Drive = Dir1.Path
            Text1.Text = Drive1.Drive
            Exit Sub
        Case Else
            MsgBox prompt:="Application error.", Buttons:=vbExclamation
    End Select
End Sub

Private Sub File1_Click()
Text1.Text = File1.Path + "\" + File1.filename
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub
