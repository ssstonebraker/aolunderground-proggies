VERSION 5.00
Begin VB.Form PWSD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PWSD sample by ÐøøM n Stock"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   0
      TabIndex        =   9
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
         TabIndex        =   10
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
            Name            =   "Arial Narrow"
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
         TabIndex        =   7
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
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2340
         Left            =   120
         Pattern         =   "*.exe"
         TabIndex        =   8
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
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete File"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "PassWord Stealer Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Sample PWSD By ÐøøM and Stock"
      BeginProperty Font 
         Name            =   "Lydian Csv BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   3840
      Width           =   2295
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
On Error Resume Next
If File1.FileName = "" Then Exit Sub
FPath$ = Dir1.Path
If Right(FPath$, 1) <> "\" Then FPath$ = FPath$ + "\"
Selectedfile$ = FPath$ + File1.FileName

SearchString1% = Scan_For(Selectedfile$, "main.idx"): DoEvents

If SearchString1% = 0 Then

    outcome$ = " "
Else
    outcome$ = " Not "
End If
MsgBox "This is" & outcome$ & "a PassWord Stealer!", vbCritical, "PWS Detector Sample by ÐøøM"
End Sub

Private Sub Command2_Click()
Call File_Delete(Text1.Text)

End Sub

Private Sub Command3_Click()
Unload Me
Text1.Text = ""
End

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
    Select Case err
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
Text1.Text = File1.Path + "\" + File1.FileName
End Sub

