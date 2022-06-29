VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dos's bas compare example [1/18/98]"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtIntro 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmMain.frx":0000
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&file"
      Begin VB.Menu mnuExit 
         Caption         =   "&exit"
      End
   End
   Begin VB.Menu mnuExamples 
      Caption         =   "&examples"
      Begin VB.Menu mnuLoading 
         Caption         =   "&loading a module"
      End
      Begin VB.Menu mnuFindingVariables 
         Caption         =   "&finding variables, etc."
      End
      Begin VB.Menu mnuCompare 
         Caption         =   "&comparing procedures"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&other"
      Begin VB.Menu mnuConclusion 
         Caption         =   "&conclusion"
      End
      Begin VB.Menu mnuWriting 
         Caption         =   "&writing a bas scanner"
      End
      Begin VB.Menu mnuHow 
         Caption         =   "&how to beat a scanner"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "&contact"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmCompare
    Unload frmLoad
    Unload frmTricks
End Sub

Private Sub mnuCompare_Click()
    frmCompare.Visible = True
End Sub

Private Sub mnuConclusion_Click()
    MsgBox "as you can see, it is very possible to scan two bas files for copied code. this example shows how to;" & vbCrLf & "load a bas file" & "identify code stealing tricks" & vbCrLf & "use a reliable way to compare procedures" & vbCrLf & vbCrLf & "writing a bas file scanner can never replace actually looking at the file for yourself. also, finding 10% or less of copied code in a file doesn't mean much. however, most bas files are well over 50%." & vbCrLf & "also, i did not write this program to catch code stealers. it was made per a request from knk." & vbCrLf & "in the end, if you understand this project, i'm sure you'll come to the conclusion that yes, a relaible bas file scanner can be created." & vbCrLf & vbCrLf & "dos"
End Sub

Private Sub mnuContact_Click()
    MsgBox "email: " & Chr(9) & "xdosx@hotmail.com" & vbCrLf & "aim: " & Chr(9) & "xdosx" & vbCrLf & vbCrLf & "please do not email or instant message me your questions without trying on your own first. i get a lot of email every day and although i read it all, i couldn't possibly answer it all. so if you're asking a question that can be answered easily (ie, your help file), i probably won't answer."
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFindingVariables_Click()
    frmTricks.Visible = True
End Sub

Private Sub mnuHow_Click()
    MsgBox "sure, this scanner can be fooled, however, there is one really good way to beat it..." & vbCrLf & vbCrLf & "... write your own code!"
End Sub

Private Sub mnuLoading_Click()
    frmLoad.Visible = True
End Sub

Private Sub mnuWriting_Click()
    MsgBox "this program does not contain everything required to write a bas file scanner. there is yet one more step. you must be able to load at least two bas files, then loop them against eachother, keeping track of matches. also, you will need to generate the related results from these totals." & vbCrLf & vbCrLf & "for those of you who can't loop through a list, i suggest you buy a book and learn. to those of you who can loop through a list, but not compare two lists, i would suggest that you relax, be patient, and learn too. both of these are really basic tricks and i'm not going to waste time on them." & vbCrLf & vbCrLf & "this is not the actual source for my scanner. instead, the code in this example is just the watered down basics. my program requires extensive use a class module and arrays. these were necessary in that i wanted to scan a list of files against another list of files"
End Sub
