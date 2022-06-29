VERSION 4.00
Begin VB.Form Form13 
   Caption         =   "Save As?"
   ClientHeight    =   3480
   ClientLeft      =   5775
   ClientTop       =   1740
   ClientWidth     =   3540
   Height          =   3885
   Left            =   5715
   LinkTopic       =   "Form13"
   ScaleHeight     =   3480
   ScaleWidth      =   3540
   Top             =   1395
   Width           =   3660
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   225
      Left            =   2640
      TabIndex        =   6
      Top             =   3210
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   3210
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1770
      TabIndex        =   3
      Top             =   540
      Width           =   1725
   End
   Begin VB.FileListBox File1 
      Height          =   2205
      Left            =   1770
      TabIndex        =   2
      Top             =   960
      Width           =   1725
   End
   Begin VB.DirListBox Dir1 
      Height          =   2730
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   1665
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   210
      Width           =   1725
   End
End
Attribute VB_Name = "Form13"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form13.Hide
Form7.Show
End Sub

Private Sub Command2_Click()
If Text1.text = " " Then
MsgBox "Enter a name for the file"
Exit Sub
End If


Open File1.Path & "\" & Text1.text For Output As #1
thing$ = Form7.RichTextBox1.text
Write #1, thing$
Close #1
frank = Text1.text & " has been saved"
freak = vbOKOnly
julie = "Saved"
Bob = MsgBox(frank, freak, julie)
Form13.Hide
Form7.Show

End Sub

Private Sub Dir1_Change()
mypath = Dir1.Path
File1.Path = mypath
End Sub

Private Sub Drive1_Change()
mypath = Drive1
Dir1.Path = mypath
End Sub


Private Sub File1_Click()
file = File1
Text1.text = file
End Sub

Private Sub Form_Load()
mypath = CurDir
Drive1.Drive = CurDir
Dir1.Path = CurDir
File1 = CurDir
Call StayOnTop(Form13)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


