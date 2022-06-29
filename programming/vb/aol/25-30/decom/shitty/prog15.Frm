VERSION 4.00
Begin VB.Form Form14 
   Caption         =   "Open which file?"
   ClientHeight    =   3480
   ClientLeft      =   1785
   ClientTop       =   1665
   ClientWidth     =   3600
   Height          =   3885
   Left            =   1725
   LinkTopic       =   "Form14"
   ScaleHeight     =   3480
   ScaleWidth      =   3600
   Top             =   1320
   Width           =   3720
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   225
      Left            =   2700
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   3210
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1860
      TabIndex        =   3
      Top             =   90
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   2595
      Left            =   1830
      TabIndex        =   2
      Top             =   600
      Width           =   1725
   End
   Begin VB.DirListBox Dir1 
      Height          =   2730
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   1635
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "Form14"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command2_Click()
If Text1.text = " " Then
MsgBox "Enter a name for the file that you want to load"
Exit Sub
End If
nacho = Dir1.Path & "\" & Text1.text
Form7.RichTextBox1.LoadFile nacho
Form7.Show
Form14.Hide
End Sub


Private Sub Dir1_Change()
mypath = Dir1.Path
File1.Path = mypath
End Sub


Private Sub Drive1_Change()
mypath = Drive1
Dir1.Path = mypath
End Sub


Private Sub File1_DblClick()
Text1.text = File1
End Sub


Private Sub Gauge1_Change()
Gauge1 = gauHoriz
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


