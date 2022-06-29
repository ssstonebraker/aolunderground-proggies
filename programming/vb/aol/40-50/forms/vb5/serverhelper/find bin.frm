VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Bin"
   ClientHeight    =   1905
   ClientLeft      =   2055
   ClientTop       =   1995
   ClientWidth     =   2025
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Menu info 
      Caption         =   "&Info"
   End
   Begin VB.Menu save 
      Caption         =   "&Save"
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.ListCount = 0 Then List1.AddItem Text1.text: Exit Sub

For i = 0 To List1.ListCount - 1
num = LCase$(List1.List(i))
If num = LCase$(Text1.text) Then Exit Sub
Next i
List1.AddItem Text1.text
List1.AddItem Form7.List4
List1.AddItem Form13.List4
End Sub

Private Sub Form_Load()
StayOnTop Me
Dim a As Variant
Dim b As Variant
On Error GoTo kook
a = 1
List1.Clear
Open CStr(App.Path + "\filename.lst") For Input As a
While (EOF(a) = False)
Line Input #a, b
List1.AddItem b
Wend

Close a
'End If
kook:
End Sub

Private Sub info_Click()
MsgBox "If you're like me, you dont like reading through ten lists for certain items.  I always search for the same things when running through the rooms.  And I hated typing each one out.  Well this options saves  your Find items.  Just click run on the pull dowm menue, and it wll ask for the items listed in your Find Bin.", vbInformation, "Info................."
End Sub

Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub save_Click()
If List1.ListCount = 0 Then Exit Sub
Dim a As Integer
Dim b As Variant
On Error GoTo C
a = 2
Open CStr(App.Path + "\Filename.lst") For Output As a
b = 0
Do While b < List1.ListCount
Print #a, List1.List(b)
b = b + 1
Loop
Close a
'End If
C:
Exit Sub
End Sub

Private Sub Timer1_Timer()

End Sub
