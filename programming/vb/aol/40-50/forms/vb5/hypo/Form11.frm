VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00404040&
   Caption         =   "Group Creator (For all those Warez Groups)"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form11"
   ScaleHeight     =   3585
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Remove Room"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3720
      Top             =   1560
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Text            =   "Room Name"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Addroom Name"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form11.frx":0000
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "How Many rooms left to scan"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Form11

Form11.Hide

End Sub

Private Sub Command4_Click()
List2.AddItem Text2
Label1.Caption = Val(Label1.Caption) + 1
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim i As Integer

  For i = 0 To List2.ListCount - 1
   If List2.Selected(i) Then
    
    List2.RemoveItem (i)
    End If
    Next i

 

End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form11
List2.AddItem ("Zeraw")
List2.AddItem ("MP3")
List2.AddItem ("Poa")
List2.AddItem ("Fate")
List2.AddItem ("Leet")
List2.AddItem ("Games")
List2.AddItem ("games3")

End Sub

Private Sub Timer1_Timer()
Timer2.Enabled = True

For i% = 0 To List2.ListCount - 1 ' or whatever List# is
Call KeyWord("aol://2719:2-2-" + List2.List(i%))
TimeOut 2
AppActivate "America  Online"
SendKeys " "
AddRoomToListBox List1
 
Label1.Caption = Val(Label1.Caption) - 1
If i% = 0 Then Timer1.Enabled = False
Next i%
End Sub

Private Sub Timer2_Timer()
If Label1.Caption = "0" Then
For i% = 0 To List1.ListCount - 1 ' or whatever List# is
Call IMKeyword(List1.List(i%), Text1.Text) 'Or whatever the Textbox is
Next i%
End If
End Sub
