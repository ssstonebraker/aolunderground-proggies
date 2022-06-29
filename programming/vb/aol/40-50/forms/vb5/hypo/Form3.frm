VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass IM"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove SN"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add SN"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Text            =   "                                                                                  ToaST's Mass Imer"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Imer"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form3.frx":0000
      Top             =   240
      Width           =   2895
   End
   Begin VB.ListBox lst 
      Height          =   1035
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Room"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mass IM"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ToaST$ = "ToaST's Mass IMer"

A = Len(ToaST$)
For w = 1 To A Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "CD2626" & Chr$(34) & "><sup><B>" & ToaSTR$ & "</sup></B>" & "<FONT COLOR=" & Chr$(34) & "#8B1A1A" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#8B3A3A" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#4E2F2F" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$

For i% = 0 To lst.ListCount - 1 ' or whatever List# is
Call IMKeyword(lst.List(i%), Text1.Text + WavYChaTRBb) 'Or whatever the Textbox is
Next i%
End Sub

Private Sub Command2_Click()
If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If
AddRoomToListBox lst
End Sub

Private Sub Command3_Click()
Unload Form3

Form3.Hide

End Sub

Private Sub Command4_Click()
On Error Resume Next
lst.AddItem Text3

End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim i As Integer

  For i = 0 To lst.ListCount - 1
   If lst.Selected(i) Then
    
    lst.RemoveItem (i)
    End If
    Next i
End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form3
End Sub

