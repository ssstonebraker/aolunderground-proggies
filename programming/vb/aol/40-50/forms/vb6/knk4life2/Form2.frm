VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Room Buster"
   ClientHeight    =   1245
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   2370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Room Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form2.frx":0000
         Left            =   120
         List            =   "Form2.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bust As Boolean


Private Sub cmdStart_Click()
 Dim MyRoom As String, CurRoom As String, Tries As Long
    MyRoom$ = LCase(ReplaceString(Combo1, " ", ""))
    bust = True
    Do
        DoEvents
        Call PrivateRoom(MyRoom$)
        Call WaitForOKOrRoom(MyRoom$)
        DoEvents
        CurRoom$ = GetCaption(FindRoom)
        CurRoom$ = LCase(ReplaceString(CurRoom$, " ", ""))
        Tries& = Tries& + 1
        Text1.Text = Tries& & " attempts completed."
    Loop Until bust = False Or CurRoom$ = MyRoom$
    If MyRoom$ = CurRoom$ Then
        ChatSend (GetUser & " busted in " & Combo1 & " in " & Tries& & " attempt(s).")
    End If
End Sub

Private Sub cmdStop_Click()
bust = False
End Sub

Private Sub Form_Load()
FormOnTop Me
End Sub

Private Sub save_Click()

End Sub

