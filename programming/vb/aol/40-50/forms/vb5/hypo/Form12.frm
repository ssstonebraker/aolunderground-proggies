VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H80000007&
   Caption         =   "Annoy Bots"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   ScaleHeight     =   1740
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Spiral Scroll"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Text            =   "HyPO By ToaST"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Soundz"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sound Annoy "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If

SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5

SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
TimeOut 0.5
SendChat "{S " + Combo1
SendChat "ToaSTs AOL Chat Annoy"
End Sub

Private Sub Command2_Click()
Unload Form12
Form12.Hide


End Sub

Private Sub Command3_Click()
If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If

X = txt.Text
thastart:
Dim MYLEN As Integer
MyString = txt.Text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
txt.Text = MYSTR
SendChat (MYSTR)
TimeOut 1
If txt.Text = X Then
Exit Sub
End If
GoTo thastart

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
Combo1.AddItem ("GoodBye")
Combo1.AddItem ("FileDone")
Combo1.AddItem ("Welcome")
Combo1.AddItem ("IM")
End Sub
