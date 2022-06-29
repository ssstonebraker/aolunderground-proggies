VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo Chat Room Tool"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "GhostTech@usa.net"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Chat demonstration tool by Ghost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormOnTop Me
Do: DoEvents
Do: DoEvents
 AORoom = FindRoom()
 AOEdit = FindWindowEx(AORoom, 0&, "RICHCNTL", vbNullString)
Loop Until AOEdit <> 0
 
    TextLength = SendMessage(AOEdit, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength, 0&)
    Call SendMessageByString(AOEdit, WM_GETTEXT, TextLength + 1, buffer$)

If lasttype = buffer$ Then GoTo skipit:

lasttype = buffer$

buffer$ = ReplaceString(buffer$, Chr(13), Chr(13) & Chr(10))

Text1.Text = buffer$
Text1.SelStart = Len(Text1.Text)
skipit:
Loop
End Sub

Private Sub Command2_Click()
MsgBox Asc(Text1)
End Sub
