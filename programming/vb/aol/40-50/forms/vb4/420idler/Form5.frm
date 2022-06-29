VERSION 4.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   615
   ClientLeft      =   6525
   ClientTop       =   4080
   ClientWidth     =   2295
   Height          =   1020
   Left            =   6465
   LinkTopic       =   "Form5"
   ScaleHeight     =   615
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   Top             =   3735
   Width           =   2415
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "enter # or word"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line16 
      X1              =   960
      X2              =   2280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line15 
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line14 
      X1              =   960
      X2              =   960
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line13 
      X1              =   840
      X2              =   840
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line12 
      X1              =   2280
      X2              =   2280
      Y1              =   600
      Y2              =   0
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line8 
      X1              =   1080
      X2              =   960
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line7 
      X1              =   2160
      X2              =   2280
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   2280
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line5 
      X1              =   1080
      X2              =   960
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line4 
      X1              =   720
      X2              =   840
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   0
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   0
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   840
      Y1              =   120
      Y2              =   0
   End
End
Attribute VB_Name = "Form5"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call ChatSend("o·.× 4²o ìdlér")
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause("1.4")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause("1.4")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause("1.4")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Call ChatSend(Text1.Text)
Call Pause(".7")
Unload Form5
End Sub

Private Sub Form_Load()
Call FormOnTop(Me)
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


    FormDrag Me
    End Sub

