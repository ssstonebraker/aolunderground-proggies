VERSION 4.00
Begin VB.Form mchatex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M-Chat Example By Elroy"
   ClientHeight    =   4605
   ClientLeft      =   1980
   ClientTop       =   2175
   ClientWidth     =   6090
   Height          =   4995
   Left            =   1920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Top             =   1845
   Width           =   6210
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   0
      Top             =   0
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "mchatex"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
Text1.SelStart = Len(Text1.Text)
Text1.SelText = vbCrLf & Screen_Name & ":     " & Chr(9) & What_Said
End Sub


Private Sub Command1_Click()
Chat1.ChatSend "" & Text2.Text & ""
Text2.Text = ""
End Sub


Private Sub Form_Load()
MsgBox "this is an m-chat example by elroy.  please dont steal the credit!", vbOKOnly, "remember!"
Text1.Text = "OnlineHost:              Welcome to elroy's m-chat example"
Chat1.ScanOn
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Chat1.ScanOff
End Sub

