VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "m-chat example - by ecco"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   1320
      Top             =   2520
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'm-chat example - by ecco (xeccox@hotmail.com)
'*******************
'welcome to another example by me, this example
'will show you how to make a simple m-chat
'using dos32.bas and dos's vb5 chat.ocx
'enjoy!

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
Text1.SelStart = Len(Text1.Text)
Text1.SelText = vbCrLf & Screen_Name & ":" & Chr(9) & What_Said
End Sub

Private Sub Command1_Click()
ChatSend "" & Text2.Text & ""
End Sub

Private Sub Form_Load()
FormOnTop Me
Chat1.ScanOn
End Sub
