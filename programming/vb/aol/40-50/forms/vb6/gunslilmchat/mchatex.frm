VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form mchatex 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "guns lil mchat example"
   ClientHeight    =   3555
   ClientLeft      =   1935
   ClientTop       =   1845
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mchatex.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "mchatex.frx":08CA
   ScaleHeight     =   3555
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   3600
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   600
      MaxLength       =   100
      TabIndex        =   1
      ToolTipText     =   "Text to send"
      Top             =   2710
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1800
      Left            =   560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   660
      Width           =   5400
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "guns lil mchat example"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Send text"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "People in room"
      Top             =   400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Room name"
      Top             =   405
      Width           =   2175
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   2760
      X2              =   2760
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   480
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   2760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   6000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   6000
      X2              =   6000
      Y1              =   600
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   6000
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   2520
   End
End
Attribute VB_Name = "mchatex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'M-Chat Example
'Written by: GuN
'Send all E-Mails to i gun l@aol.com

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
'gets the text from the chatroom
Text1.SelStart = Len(Text1.Text)
Text1.SelText = vbCrLf & Screen_Name & ":   " & What_Said
End Sub


Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
'adds the following line to text1, makes the form alway's ontop, and turns dos's chat scanner on
Text1.Text = "GuN:          *** Welcome to: guns lil mchat example ***"
StayOnTop Me
Chat1.ScanOn
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when the mouse is clicked on the form, it will drag the form to any position
FormDrag Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
'when the form is unloaded, it will turn off dos's chat scanner
Chat1.ScanOff
End Sub


Private Sub Label3_Click()
'sends the text to the chatroom
Chat1.ChatSend Text2.Text
Text2.Text = ""
End Sub


Private Sub Label4_Click()
'exits and turns off dos's chat scanner
Chat1.ScanOff
End
End Sub


Private Sub Label5_Click()
'exits and turns off dos's chat scanner
Chat1.ScanOff
End
End Sub


Private Sub Label6_Click()
'minimizes the current form
WindowState = 1
End Sub


Private Sub Label7_Click()

End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'moves the current form to where ever you want
FormDrag Me
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
'sends text to chat when enter is hit
If KeyAscii = 13 Then
Chat1.ChatSend Text2.Text
Text2.Text = ""
End If
End Sub


Private Sub Timer1_Timer()
'sets label2's caption to how many people are in the current room
Label2.Caption = RoomCount
Label1.Caption = XAOL4_GetCurrentRoomName
End Sub


