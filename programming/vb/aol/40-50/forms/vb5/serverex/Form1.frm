VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form1 
   Caption         =   "VoiD Flashmail Server"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2760
      Top             =   2520
   End
   Begin VB.TextBox txtwhat 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtwho 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   1200
      Top             =   240
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sup?
'This is Accesszero.  I made this, because I saw a post
'that said you can make a server in like 6 lines.
'well that aint true, but it is truely easy.
'feel free to use this as a starting point for your
'Server, as long as you include DoS, and Accesszero
'in your credits.
'What the hell, you don't need to include me, I only
'put subs in order.
'DoS wrote them.

'L8er,
'-Accesszero
'[Accesszero@juno.com]
'[http://www.members.tripod.com/~access_vb/]

'I am going to take a lot of shit for this,
'so use it correctly.


Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
txtwhat = What_Said
If LCase(txtwhat) = "/list" Then
For i = 0 To List1.ListCount - 1
a = a & Chr$(13) & Chr$(10) & List1.List(i)
Next i
Text3 = a
DoEvents
Call SendMail(Screen_Name, "Void Server List · 1 - " & List1.ListCount, Text3)
DoEvents
End If
If LCase(Left(txtwhat, 7)) = "/void -" Then
DoEvents
numba = Mid(txtwhat, InStr(txtwhat, "-") + 1)
Text3 = numba
If IsNumeric(numba) = True And numba < List1.ListCount Then
Pause 0.5
Call MailOpenEmailFlash(Text3 - 1)
DoEvents
DoEvents
Call MailSenderFlash(Text3 - 1)
DoEvents
DoEvents
Call MailForward(Screen_Name, "VoiD Server - " & numba & " of " & List1.ListCount, True)
DoEvents
DoEvents
Else
Chat1.ChatSend "" & Screen_Name & " - Invalid Number"
Exit Sub
End If
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Start" Then
Command1.Caption = "Stop"
Timer1.Enabled = True
Chat1.ChatSend "-VoiD Server By: Accesszero & DoS"
DoEvents
DoEvents
Chat1.ChatSend "type: /list to get a list of the mails"
DoEvents
DoEvents
Chat1.ChatSend "type: /sendX - X as index number."
DoEvents
DoEvents
Chat1.ScanOn
Else
Command1.Caption = "Start"
Timer1.Enabled = False
Chat1.ScanOff
Chat1.ChatSend "-VoiD Server By: Accesszero & DoS"
DoEvents
DoEvents
Chat1.ChatSend "Status: Inactive"
End If
End Sub

Private Sub Command2_Click()
Chat1.About
End Sub

Private Sub Form_Load()
Call stayontop(Me)
SetMailPrefs
DoEvents
List1.Clear
MailOpenFlash
Do Until List1.ListCount > 0
DoEvents
Call MailToListFlash(List1)
Loop
End Sub

Private Sub Timer1_Timer()
Chat1.ChatSend "-VoiD Server By: Accesszero & DoS"
Chat1.ChatSend "type: /list to get a list of the mails"
Chat1.ChatSend "type: /sendX - X as index number."
End Sub
