VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Mailer By Mission"
   ClientHeight    =   3420
   ClientLeft      =   2805
   ClientTop       =   4500
   ClientWidth     =   4125
   Icon            =   "MMer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "&Options"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A&dd Room"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   1320
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   2040
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   1005
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "MMer.frx":1272
      Top             =   2400
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Message"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "People on MM"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label4 
      Caption         =   "% done"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "of"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Label13.Caption = "0" Then
MsgBox "You must have at least one person to Mass Mail.", 64
Exit Sub
End If
If UserSN = "" Then
MsgBox "Please sign on before using this feature.", 64
Exit Sub
End If
If Form2.Option1.Value = True And Command1.Caption = "&Start" Then
Command1.Caption = "&Stop"
IMsOff
Call MailWaitForLoadNew
TimeOut (1)
Label11.Caption = MailCountNew
Call MailOpenEmailNew(0)
TimeOut (5)
Timer1.Enabled = True
Timer5.Enabled = True
Exit Sub
End If
If Command1.Caption = "&Stop" Then
Command1.Caption = "&Start"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Label9.Caption = "0"
Label11.Caption = "0"
Label15.Caption = "0"
Call PercentBar(Picture1, Val(Label9.Caption), Val(Label11.Caption))
Box% = FindChildByTitle(AOLMDI, UserSN & "'s Online Mailbox")
Box2% = FindChildByTitle(AOLMDI, "Incoming/Saved Mail")
If Box% Then
Call CloseWindow(Box%)
End If
If Box2% Then
Call CloseWindow(Box2%)
End If
Call CloseWindow(Box%)
TimeOut (1)
Call CloseOpenMails
TimeOut (1)
IMsOn
Exit Sub
End If
If Form2.Option2.Value = True And Command1.Caption = "&Start" Then
Command1.Caption = "&Stop"
IMsOff
Call MailWaitForLoadOld
TimeOut (1)
Label11.Caption = MailCountOld
Call MailOpenEmailOld(0)
TimeOut (5)
Timer1.Enabled = True
Timer5.Enabled = True
Exit Sub
End If
If Form2.Option3.Value = True And Command1.Caption = "&Start" Then
Command1.Caption = "&Stop"
IMsOff
Call MailWaitForLoadSent
TimeOut (1)
Label11.Caption = MailCountSent
Call MailOpenEmailSent(0)
TimeOut (5)
Timer1.Enabled = True
Timer5.Enabled = True
Exit Sub
End If
If Form2.Option4.Value = True And Command1.Caption = "&Start" Then
Command1.Caption = "&Stop"
IMsOff
MailOpenFlash
Call MailWaitForLoadFlash
TimeOut (1)
Label11.Caption = MailCountFlash
Call MailOpenEmailFlash(0)
TimeOut (5)
Timer1.Enabled = True
Timer5.Enabled = True
Exit Sub
End If
End Sub
Private Sub Command2_Click()
Dim strAdd As String
strAdd = InputBox("What is the screen name of the person you would like to add to the Mass Mail?", "Add Person")
If TrimSpaces(strAdd) = "" Then
Exit Sub
Else
List1.AddItem (strAdd)
Label13.Caption = List1.ListCount
End If
End Sub
Private Sub Command3_Click()
If List1.ListCount = 0 Then
MsgBox "You have nobody on the list to remove.", 64
Exit Sub
End If
If List1.SelCount = 0 Then
MsgBox "Please highlight who you would like to remove.", 64
Exit Sub
Else
List1.RemoveItem (List1.ListIndex)
End If
Label13.Caption = List1.ListCount
End Sub
Private Sub Command4_Click()
If FindRoom = 0 Then
Exit Sub
Else
Call AddRoomToListbox(List1)
End If
Label13.Caption = List1.ListCount
End Sub
Private Sub Command5_Click()
Dim l0160 As Variant
l0160 = MsgBox("Are you sure you want to clear the list?", 36)
If l0160 = 6 Then
List1.Clear
Label13 = 0
End If
If l0160 = 7 Then
Exit Sub
End If
End Sub
Private Sub Command6_Click()
Form2.Show
End Sub
Private Sub Form_Load()
Call StayOnTop(Me)
Call PercentBar(Picture1, Val(Label9.Caption), Val(Label11.Caption))
Text1.Text = "Enjoy this Mass Mail!"
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Timer1_Timer()
If List1.ListCount = 0 Then
Call CloseOpenMails
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
TimeOut (1)
IMsOn
Command1.Caption = "&Start"
Label9.Caption = "0"
Label11.Caption = "0"
Label15.Caption = "0"
Call PercentBar(Picture1, Val(Label9.Caption), Val(Label11.Caption))
MsgBox "You must have at least one person to Mass Mail.", 64
Exit Sub
End If
If FindSendWindow = 0 Then
Exit Sub
End If
Label15.Caption = Val(Label9.Caption + 1) / Val(Label11.Caption) * 100
If Form2.Check1.Value = 1 Then
Call MailForward(AddListToString(List1), Text1.Text & Chr(13) & Chr(13) & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "><-==(`(` " & "<A HREF=" & Chr(34) & "http://members.tripod.com/~ShaOLinXGroup/Setup.exe" & Chr(34) & ">Icy Hot 2.0 For AOL 4.0</A><FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "> ')')==-></FONT>" & Chr(13) & "<-==(`(` Mail # " & Label9.Caption + 1 & " of " & Label11.Caption & " - " & Left(Label15.Caption, 4) & "% Done ')')==->", True)
End If
If Form2.Check1.Value = 0 Then
Call MailForward(AddListToString(List1), Text1.Text & Chr(13) & Chr(13) & "<FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "><-==(`(` " & "<A HREF=" & Chr(34) & "http://members.tripod.com/~ShaOLinXGroup/Setup.exe" & Chr(34) & ">Icy Hot 2.0 For AOL 4.0</A><FONT COLOR=" & Chr(34) & "#FF0000" & Chr(34) & "> ')')==-></FONT>" & Chr(13) & "<-==(`(` Mail # " & Label9.Caption + 1 & " of " & Label11.Caption & " - " & Left(Label15.Caption, 4) & "% Done ')')==->", False)
End If
Label9.Caption = Label9.Caption + 1
Call PercentBar(Picture1, Val(Label9.Caption), Val(Label11.Caption))
Timer2.Enabled = True
Timer3.Enabled = True
Timer1.Enabled = False
End Sub
Private Sub Timer2_Timer()
If FindSendWindow <> 0 Then
Exit Sub
End If
If Form2.Option1.Value = True Then
Call MailOpenEmailNew(Val(Label9.Caption))
End If
If Form2.Option2.Value = True Then
Call MailOpenEmailOld(Val(Label9.Caption))
End If
If Form2.Option3.Value = True Then
Call MailOpenEmailSent(Val(Label9.Caption))
End If
If Form2.Option4.Value = True Then
Call MailOpenEmailFlash(Val(Label9.Caption))
End If
TimeOut (5)
Timer1.Enabled = True
Timer2.Enabled = False
End Sub
Private Sub Timer3_Timer()
If Label9.Caption = Label11.Caption And Val(Label11.Caption) > 0 Then
Timer1.Enabled = False
Timer2.Enabled = False
Box% = FindChildByTitle(AOLMDI, UserSN & "'s Online Mailbox")
Box2% = FindChildByTitle(AOLMDI, "Incoming/Saved Mail")
If Box% Then
Call CloseWindow(Box%)
End If
If Box2% Then
Call CloseWindow(Box2%)
End If
TimeOut (1)
IMsOn
Command1.Caption = "&Start"
Label9.Caption = "0"
Label11.Caption = "0"
Label15.Caption = "0"
Call PercentBar(Picture1, Val(Label9.Caption), Val(Label11.Caption))
Timer4.Enabled = False
Timer5.Enabled = False
Timer3.Enabled = False
End If
AoError% = FindChildByTitle(AOLMDI, "Error")
If AoError% Then
TimeOut (0.5)
Timer4.Enabled = True
Timer3.Enabled = False
End If
End Sub
Private Sub Timer4_Timer()
If ErrorNameCount = 0 Then
Timer4.Enabled = False
End If
If ErrorNameCount = 1 Then
Call DeleteItem(List1, ErrorName(1))
End If
If ErrorNameCount = 2 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
End If
If ErrorNameCount = 3 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
End If
If ErrorNameCount = 4 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
End If
If ErrorNameCount = 5 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
End If
If ErrorNameCount = 6 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
End If
If ErrorNameCount = 7 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
End If
If ErrorNameCount = 8 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
End If
If ErrorNameCount = 9 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
End If
If ErrorNameCount = 10 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
End If
If ErrorNameCount = 11 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
End If
If ErrorNameCount = 12 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
End If
If ErrorNameCount = 13 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
End If
If ErrorNameCount = 14 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
End If
If ErrorNameCount = 15 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
End If
If ErrorNameCount = 16 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
Call DeleteItem(List1, ErrorName(16))
End If
If ErrorNameCount = 17 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
Call DeleteItem(List1, ErrorName(16))
Call DeleteItem(List1, ErrorName(17))
End If
If ErrorNameCount = 18 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
Call DeleteItem(List1, ErrorName(16))
Call DeleteItem(List1, ErrorName(17))
Call DeleteItem(List1, ErrorName(18))
End If
If ErrorNameCount = 19 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
Call DeleteItem(List1, ErrorName(16))
Call DeleteItem(List1, ErrorName(17))
Call DeleteItem(List1, ErrorName(18))
Call DeleteItem(List1, ErrorName(19))
End If
If ErrorNameCount = 20 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
Call DeleteItem(List1, ErrorName(16))
Call DeleteItem(List1, ErrorName(17))
Call DeleteItem(List1, ErrorName(18))
Call DeleteItem(List1, ErrorName(19))
Call DeleteItem(List1, ErrorName(20))
End If
If ErrorNameCount > 20 Then
Call DeleteItem(List1, ErrorName(1))
Call DeleteItem(List1, ErrorName(2))
Call DeleteItem(List1, ErrorName(3))
Call DeleteItem(List1, ErrorName(4))
Call DeleteItem(List1, ErrorName(5))
Call DeleteItem(List1, ErrorName(6))
Call DeleteItem(List1, ErrorName(7))
Call DeleteItem(List1, ErrorName(8))
Call DeleteItem(List1, ErrorName(9))
Call DeleteItem(List1, ErrorName(10))
Call DeleteItem(List1, ErrorName(11))
Call DeleteItem(List1, ErrorName(12))
Call DeleteItem(List1, ErrorName(13))
Call DeleteItem(List1, ErrorName(14))
Call DeleteItem(List1, ErrorName(15))
Call DeleteItem(List1, ErrorName(16))
Call DeleteItem(List1, ErrorName(17))
Call DeleteItem(List1, ErrorName(18))
Call DeleteItem(List1, ErrorName(19))
Call DeleteItem(List1, ErrorName(20))
Call DeleteItem(List1, ErrorName(21))
End If
Label13.Caption = List1.ListCount
TimeOut (3)
Error% = FindChildByTitle(AOLMDI, "Error")
Call CloseWindow(Error%)
TimeOut (1)
Call ClickSendAfterError(AddListToString(List1))
If FindSendWindow = 0 Then
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
Call CloseWindow(mailwin%)
End If
Timer4.Enabled = False
End Sub
Private Sub Timer5_Timer()
If FindForwardWindow <> 0 And FindSendWindow = 0 Then
Call ClickForward
End If
If FindForwardWindow = 0 And FindSendWindow <> 0 Then
mailwin% = FindChildByTitle(AOLMDI(), "fwd: ")
Call CloseWindow(mailwin%)
End If
End Sub
