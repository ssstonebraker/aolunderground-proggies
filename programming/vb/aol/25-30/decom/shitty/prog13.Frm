VERSION 4.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   Caption         =   "Scroller"
   ClientHeight    =   3705
   ClientLeft      =   2505
   ClientTop       =   1830
   ClientWidth     =   3615
   Height          =   4110
   Icon            =   "prog13.frx":0000
   Left            =   2445
   LinkTopic       =   "Form12"
   ScaleHeight     =   3705
   ScaleWidth      =   3615
   Top             =   1485
   Width           =   3735
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   1140
      TabIndex        =   16
      Top             =   960
      Width           =   915
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Ghost"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2280
      TabIndex        =   15
      Top             =   60
      Width           =   825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Quick Talker"
      ForeColor       =   &H000000FF&
      Height          =   2025
      Left            =   60
      TabIndex        =   8
      Top             =   1200
      Width           =   3525
      Begin VB.CommandButton Command5 
         Caption         =   "Clear it"
         Height          =   225
         Left            =   2400
         TabIndex        =   13
         Top             =   1770
         Width           =   1065
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Scroll it"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   1770
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1500
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   1500
         Width           =   1095
      End
      Begin RichtextLib.RichTextBox RichTextBox1 
         Height          =   1215
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   2143
         _StockProps     =   69
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"prog13.frx":0442
         RightMargin     =   3005.071
      End
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Tornado"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1320
      TabIndex        =   7
      Top             =   570
      Width           =   1005
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Box"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1320
      TabIndex        =   6
      Top             =   300
      Width           =   645
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1320
      TabIndex        =   5
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   195
      Left            =   2790
      TabIndex        =   3
      Top             =   900
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   300
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   900
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   3300
      Width           =   3525
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Lines To Scroll"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   3900
      Width           =   1965
   End
End
Attribute VB_Name = "Form12"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
If Text1.text = "" Then
MsgBox "Umm tell me how many lines to scroll"
End If
If room% = 0 Then
MsgBox "Room could not be found"
Exit Sub
End If
View% = FindChildByClass(room%, "_AOL_View")
If View% = 0 Then
MsgBox "Room could not be verified"
Exit Sub
End If
Z = aolhwnd = room%
hChild = GetWindow(room%, GW_CHILD)
hchild55 = hChild
roomz = GetClass(hChild)
Label1.Caption = roomz
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
rooms = GetClass(hChild)
Label1.Caption = rooms
Y = aolhwnd = hChild
Loop Until Label1.Caption = "_AOL_Edit"

Do: DoEvents
hchild55 = GetWindow(hchild55, GW_HWNDNEXT)
ico = GetClass(hchild55)
Label1.Caption = ico
o = aolhwnd = hchild55
Loop Until Label1.Caption = "_AOL_Icon"

If Option1.Value = True Then
X = 0
Y = 0
i = 0
Do: DoEvents
X = X + 1
Y = Y + 1
Call sendtext(hChild, "@@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 6
Call sendtext(hChild, "@@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 6
Call sendtext(hChild, "@@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 6
Call sendtext(hChild, "@@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ @@@@@@@@@@@@@ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 6
Pause (2)
If Y > 5 Then
Call sendtext(hChild, "-=Ðê‡ç‡Ðê™=- Haz scrolled " & i & " lines so far")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Pause (2)
Y = 0
End If
Loop Until X = Text1.text
End If
If Option2.Value = True Then
X = 0
Y = 0
i = 0
Do: DoEvents
X = X + 1
Y = Y + 1
Call sendtext(hChild, "ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 5
Call sendtext(hChild, "ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 5
Call sendtext(hChild, "ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 5
Call sendtext(hChild, "ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ŽŽŽŽŽŽŽŽ ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
i = i + 5
Pause (2)
If Y > 5 Then
Call sendtext(hChild, "-=Ðê‡ç‡Ðê™=- Haz scrolled " & i & " lines so far")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Pause (2)
Y = 0
End If
Loop Until X = Text1.text
End If

If Option3.Value = True Then

Call sendtext(hChild, "| GÕÐ§ßÃÑE by Deicide402™ |                                   | ™GÕÐ§ßÃÑE by Deicide402 |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| 2™GÕÐ§ßÃÑE by Deicide40 |                                   | 02™GÕÐ§ßÃÑE by Deicide4 |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| 402™GÕÐ§ßÃÑE by Deicide |                                   | e402™GÕÐ§ßÃÑE by Deicid |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| de402™GÕÐ§ßÃÑE by Deici |                                   | ide402™GÕÐ§ßÃÑE by Deic |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Pause (2)

Call sendtext(hChild, "| cide402™GÕÐ§ßÃÑE by Dei |                                   | icide402™GÕÐ§ßÃÑE by De |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| eicide402™GÕÐ§ßÃÑE by D |                                   | Deicide402™GÕÐ§ßÃÑE by  |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "|  Deicide402™GÕÐ§ßÃÑE by |                                    | y Deicide402™GÕÐ§ßÃÑE b |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| by Deicide402™GÕÐ§ßÃÑE  |                                    |  by Deicide402™GÕÐ§ßÃÑE |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Pause (2)

Call sendtext(hChild, "| E by Deicide402™GÕÐ§ßÃÑ |                                    | ÑE by Deicide402™GÕÐ§ßÃ |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| ÃÑE by Deicide402™GÕÐ§ß |                                    | ßÃÑE by Deicide402™GÕÐ§ |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| §ßÃÑE by Deicide402™GÕÐ |                                    | Ð§ßÃÑE by Deicide402™GÕ |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "| ÕÐ§ßÃÑE by Deicide402™G |                                    | GÕÐ§ßÃÑE by Deicide402™ |")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
End If

If Option4.Value = True Then
num = Text1.text
Do: DoEvents
Counter = Counter + 1
Call sendtext(hChild, "                                                                                        ")
Do: DoEvents
Call SendCharNum(hchild55, 13)
a = AOLGetText(hChild)
Loop Until a = ""
If Counter = 4 Then
Counter = 1
Pause (2)
End If
Bob = Bob + 1
Loop Until Bob = Text1.text
End If
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "This scrolls something if you picked an option"
End Sub


Private Sub Command2_Click()
Form12.Hide
Form4.Show
End Sub


Private Sub Command3_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
If room% = 0 Then
MsgBox "Room could not be found"
Exit Sub
End If
View% = FindChildByClass(room%, "_AOL_View")
If View% = 0 Then
MsgBox "Room could not be verified"
Exit Sub
End If
Z = aolhwnd = room%
hChild = GetWindow(room%, GW_CHILD)
roomz = GetClass(hChild)
Label1.Caption = roomz
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
rooms = GetClass(hChild)
Label1.Caption = rooms
Y = aolhwnd = hChild
Loop Until Label1.Caption = "_AOL_Edit"

Call sendtext(hChild, "              _________          ")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "             /   ____     \                            __         _______")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "           /   /        \     \   ___                  \  \       /  _____  \")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "        _|    |_        \     |  \__/                   \  \     /  /         \  |")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Pause (2)
Call sendtext(hChild, "        \       /        |     |   |  |      |\         |\  \  \   /  /        _ |  |_")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "         \    /          |     |   |  |      |  \       | \  \  \/  /          \      / ")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "          \ /            |     |   |  |      |    \    /  / /  /\  \            \   /")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "                         |     |   |  |      |  |\  \ /  / /  /   \  \           \ /")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Pause (2)
Call sendtext(hChild, "                         |     |  /__\    /  /  |    / /  /      \  \           ")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "                         |\_ _\ |___|_/_/_  \_/_/_/_____\_\_")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "                         ||__ _______________________ |")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""
Call sendtext(hChild, "                                                                                 Macro by: LiNx ")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hChild)
Loop Until a = ""


End Sub


Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Clears your text"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Option4_Click()
'//roll dice-sides 1 how to roll dice
End Sub


Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Umm this button...well nevermind"
End Sub

Private Sub Command4_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
mark = 1
If room% = 0 Then
MsgBox "Room could not be found"
Exit Sub
End If
View% = FindChildByClass(room%, "_AOL_View")
If View% = 0 Then
MsgBox "Room could not be verified"
Exit Sub
End If
Z = aolhwnd = room%
hChild = GetWindow(room%, GW_CHILD)
hchild2 = hChild
rooms = GetClass(hChild)
Label1.Caption = rooms

Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label1.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label1.Caption = "_AOL_Icon"

Do: DoEvents
hchild2 = GetWindow(hchild2, GW_HWNDNEXT)
rooms = GetClass(hchild2)
Label1.Caption = rooms
Y = aolhwnd = hchild2
Loop Until Label1.Caption = "_AOL_Edit"

Bob = Len(RichTextBox1.text)
If Bob = 0 Then
Exit Sub
End If
Let inptxt$ = RichTextBox1.text
inptxts = inptxt$
Do: DoEvents

If Bob < 92 Then
Let Nextchrr$ = Mid$(inptxt$, numspc%, Bob)
Call sendtext(hchild2, Nextchrr$)
Do: DoEvents
If Counter = 4 Then
Pause (1)
Counter = 0
End If
Call SendCharNum(hChild, 13)
a = AOLGetText(hchild2)
Loop Until a = ""
Exit Sub
End If
Let Nextchrr$ = Mid$(inptxt$, mark, 92)
mark = mark + 92
If Nextchrr$ = "" Then
bobo = 0
mark = 0
Exit Sub
End If
Counter = Counter + 1
If Counter = 5 Then
Pause (2)
Counter = 0
End If
Call sendtext(hchild2, Nextchrr$)
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hchild2)
Loop Until a = ""

Loop

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Scroll that text"
End Sub


Private Sub Command5_Click()
RichTextBox1.text = ""
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Nope De Nada here"
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Allows you to say something all at once"
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Yeah...pretty huh? It doesn't do anything"
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Says how many times it will take to scroll text"
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Another one of those boxes..."
End Sub


Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "This scrolls a whole bunch of @ over and over"
End Sub


Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "K this one makes a whole bunch of boxes"
End Sub


Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Don't know how to explain it..."
End Sub


Private Sub Option4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Send's Blank Lines to the chat room"
End Sub

Private Sub RichTextBox1_Change()
a = Len(RichTextBox1.text)
Label3.Caption = "Length = " & a
X = Fix(a / 92)
X = X + 1
Label4.Caption = "Will Scroll " & X & " Times"
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Text you wanna say...all at once"
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "How many times it scrolls"
End Sub


