VERSION 4.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IM stuff"
   ClientHeight    =   2655
   ClientLeft      =   2370
   ClientTop       =   1200
   ClientWidth     =   6675
   Height          =   3060
   Left            =   2310
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Top             =   855
   Width           =   6795
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Height          =   195
      Left            =   4950
      TabIndex        =   13
      Top             =   1890
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "More"
      Height          =   195
      Left            =   1530
      TabIndex        =   12
      Top             =   2400
      Width           =   705
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00000000&
      Caption         =   "Swirly"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4110
      TabIndex        =   11
      Top             =   1680
      Width           =   795
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Blank IM"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4110
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "IM is open"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3930
      TabIndex        =   9
      Top             =   120
      Width           =   1125
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Funky Letters"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4110
      TabIndex        =   8
      Top             =   3030
      Width           =   1305
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Medium"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4110
      TabIndex        =   7
      Top             =   1050
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Fastest"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4110
      TabIndex        =   6
      Top             =   810
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   195
      Left            =   2970
      TabIndex        =   3
      Top             =   2400
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Scrolling( Regular)"
      ForeColor       =   &H000000FF&
      Height          =   945
      Left            =   3900
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
      Begin VB.OptionButton Option6 
         BackColor       =   &H00000000&
         Caption         =   "Crap"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   510
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Scrolling ( Typing)"
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   3900
      TabIndex        =   15
      Top             =   540
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Miscellaneous"
      ForeColor       =   &H000000FF&
      Height          =   1515
      Left            =   3900
      TabIndex        =   16
      Top             =   2490
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Windows"
      Height          =   345
      Left            =   840
      TabIndex        =   4
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1830
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin RichtextLib.RichTextBox RichTextBox1 
      Height          =   1905
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   3705
      _Version        =   65536
      _ExtentX        =   6535
      _ExtentY        =   3360
      _StockProps     =   69
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         name            =   "Small Fonts"
         charset         =   0
         weight          =   400
         size            =   6.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ScrollBars      =   2
      TextRTF         =   $"PROG3.frx":0000
      RightMargin     =   3375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command3_Click()
If Form3.Height < 3200 Then
Form3.Height = 4150
Form3.Width = 5715
Command3.Caption = "Less"
Exit Sub
End If

If Form3.Height > 3200 Then
Form3.Width = 3965
Form3.Height = 3050
Command3.Caption = "More"
Exit Sub
End If

End Sub

Private Sub Command4_Click()
If Text1.text = "" Then
    MsgBox "Please Enter a name first"
    Exit Sub
End If

mdicc = 0
uik = 0
frankie = 0
bobb = 0


If Option5.Value = True Then
bobb = 1
End If

If Option6.Value = True Then
bobb = 1
End If

If bobb = 0 Then
    MsgBox "Chose a option before you click on this"
    Exit Sub
End If

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mdic = GetChildCount(MDI%)

If Check1.Value <> 0 Then
    hChild = GetWindow(MDI%, GW_CHILD)
    rooms = GetClass(hChild)
    Label2.Caption = rooms
    If Label2.Caption = "AOL Child" Then 'bbb
        ax = AOLGetText(hChild)
        aaa = Len(Text1.text)
        aab = Len(ax)
        xxx = aab - aaa
        xxx = xxx + 1
        Itext = Mid(ax, xxx, aaa)
     End If                                        'bbb ended
            If Itext = Text1.text Then 'aaa
            IMs% = hChild
            Else
            Do: DoEvents
                Do: DoEvents
                    mdicc = mdicc + 1
                    If mdicc > mdic Then
                    MsgBox "The IM could not be found, make certain you have spelled the name right, and the box is open"
                    Exit Sub
                    End If
                    hChild = GetWindow(hChild, GW_HWNDNEXT)
                    bob = GetClass(hChild)
                    Label2.Caption = bob
                    Y = aolhwnd = hChild
                Loop Until Label2.Caption = "AOL Child"
                ax = AOLGetText(hChild)
                aaa = Len(Text1.text)
                aab = Len(ax)
                xxx = aab - aaa
                xxx = xxx + 1
                If xxx > 0 Then
                Itext = Mid(ax, xxx, aaa)
                End If
            Loop Until Itext = Text1.text
        End If 'aaa ended
Else
Call RunMenuByString(AOL%, "Send an instant Message")
MDI% = FindChildByClass(AOL%, "MDIClient")
Do: DoEvents
IMs% = FindChildByTitle(MDI%, "Send Instant Message")
Loop Until IMs% <> 0
Call sendtext(IMs%, "funky")
End If
X = ShowWindow(IMs%, SW_HIDE)


'To who?

If Check1.Value <> 0 Then
Else
hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label2.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Edit"
'Setting Text
Do: DoEvents
Call AOLSetText(hChild, Text1.text)
EText = AOLGetText(hChild)
Loop Until EText = Text1.text
End If

Counter = 0
'Getting text in right format
a = Len(RichTextBox1.text)


If Option5.Value = True Then
imstring = RichTextBox1.text
IMLength = Len(imstring)
If IMLength = 0 Then
MsgBox "Umm enter in some text first!"
Exit Sub
End If
ietext = Chr$(13) + Chr$(10) + "</html>" & RichTextBox1.text + Chr$(13) + Chr$(10)
Do: DoEvents
IMLength = IMLength - 1
charx = Mid(imstring, 1, IMLength)
ietext = ietext + "</html>" & charx + Chr$(13) + Chr$(10)
Loop Until IMLength = 1
imstring = RichTextBox1.text
Do: DoEvents
IMLength = IMLength + 1
charx = Mid(imstring, 1, IMLength)
ietext = ietext + "</html>" & charx + Chr$(13) + Chr$(10)
Loop Until Len(charx) = Len(RichTextBox1.text)
End If

If Option6.Value = True Then
hu = Len(RichTextBox1.text)
fu = RichTextBox1.text
hu = hu + 15
Dim Byu As Integer
Byu = 500 / hu
asdf = 1
Stringy = Chr$(13) & Chr$(10) & "</html>-=|}" & fu & "{|=-" & Chr$(13) + Chr$(10)
asd = asdf + hu
ietext = Stringy

Do: DoEvents

asdff = asdff + 1
asdf = asdf + 1
asd = asdf + hu
If asdf > hu Then
asdf = 1
asd = hu
End If
If asd = hu Then
mystring = "</html>-=|}" & fu & "{|=-" & Chr$(13) + Chr$(10)
Else
hui = asd - hu
bobe = Mid(fu, 1, hui)
bobee = Mid(fu, hui, hu)
mystring = "</html>-=|}" & bobe & bobee & "{|=-" & Chr$(13) + Chr$(10)
End If
ietext = ietext + mystring
shitty = shitty + 1
Loop Until shitty = Byu
End If

hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label2.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "RICHCNTL"
If Check1.Value <> 0 Then
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "RICHCNTL"
End If
stxt = hChild


Do: DoEvents
Call AOLSetText(hChild, ietext)
nText = AOLGetText(hChild)
Loop Until nText = ietext


'Send button
hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label2.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"

'Normal size button
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
nsi = hChild

'Size up button
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
upt = hChild

'Bold Button
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
bol = hChild

'Italics button
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"

'Underline button
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"

'send button
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"

Label1.Caption = "Sent!"
Call Sendclick(hChild)
If Check1.Value <> 0 Then
Y = ShowWindow(IMs%, SW_SHOW)
End If
Pause (1)
Bobzx% = FindChildByTitle(MDI%, "funky")
If Bobzx% <> 0 Then
SendMessage (Bobzx%), WM_CLOSE, 0, 0
End If
End Sub


Private Sub Command1_Click()
If Text1.text = "" Then
    MsgBox "Please Enter a name first"
    Exit Sub
End If

mdicc = 0
uik = 0
frankie = 0
bobb = 0

If Option1.Value = True Then
    bobb = 1 + bobb
End If

If Option2.Value = True Then
    bobb = 1 + bobb
End If

If Option3.Value = True Then
    bobb = 1 + bobb
End If

If Option4.Value = True Then
    bobb = 1 + bobb
End If

If bobb = 0 Then
    MsgBox "Chose a option before you click on this"
    Exit Sub
End If

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mdic = GetChildCount(MDI%)

If Check1.Value <> 0 Then
    hChild = GetWindow(MDI%, GW_CHILD)
    rooms = GetClass(hChild)
    Label2.Caption = rooms
    If Label2.Caption = "AOL Child" Then 'bbb
        ax = AOLGetText(hChild)
        aaa = Len(Text1.text)
        aab = Len(ax)
        xxx = aab - aaa
        xxx = xxx + 1
        Itext = Mid(ax, xxx, aaa)
     End If                                        'bbb ended
            If Itext = Text1.text Then 'aaa
            IMs% = hChild
            Else
            Do: DoEvents
                Do: DoEvents
                    mdicc = mdicc + 1
                    If mdicc > mdic Then
                    MsgBox "The IM could not be found, make certain you have spelled the name right, and the box is open"
                    Exit Sub
                    End If
                    hChild = GetWindow(hChild, GW_HWNDNEXT)
                    bob = GetClass(hChild)
                    Label2.Caption = bob
                    Y = aolhwnd = hChild
                Loop Until Label2.Caption = "AOL Child"
                ax = AOLGetText(hChild)
                aaa = Len(Text1.text)
                aab = Len(ax)
                xxx = aab - aaa
                xxx = xxx + 1
                If xxx > 0 Then
                Itext = Mid(ax, xxx, aaa)
                End If
            Loop Until Itext = Text1.text
        End If 'aaa ended
Else
Call RunMenuByString(AOL%, "Send an instant Message")
MDI% = FindChildByClass(AOL%, "MDIClient")
Do: DoEvents
IMs% = FindChildByTitle(MDI%, "Send Instant Message")
Loop Until IMs% <> 0
Call sendtext(IMs%, "funky")
End If
X = ShowWindow(IMs%, SW_HIDE)

'To who?

If Check1.Value <> 0 Then
Else
hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label2.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Edit"
'Setting Text
Do: DoEvents
Call AOLSetText(hChild, Text1.text)
EText = AOLGetText(hChild)
Loop Until EText = Text1.text
End If

Counter = 0
'Getting text in right format
a = Len(RichTextBox1.text)

Do: DoEvents
numspc% = numspc% + 1
Let nextchr$ = Mid$(RichTextBox1.text, numspc%, 1)
If Option1.Value = True Then
nextchrs$ = nextchr$ + "</html>"
End If

If Option2.Value = True Then
nextchrs$ = nextchr$ + "</html></html>"
End If

If Option4.Value = True Then
Teixt$ = "<a href"">"
End If

If Option3.Value = True Then
frankie = frankie + 1
If frankie = 4 Then
frankie = 1
End If

If frankie = 1 Then
nextchrs$ = nextchr$ + "<Sub>"
End If

If frankie = 2 Then
nextchrs$ = nextchr$ + "</sub>"
End If

If frankie = 3 Then
nextchrs$ = nextchr$ + "<sub>"
End If
End If

Teixt$ = Teixt$ + nextchrs$
Counter = Counter + 1
Loop Until Counter = a

'getting text part
hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label2.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "RICHCNTL"
If Check1.Value <> 0 Then
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "RICHCNTL"
End If

Do: DoEvents
Call AOLSetText(hChild, Teixt$)
nText = AOLGetText(hChild)
Loop Until nText = Teixt$

'Send button
hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label2.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
bob = GetClass(hChild)
Label2.Caption = bob
Y = aolhwnd = hChild
Loop Until Label2.Caption = "_AOL_Icon"
Label1.Caption = "Sent!"
Call Sendclick(hChild)
If Check1.Value <> 0 Then
Y = ShowWindow(IMs%, SW_SHOW)
End If
Pause (2)
Bobzx% = FindChildByTitle(MDI%, "funky")
If Bobzx% <> 0 Then
SendMessage (Bobzx%), WM_CLOSE, 0, 0
End If
End Sub

Private Sub Command2_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
AOL% = FindWindow("AOL FRame25", vbNullString)
Call RunMenuByString(AOL%, "send an instant message")

End Sub

Private Sub Form_Load()
Call StayOnTop(Form3)

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call sendtext(Label3.Caption, Text1.text)
Call SendCharNum(Label3.Caption, 13)
Text1.text = ""
End If
End Sub


Private Sub Timer1_Timer()
roomm% = FindChildByClass(MDI%, "AOL Child")
child = FindChildByClass(roomm%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&) 'Get some text

trimspace$ = Space$(GetTrim) 'Setup a string that is as
                    'long as the chat text
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)
                   'Place the text in our newly set up
                   'holding area
theview$ = trimspace$  'Rename our variable

Label1.Caption = theview$ 'Place the chat text in our text box
Text1.SelStart = Len(theview$) 'Scroll our text box
           
End Sub



Private Sub RichTextBox1_Change()

If Option1.Value = True Then X = 1
If Option2.Value = True Then X = 1
If Option3.Value = True Then X = 1
If Option4.Value = True Then X = 1
If Option5.Value = True Then X = 1
If Option6.Value = True Then X = 1
If X <> 1 Then
MsgBox "Click on More and select a option before you type in the text!"
End If

a = Len(RichTextBox1.text)
Label1.Caption = "Length = " & a
blah = RichTextBox1.text




If Option1.Value = True Then
If a = 128 Then
MsgBox "Using this option...the IM can not hold anymore text! :( sorry...."
End If
End If

If Option2.Value = True Then
If a = 69 Then
MsgBox "Using this option...the IM can not hold anymore stuff..damn I know"
End If
End If

If Option3.Value = True Then
If a / 3 * 10 > 498 Then
MsgBox "You are going past the limit the IM will hold, stop typing text in! :)"
End If
End If

If Option4.Value = True Then
MsgBox "Ummm hey with this option, umm you don't have ta send any text"
End If

If Option5.Value = True Then
If a = 24 Then
MsgBox "This option can only hold 24 characters"
End If
End If
End Sub

