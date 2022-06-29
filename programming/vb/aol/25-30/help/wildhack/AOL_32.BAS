Attribute VB_Name = "AOL_32"
'***************************DECLARES***********************

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Long, ByVal wFlag As Long) As Long



'*************************CONSTANTS************************

Global Stop_Busting_In

Sub ADD_AOL_LB(itm As String, lst As ListBox)
'Add a list of names to a VB ListBox
'This is usually called by another one of my functions

If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub


Sub aolclick(E1 As Integer)
'Clicks an AOL button with the given handle as E1

Exit Sub


do_wn = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
Pause 0.008
u_p = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
End Sub

Function aolhwnd()
'finds AOL's handle
a = FindWindow("AOL Frame25", vbNullString)
aolhwnd = a
End Function

Sub AOLSendMail(Person, subject, message)

'Opens an AOL Mail and fills it out to PERSON, with a
'subject of SUBJECT, and a message of MESSAGE.
'*****THIS DOES NOT SEND THE MAIL  !! ******

aol% = FindWindow("AOL Frame25", vbNullString)
If aol% = 0 Then
    MsgBox "Must Be Online"
    Exit Sub
End If
Call RunMenuByString(aol%, "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)

'AOLIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub

Sub countnewmail()
'Counts your new mail...Mail doesn't have to be open

a = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(a, "Read &New Mail")

AO% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
bb% = FindChildByClass(AO%, "MDIClient")
arf = FindChildByTitle(bb%, "New Mail")
If arf <> 0 Then Exit Do
Loop


Hand% = FindChildByClass(arf, "_AOL_TREE")
buffer = SendMessage(Hand%, LB_GETCOUNT, 0, 0)
If buffer > 1 Then
MsgBox "You have " & buffer & " messages in your E-Mailbox."
End If
If buffer = 1 Then
MsgBox "You have one message in your E-Mailbox."
End If
If buffer < 1 Then
MsgBox "You have zero messages in your E-Mailbox"
End If

End Sub


Function findchatroom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = FindChildByClass(firs%, "_AOL_Edit")
listere% = FindChildByClass(firs%, "_AOL_View")
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone

firs% = GetWindow(mdi%, GW_CHILD)
Do: DoEvents
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "_AOL_Edit")
listere% = FindChildByClass(firs%, "_AOL_View")
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then Exit Do

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = FindChildByClass(firs%, "_AOL_Edit")
listere% = FindChildByClass(firs%, "_AOL_View")
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
l = l + 1
If l = 100 Then GoTo begis
Loop

bone:
room% = firs%
findchatroom = room%
Exit Function
begis:
findchatroom = 0
End Function

Function findcomposemail()
'Finds the Compose mail window's handle

Dim bb As Integer
Dim dis_win As Integer

dis_win = FindChildByClass(aolhwnd(), "AOL Child")

begin_find_composemail:

bb = FindChildByTitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "To:")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Subject:")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Send" & Chr(13) & "Later")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Attach")
    If bb <> 0 Then Let countt = countt + 1

bb = FindChildByTitle(dis_win, "Address" & Chr(13) & "Book")
    If bb <> 0 Then Let countt = countt + 1

If countt = 6 Then
  findcomposemail = dis_win
  Exit Function
End If
Let countt = 0
dis_win = GetNextWindow(dis_win, 2)
If dis_win = GetWindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function


Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
r_backwards = newsent$

End Function

Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "|V|"
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "º"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function

Function r_hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
r_hacker = newsent$

End Function

Function r_same(strr As String)
'Returns the strin the same
Let r_same = Trim(strr)

End Function

Function r_spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
r_spaced = newsent$

End Function


Sub Sendclick(Handle)
'Clicks something
X% = SendMessage(Handle, WM_LBUTTONDOWN, 0, 0&)
Pause 0.05
X% = SendMessage(Handle, WM_LBUTTONUP, 0, 0&)
End Sub

Sub sendtext(handl As Integer, msgg As String)
'Sends msgg to handl
send_txt = SendMessageByString(handl, WM_SETTEXT, 0, msgg)
End Sub

Sub showaolwins()
'Shows all AOL Windows
fc = FindChildByClass(aolhwnd(), "AOL Child")
req = ShowWindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = GetNextWindow(faa, 2)
res = ShowWindow(faa, 1)
DoEvents
Loop Until faf = faa


End Sub

Sub StayOnTop(frm As Form)
'Allows a window to stay on top
Dim success%
success% = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Sub

Function trim_null(wstr As String)
'Trims null characters from a string
wstr = Trim(wstr)
Do Until xx = Len(wstr)
Let xx = xx + 1
Let this_chr = Asc(Mid$(wstr, xx, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(wstr, xx, 1)
Loop
trim_null = wordd

End Function

Sub waitforok()
'Waits for the AOL OK messages that popup up
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function windowcaption(hWndd As Integer)
'Gets the caption of a window
Dim WindowText As String * 255
Dim getWinText As Integer
getWinText = GetWindowText(hWndd, WindowText, 255)
windowcaption = (WindowText)
End Function

