'***************************DECLARES***********************
Declare Sub setwindowtext Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Sub CloseWindow Lib "User" (ByVal hWnd As Integer)
Declare Sub MoveWindow Lib "User" (ByVal hWnd As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Declare Sub bringwindowtotop Lib "User" (ByVal hWnd As Integer)
Declare Function GetCursor Lib "User" () As Integer
Declare Function EnableMenuItem Lib "User" (ByVal hMenu As Integer, ByVal wIDEnableItem As Integer, ByVal wEnable As Integer) As Integer
Declare Function DestroyMenu Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer


Declare Function getwindowtask Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function enablewindow Lib "User" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function GetActiveWindow Lib "User" () As Integer
Declare Function destroywindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function createwindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)

Declare Function setactivewindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function setfocusapi Lib "User" Alias "SetFocus" (ByVal hWnd As Integer) As Integer
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function DeleteMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function Getwindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function getwindowtext Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function getfocus Lib "User" () As Integer
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetMenu% Lib "User" (ByVal hWnd%)
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuItemID% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuState Lib "User" (ByVal hMenu As Integer, ByVal wId As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetSubMenu% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function findwindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function gettopwindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function sendmessagebynum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function sendmessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function sendmessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function AOLGetList% Lib "311.Dll" (ByVal Index%, ByVal Buf$)
Declare Function getnames Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByValp2$) As Integer

Declare Function findchildbytitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal title$)
Declare Function findchildbyclass% Lib "vbwfind.dll" (ByVal Parent%, ByVal title$)
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal Cx%, ByVal Cy%, ByVal f%) As Integer

'*************************CONSTANTS************************
Global Const SW_HIDE = 0
Global Const SW_SHOW = 5
Global Const WM_USER = &H400
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_DELETESTRING = (WM_USER + 3)
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LB_FINDSTRING = (WM_USER + 16)
Global Const LB_GETtext = (WM_USER + 10)
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const GW_HWNDNEXT = 2
Global Const GW_CHILD = 5
Global Const WM_CLOSE = &H10
Global Const WM_DESTROY = &H2
Global Const WM_SETTEXT = &HC
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_CHAR = &H102
Global Const WM_COMMAND = &H111
Global Const MB_ICONSTOP = 16
Global Const MB_OK = 0
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Global Const MF_BYCOMMAND = &H0
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const WM_gettext = &HD
Global Const CB_ADDSTRING = (WM_USER + 3)
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_DELETESTRING = (WM_USER + 4)
Global Program_Title
Global da_num#
Global daaf%
Global entr
Global Stop_Busting_In
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Sub ADD_AOL_LB (itm As String, lst As ListBox)
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

Sub AddRoom (lst As ListBox)
'This calls a function in 311.dll that retreives the names
'from the AOL listbox.
'PLEASE NOTE THE FOLLOWING:
'1)  I don't support this dll..its hacked and illegal
'2)  This only works on 16 bit versions of AOL
'3)  Its a good idea to bring the chat room to the top
'    of the AOL client before doing this.  Sometimes it
'    gets text from other AOL listboxes


For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_addr
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)

ADD_AOL_LB namez$, lst
Next Index%
end_addr:

End Sub

Sub aolclick (E1 As Integer)
'Clicks an AOL button with the given handle as E1

Exit Sub


do_wn = sendmessagebynum(E1, WM_LBUTTONDOWN, 0, 0&)
pause .008
u_p = sendmessagebynum(E1, WM_LBUTTONUP, 0, 0&)
End Sub

Function aolhwnd ()
'finds AOL's handle
a = findwindow("AOL Frame25", 0&)
aolhwnd = a
End Function

Sub AOLSendMail (PERSON, SUBJECT, MESSAGE)

'Opens an AOL Mail and fills it out to PERSON, with a
'subject of SUBJECT, and a message of MESSAGE.
'*****THIS DOES NOT SEND THE MAIL  !! ******

AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then
    MsgBox "Must Be Online"
    Exit Sub
End If
Call RunMenuByString(AOL%, "Compose Mail")

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(mdi%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = Getwindow(subjt%, 2)
mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, PERSON)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, SUBJECT)
a = sendmessagebystring(mess%, WM_SETTEXT, 0, MESSAGE)

'AOLIcon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(mdi%, "Compose Mail")
erro% = findchildbytitle(mdi%, "Error")
aolw% = findwindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = sendmessage(aolw%, WM_CLOSE, 0, 0)
a = sendmessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
If erro% <> 0 Then
a = sendmessage(erro%, WM_CLOSE, 0, 0)
a = sendmessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub

Sub countnewmail ()
'Counts your new mail...Mail doesn't have to be open

a = findwindow("AOL Frame25", 0&)
Call RunMenuByString(a, "Read &New Mail")

AO% = findwindow("AOL Frame25", 0&)
Do: DoEvents
arf = findchildbytitle(AO%, "New Mail")
If arf <> 0 Then Exit Do
Loop


Hand% = findchildbyclass(arf, "_AOL_TREE")
buffer = sendmessagebynum(Hand%, LB_GETCOUNT, 0, 0)
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

Sub enableaolwins ()
'Enables all AOL Child Windows

Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 1)

fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa
End Sub

Function findchatroom ()
'Finds the handle of the AOL Chatroom by looking for a
'Window with a ListBox (Chat ScreenNames), Edit Box,
'(Where you type chat text), and an _AOL_VIEW.  If another
'AOL window is present that also has these 3 controls, it
'may find the wrong window.  I have never seen another AOL
'window with these 3 controls at once

AOL = findwindow("AOL Frame25", 0&)
If AOL = 0 Then Exit Function
b = findchildbyclass(AOL, "AOL Child")

start:
c = findchildbyclass(b, "_AOL_VIEW")
If c = 0 Then GoTo nextwnd
d = findchildbyclass(b, "_AOL_EDIT")
If d = 0 Then GoTo nextwnd
e = findchildbyclass(b, "_AOL_LISTBOX")
If e = 0 Then GoTo nextwnd
'We've found it
findchatroom = b
Exit Function

nextwnd:
b = getnextwindow(b, 2)
If b = Getwindow(b, GW_HWNDLAST) Then Exit Function
GoTo start


End Function

Function findcomposemail ()
'Finds the Compose mail window's handle

Dim bb As Integer
Dim dis_win As Integer

dis_win = findchildbyclass(aolhwnd(), "AOL Child")

begin_find_composemail:

bb = findchildbytitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "To:")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Subject:")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Send" & Chr(13) & "Later")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Attach")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Address" & Chr(13) & "Book")
    If bb <> 0 Then Let countt = countt + 1

If countt = 6 Then
  findcomposemail = dis_win
  Exit Function
End If
Let countt = 0
dis_win = getnextwindow(dis_win, 2)
If dis_win = Getwindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function

Function findsn ()
'Finds the user's Screen Name...they must be signed on!

Dim dis_win2 As Integer
a = findwindow("AOL Frame25", 0&)
dis_win2 = findchildbyclass(a, "AOL Child")

begin_find_SN:

bb$ = windowcaption(dis_win2)
    If Left(bb$, 9) = "Welcome, " Then Let countt = countt + 1
If countt = 1 Then
  val1 = InStr(bb$, " ")
  val2 = InStr(bb$, "!")
  Let sn$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  findsn = Trim(sn$) '_win
  Exit Function
End If
Let countt = 0
dis_win2 = getnextwindow(dis_win2, 2)
If dis_win2 = Getwindow(dis_win2, GW_HWNDLAST) Then
   findsn = 0
   Exit Function
End If

GoTo begin_find_SN

End Function

Sub pause (duratn As Integer)
'This pauses for duratn seconds
Let curent = Timer

Do Until Timer - curent >= duratn
DoEvents
Loop

End Sub

Function r_backwards (strin As String)
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

Function r_elite (strin As String)
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
If neRechr$ = "E" Then Let nextchr$ = "Ê"
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

Function r_hacker (strin As String)
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

Function r_same (strr As String)
'Returns the strin the same
Let r_same = Trim(strr)

End Function

Function r_spaced (strin As String)
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

Sub runmenu (horz, vert)
'Runs the specified AOL Menu (Horizonatl,Verticle)
'Each Position starts at 0 not 1

Dim f, gi, sm, M, a As Integer
a = findwindow("AOL Frame25", 0&)
M = GetMenu(a)
sm = GetSubMenu(M, horz)
gi = GetMenuItemID(sm, vert)
f = sendmessagebynum(a, WM_COMMAND, gi, 0)

End Sub

Sub RunMenuByString (ApplicationOfMenu, STringToSearchFor)
'This runs an application's menu by its text.  This
'includes & signs (for underlined letters)

SearchString$ = STringToSearchFor
hMenu = GetMenu(ApplicationOfMenu)
Cnt = GetMenuItemCount(hMenu)
For i = 0 To Cnt - 1
PopUphMenu = GetSubMenu(hMenu, i)
Cnt2 = GetMenuItemCount(PopUphMenu)
For O = 0 To Cnt2 - 1
    hMenuID = GetMenuItemID(PopUphMenu, O)
    MenuString$ = String$(100, " ")
    x = GetMenuString(PopUphMenu, hMenuID, MenuString$, 100, 1)
    If InStr(UCase(MenuString$), UCase(SearchString$)) Then
        SendtoID = hMenuID
        GoTo Initiate
    End If
Next O
Next i
Initiate:
x = sendmessagebynum(ApplicationOfMenu, &H111, SendtoID, 0)
End Sub

Sub Sendclick (Handle)
'Clicks something
x% = sendmessage(Handle, WM_LBUTTONDOWN, 0, 0&)
pause .05
x% = sendmessage(Handle, WM_LBUTTONUP, 0, 0&)
End Sub

Sub sendtext (handl As Integer, msgg As String)
'Sends msgg to handl
send_txt = sendmessagebystring(handl, WM_SETTEXT, 0, msgg)
End Sub

Sub showaolwins ()
'Shows all AOL Windows
fc = findchildbyclass(aolhwnd(), "AOL Child")
req = showwindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = showwindow(faa, 1)
DoEvents
Loop Until faf = faa


End Sub

Sub stayontop (frm As Form)
'Allows a window to stay on top
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Sub

Function trim_null (wstr As String)
'Trims null characters from a string
wstr = Trim(wstr)
Do Until xx = Len(wstr)
Let xx = xx + 1
Let this_chr = Asc(Mid$(wstr, xx, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(wstr, xx, 1)
Loop
trim_null = wordd

End Function

Sub waitforok ()
'Waits for the AOL OK messages that popup up
Do
DoEvents
okw = findwindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = findchildbytitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function windowcaption (hWndd As Integer)
'Gets the caption of a window
Dim WindowText As String * 255
Dim getWinText As Integer
getWinText = getwindowtext(hWndd, WindowText, 255)
windowcaption = (WindowText)
End Function

