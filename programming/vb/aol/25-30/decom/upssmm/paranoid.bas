'***************************DECLARES***********************
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal p1$, ByVal p2%) As Integer
Declare Function setparent Lib "User" (ByVal p1%, ByVal p2%) As Integer
Declare Function GetApiText Lib "xtra.dll" (ByVal p1%) As String
Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName$, ByVal wFlags%) As Integer
Declare Sub setwindowtext Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
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
Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function DeleteMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function Getwindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function getfocus Lib "User" () As Integer
Declare Function GetNextWindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetMenu% Lib "User" (ByVal hWnd%)
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuItemID% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuState Lib "User" (ByVal hMenu As Integer, ByVal wId As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetSubMenu% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function findwindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GetTopWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal Lparam&)
Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal Lparam$)
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, Lparam As Any) As Long
Declare Function AOLGetList% Lib "311.Dll" (ByVal index%, ByVal Buf$)
Declare Function getnames Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByValp2$) As Integer

Declare Function findchildbytitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function findchildbyclass% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal Cx%, ByVal Cy%, ByVal f%) As Integer
Declare Function GetParent Lib "User" (ByVal p1%) As Integer
Declare Function iswindowvisible Lib "User" (ByVal p1%) As Integer
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, lpdwProcessId As Integer) As Integer

'*************************CONSTANTS************************
Global Const SW_HIDE = 0
Global Const SW_SHOW = 5
Global Const SW_SHOWMAXIMIZED = 3
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
Global stop_busting_in
Global imon
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Type RECT
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
End Type
Dim picopen As Variant
Dim mmmsg As Variant
Declare Function FillRect Lib "User" (ByVal hDC As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As RECT)
Declare Function GetDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
Declare Sub SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Dim hBrush%
Type MODEL
  usVersion         As Integer
  fl                As Long
  pctlproc          As Long
  fsClassStyle      As Integer
  flWndStyle        As Long
  cbCtlExtra        As Integer
  idBmpPalette      As Integer
  npszDefCtlName    As Integer
  npszClassName     As Integer
  npszParentClassName As Integer
  npproplist        As Integer
  npeventlist       As Integer
  nDefProp          As String * 1
  nDefEvent         As String * 1
  nValueProp        As String * 1
  usCtlVersion      As Integer
End Type

'Subs and Functions for "APIGuide.Dll"
Declare Sub agCopyData Lib "APIGuide.Dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "APIGuide.Dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "APIGuide.Dll" (ByVal l&, lw%, lh%)
Declare Sub agOutp Lib "APIGuide.Dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "APIGuide.Dll" (ByVal portid%, ByVal outval%)
Declare Function agGetControlHwnd% Lib "APIGuide.Dll" (hctl As Control)
Declare Function agGetInstance% Lib "APIGuide.Dll" ()
Declare Function agGetAddressForObject& Lib "APIGuide.Dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (ByVal lpString$)
Declare Function agGetAddressForVBString& Lib "APIGuide.Dll" (vbstring$)
Declare Function agGetControlName$ Lib "APIGuide.Dll" (ByVal hWnd%)
Declare Function agXPixelsToTwips& Lib "APIGuide.Dll" (ByVal pixels%)
Declare Function agYPixelsToTwips& Lib "APIGuide.Dll" (ByVal pixels%)
Declare Function agXTwipsToPixels% Lib "APIGuide.Dll" (ByVal twips&)
Declare Function agYTwipsToPixels% Lib "APIGuide.Dll" (ByVal twips&)
Declare Function agDeviceCapabilities& Lib "APIGuide.Dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agDeviceMode% Lib "APIGuide.Dll" (ByVal hWnd%, ByVal hModule%, ByVal lpszDevice$, ByVal lpszOutput$)
Declare Function agExtDeviceMode% Lib "APIGuide.Dll" (ByVal hWnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)
Declare Function agInp% Lib "APIGuide.Dll" (ByVal portid%)
Declare Function agInpw% Lib "APIGuide.Dll" (ByVal portid%)
Declare Function agHugeOffset& Lib "APIGuide.Dll" (ByVal addr&, ByVal offset&)
Declare Function agVBGetVersion% Lib "APIGuide.Dll" ()
Declare Function agVBSendControlMsg& Lib "APIGuide.Dll" (ctl As Control, ByVal msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal Value&)
Declare Function dwVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal Value&)
Declare Function AGGetStringFromLPStr Lib "APIGuide.dll" (ByVal p1&) As String


'Subs and Functions for "VBMsg.Vbx"
Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal S As String) As Long
Declare Function ptGetLongAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (l As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (i As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal i As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMsg.Vbx" (ByVal l As Long) As Long
Declare Function ptGetStringFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, ByVal cbBytes As Integer) As String
Declare Function ptMakelParam Lib "VBMsg.Vbx" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMsg.Vbx" (ByVal Lparam As Long) As Integer
Declare Function ptHiWord Lib "VBMsg.Vbx" (ByVal Lparam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMsg.Vbx" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMsg.Vbx" (ByVal ushortVal As Integer) As Long
Declare Function ptMessagetoText Lib "VBMsg.Vbx" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String
Declare Function ptRecreateControlHwnd Lib "VBMsg.Vbx" (ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL) As Long
Declare Function ptGetControlName Lib "VBMsg.Vbx" (ctl As Control) As String

'Open WinDir() + "\fate.ini" For Binary As #1
'l004C$ = String(LOF(1), 0)
'Get #1, 1, l004C$
'Close #1

'Open WinDir() + "\fate.ini" For Output As #1
'l004C$ = String(LOF(1), 0)
'Send #1, l004C$
'Close #1

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


For index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_addr
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)

ADD_AOL_LB namez$, lst
Next index%
end_addr:

End Sub

Sub AddRoomWithoutMe (lst As ListBox)
aol% = findwindow("AOL Frame25", 0&)
Chatroom% = FindChatRoom()
List% = findchildbyclass(Chatroom%, "_AOL_Listbox")
If List% = 0 Then Exit Sub
thatlb = SendMessageByNum(List%, LB_GETCOUNT, 0&, 0&)
For RoomNames = 0 To thatlb - 1
Buffer$ = String$(64, 0)
BuddyName% = AOLGetList(RoomNames, Buffer$)
FinalBuddyname$ = Left$(Buffer$, BuddyName%)
MyName$ = findsn()
If MyName$ = FinalBuddyname$ Then GoTo 18
For names = 0 To lst.ListCount - 1
 If FinalBuddyname$ = lst.List(names) Then GoTo 18
Next names
lst.AddItem FinalBuddyname$
18 :
Next RoomNames
End Sub

Sub aolclick (E1 As Integer)
'Clicks an AOL button with the given handle as E1

Exit Sub


do_wn = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
Pause (.008)
u_p = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
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

aol% = findwindow("AOL Frame25", 0&)
If aol% = 0 Then
    MsgBox "Must Be Online"
    Exit Sub
End If
Call RunMenuByString(aol%, "Compose Mail")

Do: DoEvents
aol% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(aol%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = Getwindow(subjt%, 2)
mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, MESSAGE)

'AOLIcon (icone%)

Do: DoEvents
aol% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(aol%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
aolw% = findwindow("#32770", "America Online")
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

Function bsort (tl As ListBox, c As Integer)
   Dim a As Integer, b As Integer

   For a = a To c
      DoEvents
      For b = c To a Step -1
         DoEvents
         If (StrComp(tl.List(b - 1), tl.List(b), 0) = 1) Then
            tmp = tl.List(b - 1)
            tl.RemoveItem (b - 1)
            tl.AddItem tl.List(b - 1), (b - 1)
            tl.RemoveItem b
            tl.AddItem tmp, b

            tmp = mmeroptfrm.List3.List(b - 1)
            mmeroptfrm.List3.RemoveItem (b - 1)
            mmeroptfrm.List3.AddItem mmeroptfrm.List3.List(b - 1), (b - 1)
            mmeroptfrm.List3.RemoveItem b
            mmeroptfrm.List3.AddItem tmp, b
         End If
      Next b
   Next a
End Function

Function centerform (p0132 As Form)
    p0132.Top = Screen.Height / 2 - p0132.Height / 2
    p0132.Left = Screen.Width / 2 - p0132.Width / 2
End Function

Function Click (p020E As Variant) As Variant
Dim l0212 As Variant
Dim l0216 As Variant
Pause (.01)
l0212 = SendMessageByNum(p020E, WM_LBUTTONDOWN, 0, 0&)
l0216 = SendMessageByNum(p020E, WM_LBUTTONUP, 0, 0&)
DoEvents
DoEvents
End Function

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
Buffer = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
If Buffer > 1 Then
MsgBox "You have " & Buffer & " messages in your E-Mailbox."
End If
If Buffer = 1 Then
MsgBox "You have one message in your E-Mailbox."
End If
If Buffer < 1 Then
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
faa = GetNextWindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa
End Sub

Function FindChatRoom ()
'Finds the handle of the AOL Chatroom by looking for a
'Window with a ListBox (Chat ScreenNames), Edit Box,
'(Where you type chat text), and an _AOL_VIEW.  If another
'AOL window is present that also has these 3 controls, it
'may find the wrong window.  I have never seen another AOL
'window with these 3 controls at once

aol = findwindow("AOL Frame25", 0&)
If aol = 0 Then Exit Function
b = findchildbyclass(aol, "AOL Child")

start:
c = findchildbyclass(b, "_AOL_VIEW")
If c = 0 Then GoTo nextwnd
D = findchildbyclass(b, "_AOL_EDIT")
If D = 0 Then GoTo nextwnd
e = findchildbyclass(b, "_AOL_LISTBOX")
If e = 0 Then GoTo nextwnd
'We've found it
FindChatRoom = b
Exit Function

nextwnd:
b = GetNextWindow(b, 2)
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
dis_win = GetNextWindow(dis_win, 2)
If dis_win = Getwindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function

Function FindIt (p040C As String)
l040E = findwindow("AOL FRAME25", 0&)

2 :
DoEvents
If l0416 = 0 Then
l0416 = 1
l041A = l0412
Else
l041A = GetNextWindow(a, 2)
End If
l041E$ = String(255, 0)
l0420 = GetWindowText(l041A, l041E$, 255)
l0424 = LookFor0(l041E$)
If l0424 = p040C Then
FindIt = l041A
Exit Function
End If
If l0412 = 0 Then
FindIt = 0
Exit Function
End If
l0412 = l041A
GoTo 2
End Function

Function findsn () As String
Dim yoursn$
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, "Welcome")
yoursn1$ = String(30, 0)
whocares = GetWindowText(b, yoursn1$, 250)
yoursn$ = Mid(yoursn1$, 10, 10)
window2 = InStr(yoursn$, "!")
If window2 Then
yoursn$ = Mid(yoursn$, 1, window2 - 1)
End If
findsn = yoursn$
End Function

Function FixAPIString (ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function

Function GetAOlVersion () As Integer
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
l00D2 = findwindow("AOL FRAME25", 0&)
l00D6 = findchildbyclass(l00D2, "AOL TOOLBAR")
l00DA = findchildbyclass(l00D6, "_AOL_ICON")
l00DE = GetNextWindow(l00DA, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
l00DE = GetNextWindow(l00DE, 2)
If l00DE > 0 Then
GetAOlVersion = 5
End If
If l00DE = 0 Then
GetAOlVersion = 8
End If
End Function

Function GetChatRoomName () As String
On Error Resume Next
chat1% = FindChatRoom()
x = GetWindowTextLength(chat1%)
Title$ = Space(x + 1)
x = GetWindowText(chat1%, Title$, x + 1)
Title$ = FixAPIString(Title$)
GetChatRoomName = Title$
End Function

Sub GradientB (TheForm As Form)
    Dim hBrush%
    Dim FormHeight%, red%, StepInterval%, x%, retVal%, OldMode%
    Dim FillArea As RECT
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3  'Pixel
    FormHeight = TheForm.ScaleHeight
' Divide the form into 63 regions
    StepInterval = FormHeight \ 63
    red = 255
    green = 0
    blue = 0
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For x = 1 To 63
        hBrush% = CreateSolidBrush(RGB(red, green, blue))
        retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
        retVal% = DeleteObject(hBrush)
        red = red - 4
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
' Fill the remainder of the form with black
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush% = CreateSolidBrush(RGB(0, 0, 0))
    retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
    retVal% = DeleteObject(hBrush)
    TheForm.ScaleMode = OldMode
End Sub

Function hold (duratn As Integer)
'This pauses for duratn seconds
Let curent = Timer

Do Until Timer - curent >= duratn
DoEvents
Loop
End Function

Function instantmessage (to_who As String, what As String)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
towho = to_who
c = SendMessageByString(b, WM_SETTEXT, 0, towho)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
TheText = what
c = SendMessageByString(b, WM_SETTEXT, 0, TheText) 'Put msg in



D = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = GetNextWindow(D, 2) 'Next Button
f = GetNextWindow(e, 2) 'Next
g = GetNextWindow(f, 2) '
h = GetNextWindow(g, 2) '
i = GetNextWindow(h, 2) '
j = GetNextWindow(i, 2) '
k = GetNextWindow(j, 2) '
l = GetNextWindow(k, 2) '
m = GetNextWindow(l, 2) '

x = SendMessageByNum(m, WM_LBUTTONDOWN, 0, 0&) 'Click send"
DoEvents
x = SendMessageByNum(m, WM_LBUTTONUP, 0, 0&)
End Function

Function instantmessage25 (to_who As String, what As String)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
towho = to_who
c = SendMessageByString(b, WM_SETTEXT, 0, towho)
            
b = GetNextWindow(b, 2) 'Find msg area
TheText = what
c = SendMessageByString(b, WM_SETTEXT, 0, TheText) 'Put msg in

e = findchildbytitle(x, "Send")

whocares = Click(e)
End Function

Function IsAolOn ()
a = findwindow("AOL Frame25", 0&)
If a = 0 Then
   MsgBox "AOL Isn't Running!", 16
   IsAolOn = 0
   GoTo Place
End If
b = findchildbytitle(a, "Welcome")
c = String(30, 0)
D = GetWindowText(b, c, 250)
If D <= 7 Then
   MsgBox "Not Signed On!", 16
   IsAolOn = 0
   GoTo Place
End If
IsAolOn = 1
Place:
End Function

Function isaolonwithoutmsg ()
a = findwindow("AOL Frame25", 0&)
If a = 0 Then
   isaolonwithoutmsg = 0
   GoTo Place4
End If
b = findchildbytitle(a, "Welcome")
c = String(30, 0)
D = GetWindowText(b, c, 250)
If D <= 7 Then
   isaolonwithoutmsg = 0
   GoTo Place4
End If
isaolonwithoutmsg = 1
Place4:
End Function

Function keywor (where$)
   a = findwindow("AOL Frame25", 0&)
   Call RunMenuByString(a, "Keyword...")
   Do: DoEvents                          'this loads the KW screen.
      x = findchildbytitle(a, "Keyword")    'Find the KW Screen.
      If x <> 0 Then Exit Do
   Loop
   b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                 'place the Keyword in.
   c = SendMessageByString(b, WM_SETTEXT, 0, where$) 'Put our KW in.
   D = findchildbyclass(x, "_AOL_ICON")        'Find the GO Button.
   e = SendMessageByNum(D, WM_LBUTTONDOWN, 0, 0&) 'Down click
   e = SendMessageByNum(D, WM_LBUTTONUP, 0, 0&)   'Up Click
End Function

Function KTEncrypt (ByVal password, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(password) = 0 Then Error 31100
  
  'Is password too long
  If Len(password) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)

    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Function LineToChat (thestring As String) As Variant
b = FindChatRoom()
l = findchildbyclass(b, "_AOL_EDIT")
m = SendMessageByString(l, WM_SETTEXT, 0, thestring)
n = GetNextWindow(l, 2)
o = Click(n)
DoEvents
End Function

Function LookFor0 (p0BB2 As Variant) As Variant
Dim l0BB6 As Variant
l0BB6 = InStr(p0BB2, Chr(0))
If l0BB6 Then
LookFor0 = Mid(p0BB2, 1, l0BB6 - 1)
Else
LookFor0 = p0BB2
End If
End Function

Function ParaLeft () As Variant
ParaLeft = "^v•UPS•v^ "
End Function

Function ParaRight () As Variant
ParaRight = " ^v•UPS•v^"
End Function

Sub Pause (duratn As Integer)
'This pauses for duratn seconds
Let curent = Timer

Do Until Timer - curent >= duratn
DoEvents
Loop
End Sub

Function PlayWav (File)
   SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySound(SoundName$, wFlags%)
End Function

Sub qsort (tl As Control, l As Integer, r As Integer)
   Dim i As Integer
   Dim j As Integer
   i = l
   j = r

   x = tl.List((l + r) / 2)

   While (i <= j)
      While ((StrComp(tl.List(i), x, 0) < 0) And (i < r))
         DoEvents
         i = i + 1
      Wend
      While ((StrComp(tl.List(j), x, 0) > 0) And (j > l))
         DoEvents
         j = j - 1
      Wend
      If (i <= j) Then
         tmp = tl.List(i)
         tl.RemoveItem i
         tl.AddItem tl.List(j - 1), i
         tl.RemoveItem j
         tl.AddItem tmp, j
         tmp = mmeroptfrm.List3.List(i)
         mmeroptfrm.List3.RemoveItem (i)
         mmeroptfrm.List3.AddItem mmeroptfrm.List3.List(j - 1), i
         mmeroptfrm.List3.RemoveItem j
         mmeroptfrm.List3.AddItem tmp, j
         i = i + 1
         j = j - 1
      End If
      DoEvents
   Wend

   If (l < j) Then
      qsort tl, l, j
   End If
   If (i < r) Then
      qsort tl, i, r
   End If
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

Sub roombustcode ()
End Sub

Sub runmenu (horz, vert)
'Runs the specified AOL Menu (Horizonatl,Verticle)
'Each Position starts at 0 not 1

Dim f, gi, sm, m, a As Integer
a = findwindow("AOL Frame25", 0&)
m = GetMenu(a)
sm = GetSubMenu(m, horz)
gi = GetMenuItemID(sm, vert)
f = SendMessageByNum(a, WM_COMMAND, gi, 0)

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
For o = 0 To Cnt2 - 1
    hMenuID = GetMenuItemID(PopUphMenu, o)
    MenuString$ = String$(100, " ")
    x = GetMenuString(PopUphMenu, hMenuID, MenuString$, 100, 1)
    If InStr(UCase(MenuString$), UCase(SearchString$)) Then
        SendtoID = hMenuID
        GoTo Initiate
    End If
Next o
Next i
Initiate:
x = SendMessageByNum(ApplicationOfMenu, &H111, SendtoID, 0)
End Sub

Function scrollform (TheForm As Form)

realheight = TheForm.Height
realwidth = TheForm.Width
TheForm.Height = 0
TheForm.Visible = True
scalefac = 30

Do
   If TheForm.Height + scalefac > realheight Then
      TheForm.Height = realheight
      TheForm.Width = realwidth
      Exit Function
   End If
   TheForm.Height = TheForm.Height + (scalefac * (realwidth / realheight))
   DoEvents
Loop

TheForm.Height = realheight

End Function

Sub Sendclick (Handle)
'Clicks something
x% = SendMessage(Handle, WM_LBUTTONDOWN, 0, 0&)
Pause .05
x% = SendMessage(Handle, WM_LBUTTONUP, 0, 0&)
End Sub

Sub sendtext (TheText)
Chatroom% = FindChatRoom()
AolEdit% = findchildbyclass(Chatroom%, "_AOL_Edit")
x = SendMessageByString(AolEdit%, WM_SETTEXT, 0&, TheText)
D = SendMessageByNum(AolEdit%, WM_CHAR, 13, 0)
DoEvents
DoEvents
End Sub

Sub showaolwins ()
'Shows all AOL Windows
fc = findchildbyclass(aolhwnd(), "AOL Child")
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

Function SignOffQuick ()
test = IsAolOn()
If test = 0 Then Exit Function
If GetAOlVersion() = 0 Then
MsgBox "Could not detect your version of America Online", 16
Exit Function
End If
If GetAOlVersion() = 5 Then
l0126 = findwindow("AOL FRAME25", 0&)
l012A = findchildbytitle(l0126, "Welcome")
l012E$ = String(30, 0)
l0130 = GetWindowText(l012A, l012E$, 250)
If l0130 <= 7 Then
MsgBox "Not Signed On!", 16
Exit Function
End If
l0134 = findwindow("AOL FRAME25", 0&)
l013A = SendMessage(l0134, 16, 0, 0)
it:
DoEvents
whocares = findwindow("_AOL_MODAL", 0&)
If whocares = 0 Then GoTo it
whocares2 = findchildbytitle(whocares, "Cancel")
whocares3 = Click(whocares2)
12 :
DoEvents
l013E = findwindow("_AOL_MODAL", 0&)
If l013E = 0 Then GoTo 12
l0142 = findchildbytitle(l013E, "&Yes")
m002A = Click(l0142)
Do Until 2 > 3
l0148 = findwindow("#32770", "Download Manager")
If l0148 > 0 Then
l014C = findchildbytitle(l0148, "&No")
l0150 = Click(l014C)
End If
l014C = findchildbytitle(l0134, "Goodbye")
If l014C > 0 Then
Exit Function
End If
DoEvents
Loop
Else
l0126 = findwindow("AOL FRAME25", 0&)
l012A = findchildbytitle(l0126, "Welcome")
l012E$ = String(30, 0)
l0130 = GetWindowText(l012A, l012E$, 250)
If l0130 <= 7 Then
MsgBox "Not Signed On!", 16
Exit Function
End If
l0134 = findwindow("AOL FRAME25", 0&)
l013A = SendMessage(l0134, 16, 0, 0)
29 :
DoEvents
l013E = findchildbytitle(l0134, "Exit?")
If l013E = 0 Then GoTo 29
l0142 = findchildbyclass(l013E, "_AOL_icon")
m002A = GetNextWindow(l0142, 2)
l0148 = GetNextWindow(m002A, 2)
l014C = GetNextWindow(l0148, 2)
l0150 = GetNextWindow(l014C, 2)
l0156 = GetNextWindow(l0150, 2)
l015A = Click(l0156)
End If
End Function

Sub Stayontop (frm As Form)
'Allows a window to stay on top
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function Trim_Null (wstr As String)

End Function

Function trimnull (p0BB2 As Variant) As Variant
Dim l0BB6 As Variant
l0BB6 = InStr(p0BB2, Chr(0))
If l0BB6 Then
trimnull = Mid(p0BB2, 1, l0BB6 - 1)
Else
trimnull = p0BB2
End If
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
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function WaitForServerWnd (p0C0C As Variant)
15 :
DoEvents
a = findwindow("AOL FRAME25", 0&)
b = findchildbytitle(a, p0C0C)
c = findchildbyclass(b, "_AOL_TREE")
D = findwindow("#32770", "America Online")
e = findchildbytitle(D, "OK")
f = GetNextWindow(e, 1)
thestring$ = String(255, 0)
whocares = SendMessageByString(f, 13, 255, thestring$)
temp = InStr(thestring$, Chr(0))
If temp Then
valofret = Mid(thestring$, 1, temp - 1)
Else
valofret = thestring$
End If
nextstring = Mid(thestring$, 1, 11)
If nextstring = "You have no" Then
whocares = Click(e)
MsgBox "You have no " + p0C0C + "!", 16
WaitForServerWnd = 1
Exit Function
End If
If c = 0 Then GoTo 15
Do Until 2 > 3
pwait.Label1.Caption = "Waiting for " + p0C0C + " to load up..."
whocares3 = SendMessageByNum(c, 1036, 0, 0)
DoEvents
Pause 2
whocares4 = SendMessageByNum(c, 1036, 0, 0)
DoEvents
If whocares3 = whocares4 Then Exit Do
Loop
pwait.Label1.Caption = "Please Wait, Creating Mail List..."
End Function

Function waitkill ()
   a = findwindow("AOL Frame25", 0&)
   Call RunMenuByString(a, "Get A Member's Profile") 'Run the
                    'about AOL menu by its name via the function
                    'in the main bas file
   Do: DoEvents 'We ran the menu, now wait for it to appear
   Loop Until findchildbytitle(a, "Get A Member's Profile")
 
   x = SendMessageByNum(findchildbytitle(a, "Get A Member's Profile"), WM_CLOSE, 0, 0&)
        'Tell the about window to close
End Function

Function WinDir ()
l0C46$ = String(255, 0)
l0C48 = GetWindowsDirectory(l0C46$, 255)
WinDir = trimnull(l0C46$)
End Function

Function windowcaption (hWndd As Integer)
'Gets the caption of a window
Dim WindowText As String * 255
Dim getWinText As Integer
getWinText = GetWindowText(hWndd, WindowText, 255)
windowcaption = (WindowText)
End Function

