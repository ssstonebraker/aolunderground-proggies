Attribute VB_Name = "ryan_"
'sup this is sic akak ryan,
'here is my bas entitle ryan_
'ryan is my real name, hence
'the name of the bas

'ok not everything works, and thats why im asking you to help me out
'im in need of people to test out my subs and ell me which are working
'and any and i mean any problems you encounter while using it

'this is only one of the many beta versions of this bas, because well there
'is still a lot of stuff i am going to add but i want to make sure that all
'the existing subs work before adding anymore

'please, if you see, or encounter a problem please email me at: ryan_@hotmail.com
'yeas that is my own hotmail nobody has access to it...

'please test out all of the subs, and if you dont understand what a sub is
'is supposed to do be sure to ask me

'contacts:
'email1: rjx@geeklife.com
'email2: codis1@hotmail.com
'aim: v8i, visualcpp
'aol: itscarded, RYAN AKA SIC, isickening

'thanks again!
 


'all variables MUST be declared
Option Explicit

'normal, everyday api declarations
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


'constants
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETITEMDATA = &H150
Public Const CB_SETCURSEL = &H14E
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const VK_DOWN = &H28
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SPACE = &H20
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

'various types
Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Sub aol_caption(newcap As String)
Dim aol As Long
 aol& = FindWindow("AOL Frame25", vbNullString)
 Call SendMessageByString(aol&, WM_SETTEXT, 0&, newcap$)
End Sub

Public Function mail_findbox() As Long
Dim aol As Long, mdi As Long, child As Long, thecap As String
 aol& = FindWindow("AOL Frame25", vbNullString)
 mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
 child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
 thecap$ = win_getcaption(child&)
 If InStr(thecap$, "'s Online Mailbox") Then
  aol_findmailbox& = child&
  Exit Function
 Else
  Do:
   DoEvents
    child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
    thecap$ = win_getcaption(child&)
    If InStr(thecap$, "'s Online Mailbox") Then
     aol_findmailbox& = child&
     Exit Function
    End If
   Loop Until child& = 0
 End If
aol_findmailbox& = child&
End Function

Public Sub form_move(TheForm As Form)
 Call ReleaseCapture
 Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Sub list_save(nameANDdir As String, thelist As listbox)
Dim saving As Long
On Error Resume Next
Open nameANDdir$ For Output As #1
 For saving = 0 To thelist.ListCount - 1
  Print #1, thelist.List(saving)
 Next saving
Close #1
End Sub

Public Sub combo_save(nameANDdir As String, thecombo As ComboBox)
Dim saving As Long
On Error Resume Next
Open nameANDdir$ For Output As #1
 For saving = 0 To thecombo.ListCount - 1
  Print #1, thecombo.List(saving)
 Next saving
Close #1
End Sub

Public Sub list_load(nameANDdir As String, thelist As listbox)
Dim filetxt As String
On Error Resume Next
Open nameANDdir$ For Input As #1
 While Not EOF(1)
  Input #1, filetxt$
  thelist.AddItem filetxt$
 Wend
Close #1
End Sub
Public Sub combo_load(nameANDdir As String, thecombo As ComboBox)
Dim filetxt As String
On Error Resume Next
Open nameANDdir$ For Input As #1
 While Not EOF(1)
  Input #1, filetxt$
  thecombo.AddItem filetxt$
 Wend
Close #1
End Sub

Public Function mail_liststring(thelist As listbox) As String
Dim x As Integer
 For x = 0 To thelist.ListCount - 1
  If x = 0 Then mail_liststring$ = mail_liststring$ + thelist.List(x) Else:
  If x = thelist.ListCount - 1 Then mail_liststring$ = mail_liststring$ + thelist.List(x) Else:
  mail_liststring$ = mail_liststring$ + ", " + thelist.List(x)
 Next x
End Function
Public Function mail_combostring(thecombo As ComboBox) As String
Dim x As Integer
 For x = 0 To thecombo.ListCount - 1
  If x = 0 Then mail_combostring$ = mail_combostring$ + thecombo.List(x) Else:
  If x = thecombo.ListCount - 1 Then mail_combostring$ = mail_combostring$ + thecombo.List(x) Else:
  mail_combostring$ = mail_combostring$ + ", " + thecombo.List(x)
 Next x
End Function

Public Function pc_fulldate(whichone As Integer) As String
Dim tempdate As String
 If whichone = 1 Then: tempdate$ = Format(Now, "m/d/y")
 If whichone = 2 Then: tempdate$ = Format(Now, "dddd, mmmm dd, yyyy")
 If whichone = 3 Then: tempdate$ = Format(Now, "d-mmm")
 If whichone = 4 Then: tempdate$ = Format(Now, "mmmm-yy")
 If whichone = 5 Then: tempdate$ = Format(Now, "d-mmmm")
pc_fulldate$ = tempdate$
End Function



Function file_isexisting(ByVal thefile As String) As Boolean
Dim leng As Integer
On Error Resume Next
leng = Len(Dir$(thefile$))
 If Err Or leng = 0 Then file_isexisting = False Else:
  file_isexisting = True
End Function

Public Sub txt_paste(Where As TextBox)
 Where.SelText = Clipboard.GetText
End Sub

Public Sub txt_copy(thetext As TextBox)
 Clipboard.Clear
 Clipboard.SetText thetext.Text
End Sub

Public Function misc_depart(snpw As String, thechar As String) As String
Dim full As String, SN As String, pw As String
If Len(thechar$) > 1 Then Exit Function Else:
 full$ = snpw$
 full$ = Right(full$, Len(full$) + 1)
 SN$ = Left(full$, InStr(full$, thechar$) - 1)
 pw$ = Right(full$, Len(full$) - Len(SN$) - 1)
  'MsgBox "sn: " & sn$
  'MsgBox "pw: " & pw$
End Function

Public Function pc_lastboot() As String
 Dim hrs As Long, mins As Long, fullcount As Long
 fullcount& = GetTickCount
 hrs = ((fullcount& / 1000) / 60) / 60
 mins = ((fullcount& / 1000) / 60) Mod 60
 pc_lastboot$ = hrs& & ":" & mins&
End Function

Sub misc_timeout(Duration As Integer)
Dim begin
 begin = Timer
 Do While Timer - begin > Duration
 Loop
End Sub

Public Function file_stringscan(Path As String, query As String) As Boolean
Dim hah As String
On Error Resume Next
Open Path$ For Binary Access Read Write As #1
 hah$ = String(LOF(1), " ")
 Get #1, 1, hah$
 If InStr(LCase(hah$), LCase(query$)) Then file_stringscan = True
Close #1
End Function

Public Sub form_stayontop(TheForm As Form)
 Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub form_stayofftop(TheForm As Form)
 Call SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Function cd_TrackNum(Track As Integer)
 mciSendString "seek cd to " & Str(Track), 0, 0, 0
End Function

Function cd_musicornot()
 Dim hah As String * 30, a As Long, ismusic As Boolean
 mciSendString "status cd media present", a, Len(a), 0
 cd_musicornot = hah
End Function

Public Function cd_trackcount()
 Dim hah As String * 30, a As Long
 mciSendString "status cd number of tracks wait", a, Len(a), 0
 cd_trackcount = CInt(Mid$(hah, 1, 2))
End Function

Public Sub cd_trayopen()
 mciSendString "set cd door open", 0, 0, 0
End Sub

Public Sub CD_Pause()
 mciSendString "pause cd", 0, 0, 0
End Sub

Public Function CD_Stop()
 mciSendString "stop cd wait", 0, 0, 0
End Function

Public Sub CD_Play()
 mciSendString "play cd", 0, 0, 0
End Sub

Public Sub cd_trayclose()
 mciSendString "set cd door closed", 0, 0, 0
End Sub

Public Function txt_hacker(thetext As String) As String
Dim thetext2 As String, newchar As String
Dim leng As Integer, x As Integer, numspaces As Integer
leng = Len(thetext$)
Do While numspaces < leng
numspaces = numspaces + 1
newchar$ = Mid(thetext$, numspaces, 1)
 If newchar$ = "A" Then: newchar$ = "a"
 If newchar$ = "E" Then: newchar$ = "e"
 If newchar$ = "I" Then: newchar$ = "i"
 If newchar$ = "O" Then: newchar$ = "o"
 If newchar$ = "U" Then: newchar$ = "u"
 If newchar$ = "Y" Then: newchar$ = "y"
 If newchar$ = "b" Then: newchar$ = "B"
 If newchar$ = "c" Then: newchar$ = "C"
 If newchar$ = "d" Then: newchar$ = "D"
 If newchar$ = "f" Then: newchar$ = "F"
 If newchar$ = "g" Then: newchar$ = "G"
 If newchar$ = "h" Then: newchar$ = "H"
 If newchar$ = "j" Then: newchar$ = "J"
 If newchar$ = "k" Then: newchar$ = "K"
 If newchar$ = "l" Then: newchar$ = "L"
 If newchar$ = "m" Then: newchar$ = "M"
 If newchar$ = "n" Then: newchar$ = "N"
 If newchar$ = "p" Then: newchar$ = "P"
 If newchar$ = "q" Then: newchar$ = "Q"
 If newchar$ = "r" Then: newchar$ = "R"
 If newchar$ = "s" Then: newchar$ = "S"
 If newchar$ = "t" Then: newchar$ = "T"
 If newchar$ = "v" Then: newchar$ = "V"
 If newchar$ = "w" Then: newchar$ = "W"
 If newchar$ = "x" Then: newchar$ = "X"
 If newchar$ = "z" Then: newchar$ = "z"
txt_hacker$ = txthacker$ + newchar$
Loop
End Function

Public Function txt_bold1(thetext As String) As String
Dim thetext2 As String, leng As Integer, txt As String, taken As Integer
Dim numspaces As Integer, newchar As String, firstchar As String
txt$ = thetext$
leng = Len(txt$)
firstchar$ = Left(txt$, 1)
st:
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  If newchar$ = firstchar$ And taken <> 1 Then GoTo ack Else:
  If newchar$ = " " Then GoTo ox Else GoTo hoo
ox:
   thetext2$ = thetext2$ + newchar$
   numspaces = numspaces + 1
   newchar$ = Mid(txt$, numspaces, 1)
   thetext2$ = thetext2$ + "<b>" + newchar$ + "</b>"
   GoTo st
ack:
   thetext2$ = thetext2$ + "<b>" + newchar$ + "</b>"
   taken = 1
   GoTo st
hoo:
   thetext2$ = thetext2$ + newchar$
   GoTo st
 Loop
txt_bold1$ = thetext2$
End Function

Function app_runcount(thefile As String) As Integer
Dim num
On Error Resume Next
num = txt_autoload("" & thefile$)
 If string_isnumeric("" & num) = False Then num = 0
  num = num + 1
  Call txt_autosave("" & thefile$ & "", "" & num)
  app_runcount = num
 
End Function

Function txt_autoload(thefile As String)
Dim hah, heh, txt As String
'f = hah, heh = cha, txt = textda
On Error Resume Next
hah = FreeFile
txt$ = ""
 If file_isexisting("" & thefile$) = True Then
  If Len(thefile$) Then
   Open thefile$ For Input As #hah
    Do While Not EOF(hah)
     heh = Input(1, #hah)
      If heh <> Chr(10) Then
       txt$ = "" & txt$ & heh & ""
      End If
    Loop
   Close #hah
  End If
hek:
If Right(txt$, 1) = Chr(13) Then
 txt$ = Left(txt$, Len(txt$) - 1)
 GoTo hek
End If
If Right(txt$, 1) = Chr(10) Then
 txt$ = Left(txt$, Len(txt$) - 1)
 GoTo hek
End If
 txt_autoload = txt$
Else
 txt_autoload = ""
End If
End Function


Sub txt_autosave(thefile As String, thetext As String)
Dim hah
On Error Resume Next
If Not file_isexisting("" & thefile$ & "") Then
 hah = FreeFile
 Open thefile$ For Output Access Write As #5 Len = 4096
 Print #5, thetext$
 Close #5
 Exit Sub
End If
hah = FreeFile
Open thefile$ For Output As #5
Print #5, thetext$
Close #5
End Sub

Public Function txt_strike1(thetext As String) As String
Dim thetext2 As String, leng As Integer, txt As String, taken As Integer
Dim numspaces As Integer, newchar As String, firstchar As String
txt$ = thetext$
leng = Len(txt$)
firstchar$ = Left(txt$, 1)
st:
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  If newchar$ = firstchar$ And taken <> 1 Then GoTo ack Else:
  If newchar$ = " " Then GoTo ox Else GoTo hoo
ox:
   thetext2$ = thetext2$ + newchar$
   numspaces = numspaces + 1
   newchar$ = Mid(txt$, numspaces, 1)
   thetext2$ = thetext2$ + "<s>" + newchar$ + "</s>"
   GoTo st
ack:
   thetext2$ = thetext2$ + "<s>" + newchar$ + "</s>"
   taken = 1
   GoTo st
hoo:
   thetext2$ = thetext2$ + newchar$
   GoTo st
 Loop
txt_bold1$ = thetext2$
End Function


Public Function txt_uline1(thetext As String) As String
Dim thetext2 As String, leng As Integer, txt As String
Dim numspaces As Integer, newchar As String, firstchar As String
txt$ = thetext$
leng = Len(txt$)
firstchar$ = Left(txt$, 1)
st:
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  If newchar$ = firstchar$ And taken <> 1 Then GoTo ack Else:
  If newchar$ = " " Then GoTo ox Else GoTo hoo
ox:
   thetext2$ = thetext2$ + newchar$
   numspaces = numspaces + 1
   newchar$ = Mid(txt$, numspaces, 1)
   thetext2$ = thetext2$ + "<u>" + newchar$ + "</u>"
   GoTo st
ack:
   thetext2$ = thetext2$ + "<u>" + newchar$ + "</u>"
   taken = 1
   GoTo st
hoo:
   thetext2$ = thetext2$ + newchar$
   GoTo st
 Loop
txt_uline1$ = thetext2$
End Function

Public Function txt_italic1(thetext As String) As String
Dim thetext2 As String, leng As Integer, txt As String, taken As Integer
Dim numspaces As Integer, newchar As String, firstchar As String
txt$ = thetext$
leng = Len(txt$)
firstchar$ = Left(txt$, 1)
st:
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  If newchar$ = firstchar$ And taken <> 1 Then GoTo ack Else:
  If newchar$ = " " Then GoTo ox Else GoTo hoo
ox:
   thetext2$ = thetext2$ + newchar$
   numspaces = numspaces + 1
   newchar$ = Mid(txt$, numspaces, 1)
   thetext2$ = thetext2$ + "<i>" + newchar$ + "</i>"
   GoTo st
ack:
   thetext2$ = thetext2$ + "<i>" + newchar$ + "</i>"
   taken = 1
   GoTo st
hoo:
   thetext2$ = thetext2$ + newchar$
   GoTo st
 Loop
txt_italic1$ = thetext2$
End Function

Public Function txt_spaced(thetext As String) As String
Dim thetext2 As String, leng As Integer
Dim numspaces As Integer, newchar As String
leng = Len(thetext$)
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(thetext$, numspaces, 1)
  newchar$ = newchar$ + " "
  thetext2$ = thetext2$ + newchar$
 Loop
txt_spaced$ = thetext2$
End Function

Public Function txt_reverse(thetext As String) As String
Dim thetext2 As String, leng As Integer, txt As String
Dim numspaces As Integer, newchar As String
txt$ = thetext$
leng = Len(txt$)
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  thetext2$ = newchar$ + thetext2$
 Loop
txt_reverse$ = thetext2$
End Function

Public Function string_isnumeric(hah) As Boolean
Dim boink As Integer, blah As Integer, hek As Integer
Dim thechar As String, hic As Integer
boink = 1
For hek = 1 To Len(blah)
thechar$ = LCase(Right(Left(blah, hek), 1))
 hic = 0
  If thechar$ = "0" Then hic = 1
  If thechar$ = "1" Then hic = 1
  If thechar$ = "2" Then hic = 1
  If thechar$ = "3" Then hic = 1
  If thechar$ = "4" Then hic = 1
  If thechar$ = "5" Then hic = 1
  If thechar$ = "6" Then hic = 1
  If thechar$ = "7" Then hic = 1
  If thechar$ = "8" Then hic = 1
  If thechar$ = "9" Then hic = 1
  If hic = 0 Then boink = 0
 Next hek
  If boink = 0 Then string_isnumeric = True Else
   string_isnumeric = True
End Function

Function txt_nospaces(thetext As String) As String
Dim nospace As Integer, thechar As String, thechars As String
If InStr(thetext$, " ") = 0 Then
 nospaces$ = thetext$
 Exit Function
End If
 For nospace = 1 To Len(thetext$)
  thechar$ = Mid(thetext$, nospace, 1)
  thechars$ = thechars$ & thechar$
   If thechars$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
   End If
 Next nospace
txt_nospaces$ = thechars$
End Function

Public Function txt_dotted(thetext As String) As String
Dim thetext2 As String, leng As Integer, txt As String
Dim numspaces As Integer, newchar As String
txt$ = thetext$
leng = Len(txt$)
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  newchar$ = newchar$ + "."
  thetext2$ = thetext2$ + newchar$
 Loop
txt_dotted$ = thetext2$
End Function

Public Function txt_varied(thetext As String, thechar As String) As String
Dim thetext2 As String, leng As Integer, txt As String
Dim numspaces As Integer, newchar As String
txt$ = thetext$
leng = Len(txt$)
 Do While numspaces < leng
  numspaces = numspaces + 1
  newchar$ = Mid(txt$, numspaces, 1)
  newchar$ = newchar$ + thechar$
  thetext2$ = thetext2$ + newchar$
 Loop
txt_varied$ = thetext2$
End Function

Public Sub pc_openwww(theurl As String)
 Dim hwnd As Long
 Call ShellExecute(hwnd, "Open", theurl$, "", App.Path, 1)
End Sub

Public Function tnet_findtnet() As Boolean
Dim tnet As Long
tnet& = FindWindow("TForm1", vbNullString)
If tnet& <> 0 Then tnet_findtnet = True Else tnet_findtnet = False
End Function

Public Sub aol_keyword(Keyword As String)
Dim aol As Long, Toolbar As Long, toolbar2 As Long
Dim aolcombo As Long, keytext As Long, usertext As String
Dim TextLen As Long
aol& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
aolcombo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
keytext& = FindWindowEx(aolcombo&, 0&, "Edit", vbNullString)
 TextLen& = SendMessage(keytext&, WM_GETTEXTLENGTH, 0&, 0&)
 usertext$ = String(TextLen&, 0&)
Call SendMessageByString(keytext&, WM_GETTEXT, TextLen& + 1&, usertext$)
Call SendMessageByString(keytext&, WM_SETTEXT, 0&, "")
Call SendMessageByString(keytext&, WM_SETTEXT, 0&, Keyword$)
Call SendMessageLong(keytext&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(keytext&, WM_CHAR, VK_RETURN, 0&)
 DoEvents
Call SendMessageByString(keytext&, WM_SETTEXT, 0&, usertext$)
End Sub

Public Function chat_findchat() As Long 'find the aol chatroom window
 Dim aol As Long, mdi As Long, child As Long
 Dim AOLList As Long, richcntl As Long
 Dim aolicon As Long, AOLStatic As Long
 aol& = FindWindow("AOL Frame25", vbNullString)
 mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
 child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
 richcntl& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
 AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
 aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
 AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
 If child& <> 0 And richcntl& <> 0 And AOLList& <> 0 And aolicon& <> 0 And AOLStatic& <> 0 Then
  chat_findchat& = child&
  Exit Function
Else
 Do:
  DoEvents
   child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
   richcntl& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
   AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
   aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
   AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If child& <> 0 And richcntl& <> 0 And AOLList& <> 0 And aolicon& <> 0 And AOLStatic& <> 0 Then
     chat_findchat& = child&
     Exit Function
    End If
   Loop Until child& = 0
 End If
chat_findchat& = child&
End Function

Public Sub chat_sendtext(Text As String) 'sends text to chat
 Dim chatroom As Long, richcntl As Long, RICHCNTL2 As Long
 Dim usertext As String, TextLen As Long
 chatroom& = theroom&
 richcntl& = FindWindowEx(chatroom&, 0&, "RICHCNTL", vbNullString)
 RICHCNTL2& = FindWindowEx(chatroom&, richcntl&, "RICHCNTL", vbNullString)
  TextLen& = SendMessage(RICHCNTL2&, WM_GETTEXTLENGTH, 0&, 0&)
  usertext$ = String(TextLen&, 0&)
 Call SendMessageByString(RICHCNTL2&, WM_GETTEXT, TextLen& + 1&, usertext$)
 Call SendMessageByString(RICHCNTL2&, WM_SETTEXT, 0&, "")
 Call SendMessageByString(RICHCNTL2&, WM_SETTEXT, 0&, Text$)
 Call SendMessageLong(RICHCNTL2&, WM_CHAR, ENTER_KEY, 0&)
 Call SendMessageByString(RICHCNTL2&, WM_SETTEXT, 0&, usertext$)
End Sub

Public Sub people_peopleconn()
Call aol_runmenubynum(10&, 1&, False)
End Sub


Public Sub people_chatnow()
Call aol_runmenubynum(10&, 2&, False)
End Sub

Public Sub people_findachat()
Call aol_runmenubynum(10&, 3&, False)
End Sub

Public Sub people_startchat()
Call aol_runmenubynum(10&, 4&, False)
End Sub

Public Sub people_aollive()
Call aol_runmenubynum(10&, 5&, False)
End Sub

Public Sub people_openim()
Call aol_runmenubynum(10&, 6&, False)
End Sub

Public Sub people_viewbuddy()
Call aol_runmenubynum(10&, 7&, False)
End Sub

Public Sub people_msgtopager()
Call aol_runmenubynum(10&, 8&, False)
End Sub

Public Sub people_openmemdir()
Call aol_runmenubynum(10&, 9&, False)
End Sub

Public Sub people_openlocate()
Call aol_runmenubynum(10&, 10&, False)
End Sub

Public Sub people_opengetpro()
Call aol_runmenubynum(10&, 11&, False)
End Sub

Public Sub people_whitepages()
Call aol_runmenubynum(10&, 12&, False)
End Sub

Public Sub channel_aoltoday()
Call aol_runmenubynum(9&, 1&, False)
End Sub

Public Sub channel_news()
Call aol_runmenubynum(9&, 2&, False)
End Sub

Public Sub channel_sports()
Call aol_runmenubynum(10&, 3&, False)
End Sub

Public Sub channel_influence()
Call aol_runmenubynum(9&, 4&, False)
End Sub

Public Sub channel_travel()
Call aol_runmenubynum(9&, 5&, False)
End Sub

Public Sub channel_international()
Call aol_runmenubynum(9&, 6&, False)
End Sub

Public Sub channel_personalfinace()
Call aol_runmenubynum(9&, 7&, False)
End Sub

Public Sub channel_workplace()
Call aol_runmenubynum(9&, 8&, False)
End Sub

Public Sub channel_computing()
Call aol_runmenubynum(9&, 10&, False)
End Sub

Public Sub channel_research()
Call aol_runmenubynum(9&, 11&, False)
End Sub

Public Sub channel_entertainment()
Call aol_runmenubynum(9&, 12&, False)
End Sub

Public Sub channel_gaming()
Call aol_runmenubynum(9&, 13&, False)
End Sub

Public Sub channel_interests()
Call aol_runmenubynum(9&, 14&, False)
End Sub

Public Sub channel_lifestyles()
Call aol_runmenubynum(10&, 15&, False)
End Sub

Public Sub channel_shopping()
Call aol_runmenubynum(9&, 16&, False)
End Sub

Public Sub channel_health()
Call aol_runmenubynum(9&, 17&, False)
End Sub

Public Sub channel_families()
Call aol_runmenubynum(10&, 18&, False)
End Sub

Public Sub channel_kidsonly()
Call aol_runmenubynum(10&, 19&, False)
End Sub

Public Sub channel_local()
Call aol_runmenubynum(10&, 20&, False)
End Sub

Public Function tnet_sendtext(Text As String)
Dim tnet As Long, tnetedit As Long
tnet& = FindWindow("TForm1", vbNullString)
tnetedit& = FindWindowEx(tnet&, 0&, "TPanel", vbNullString)
 Call SendMessageByString(tnetedit&, WM_SETTEXT, 0&, Text$)
 Call SendMessageLong(tnetedit&, WM_CHAR, ENTER_KEY, 0&)
End Function

Public Sub chat_clear()
Dim Chat As Long, richcntl As Long
 Chat& = theroom()
 richcntl& = FindWindowEx(Chat&, 0&, "RICHCNTL", vbNullString)
 Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, "")
End Sub

Public Function chat_addtolist(thelist As listbox, usertoo As Boolean) 'listbox
 Dim cpro As Long, itmh As Long, SN As String, psnh As Long
 Dim rb As Long, index As Long, Room As Long, rl As Long
 Dim st As Long, mt As Long
 Room& = theroom&
  If Room& = 0 Then Exit Function
   rl& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
   st& = GetWindowThreadProcessId(rl, cpro&)
   mt& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cpro&)
    If mt& Then
     For index& = 0 To SendMessage(rl, LB_GETCOUNT, 0, 0) - 1
      SN$ = String$(4, vbNullChar)
      itmh& = SendMessage(rl, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
      itmh& = itmh& + 24
       Call ReadProcessMemory(mt&, itmh&, SN$, 4, rb)
       Call CopyMemory(psnh&, ByVal SN$, 4)
      psnh& = psnh& + 6
      SN$ = String$(16, vbNullChar)
       Call ReadProcessMemory(mt&, psnh&, SN$, Len(SN$), rb&)
      SN$ = Left$(SN$, InStr(SN$, vbNullChar) - 1)
    If SN$ <> GetUser$ Or usertoo = True Then
     thelist.AddItem LCase(SN$) 'add nospaces!!!!!
    End If
   Next index&
Call CloseHandle(mt)
End If
End Function

Public Function chat_addtocombo(thecombo As ComboBox, usertoo As Boolean) 'listbox
 Dim cpro As Long, itmh As Long, SN As String, psnh As Long
 Dim rb As Long, index As Long, Room As Long, rl As Long
 Dim st As Long, mt As Long
 Room& = theroom&
  If Room& = 0 Then Exit Function
   rl& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
   st& = GetWindowThreadProcessId(rl, cpro&)
   mt& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cpro&)
    If mt& Then
     For index& = 0 To SendMessage(rl, LB_GETCOUNT, 0, 0) - 1
      SN$ = String$(4, vbNullChar)
      itmh& = SendMessage(rl, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
      itmh& = itmh& + 24
       Call ReadProcessMemory(mt&, itmh&, SN$, 4, rb)
       Call CopyMemory(psnh&, ByVal SN$, 4)
      psnh& = psnh& + 6
      SN$ = String$(16, vbNullChar)
       Call ReadProcessMemory(mt&, psnh&, SN$, Len(SN$), rb&)
      SN$ = Left$(SN$, InStr(SN$, vbNullChar) - 1)
    If SN$ <> GetUser$ Or usertoo = True Then
     thecombo.AddItem LCase(SN$) 'add nospaces!!!!!
    End If
   Next index&
Call CloseHandle(mt)
End If
End Function

Public Sub list_addascii(thelist As listbox)
Dim i As Integer
 For i = 33 To 255
  thelist.AddItem Chr(i)
 Next i
End Sub

Public Sub combo_addascii(thecombo As ComboBox)
Dim i As Integer
 For i = 33 To 255
  thecombo.AddItem Chr(i)
 Next i
End Sub

Public Function im_findim() As Long
Dim aol As Long, mdi As Long, child As Long, cap As String
 aol& = FindWindow("AOL Frame25", vbNullString)
 mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
 child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
 cap$ = win_getcaption(child&)
 If InStr(cap$, "Instant Message") = 1 Or InStr(cap$, "Instant Message") = 2 Or InStr(cap, "Instant Message") = 3 Then
  im_findim& = child&
  Exit Function
 Else
  Do
   child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
   cap$ = win_getcaption(child&)
    If InStr(cap$, "Instant Message") = 1 Or InStr(cap$, "Instant Message") = 2 Or InStr(cap$, "Instant Message") = 3 Then
     im_findim& = child&
     Exit Function
    End If
  Loop Until child& = 0&
 End If
 im_findim& = child&
End Function

Public Function win_getcaption(thewindow As Long) As String
Dim hah As String, leng As Long
 leng& = GetWindowTextLength(thewindow&)
 hah$ = String(leng&, 0&)
 Call GetWindowText(thewindow&, hah$, leng& + 1)
 win_getcaption$ = hah$
End Function

Function win_getcaption2(thehwnd As Long) As String
Dim hah As String, leng As Long
 leng& = GetWindowTextLength(thehwnd&)
 hah$ = String(leng&, 0&)
 Call GetWindowText(thehwnd, hah$, (leng& + 1))
 win_getcaption2$ = hah$
End Function

Public Function list_getlisttext(thewindow As Long) As String
Dim hah As String, leng As Long
 leng& = SendMessage(thewindow&, LB_GETTEXTLEN, 0&, 0&)
 hah$ = String(leng&, 0&)
 Call SendMessageByString(thewindow&, LB_GETTEXT, leng& + 1, hah$)
 list_getlisttext$ = hah$
End Function

Public Function win_gettext(thehandle As Long) As String
Dim hah As String, leng As Long
 leng& = SendMessage(thehandle&, WM_GETTEXTLENGTH, 0&, 0&)
 hah$ = String(leng&, 0&)
 Call SendMessageByString(thehandle&, WM_GETTEXT, leng& + 1, hah$)
 win_gettext$ = hah$
End Function

Public Function aol_getusersn() As String
Dim aol As Long, mdi As Long, welcome As Long
Dim child As Long, thesn As String
 aol& = FindWindow("AOL Frame25", vbNullString)
 mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
 child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
 thesn$ = GetCaption(child&)
  If InStr(thesn$, "Welcome, ") = 1 Then
   thesn$ = Mid$(thesn$, 10, (InStr(thesn$, "!") - 10))
   aol_getusersn$ = thesn$
   Exit Function
  Else
   Do
    child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
    thesn$ = GetCaption(child&)
     If InStr(thesn$, "Welcome, ") = 1 Then
      thesn$ = Mid$(thesn$, 10, (InStr(thesn$, "!") - 10))
      aol_getusersn$ = thesn$
      Exit Function
     End If
   Loop Until child& = 0&
  End If
 aol_getusersn$ = "n/a"
End Function

Public Sub aol_runmenu(thetopmenu As Long, thesubmenu As Long)
Dim aol As Long, mnu As Long, submnu As Long, mnuid As Long, mnuval As Long
 aol& = FindWindow("AOL Frame25", vbNullString)
 amnu& = GetMenu(aol&)
 smnu& = GetSubMenu(aMenu&, thetopmenu&)
 mnuid& = GetMenuItemID(smnu&, thesubmenu&)
 Call SendMessageLong(aol&, WM_COMMAND, mnuid&, 0&)
End Sub

Public Sub aol_runmenubystring(thestring As String)
Dim aol As Long, amnu As Long, mcnt As Long
Dim lf As Long, smnu As Long, scnt As Long
Dim ls As Long, sID As Long, thestring2 As String
 aol& = FindWindow("AOL Frame25", vbNullString)
 amnu& = GetMenu(aol&)
 mcnt& = GetMenuItemCount(amnu&)
  For lf& = 0& To mnt& - 1
   smnu& = GetSubMenu(anu&, lf&)
   scnt& = GetMenuItemCount(smnu&)
  For ls& = 0 To scnt& - 1
   sID& = GetMenuItemID(smnu&, ls&)
   thestring2$ = String$(100, " ")
   Call GetMenuString(smnu&, sID&, thestring2$, 100&, 1&)
    If InStr(LCase(thestring2$), LCase(thestring$)) Then
     Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
     Exit Sub
    End If
  Next ls&
  Next lf&
End Sub

Public Sub aol_runmenubynum(a As Long, b As Long, smnu As Boolean, Optional c As Long)
Dim aol As Long, leng As Long, aolicon As Long, Toolbar As Long
Dim toolbar2 As Long, pmnu As Long, pmnuvis As Long, i As Long
Dim aoltxt As String, CurPos As POINTAPI
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
aol& = FindWindow("AOL Frame25", vbNullString)
leng& = SendMessage(aol&, WM_GETTEXTLENGTH, 0&, 0&)
aoltxt$ = String(leng&, 0&)
Call SendMessageByString(aol&, WM_GETTEXT, leng& + 1&, aoltxt$)
Toolbar& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
 If a& = 1& Then
  aolicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
 Else
  aolicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
  For i& = 1& To a - 1&
   aolicon& = FindWindowEx(toolbar2&, aolicon&, "_AOL_Icon", vbNullString)
  Next i&
 End If
Call PostMessage(aolicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon&, WM_LBUTTONUP, 0&, 0&)
 Do: DoEvents
  pmnu& = FindWindow("#32768", vbNullString)
  pmnuvis& = IsWindowVisible(pmnu&)
 Loop Until pmnuvis& = 1&
For i& = 1& To b&
 Call PostMessage(pmnu&, WM_KEYDOWN, VK_DOWN, 0&)
 Call PostMessage(pmnu&, WM_KEYUP, VK_DOWN, 0&)
Next i&
 If smnu = True Then
  Call PostMessage(pmnu&, WM_KEYDOWN, VK_RIGHT, 0&)
  Call PostMessage(pmnu&, WM_KEYUP, VK_RIGHT, 0&)
  For i& = 1& To c& - 1&
   Call PostMessage(pmnu&, WM_KEYDOWN, VK_DOWN, 0&)
   Call PostMessage(pmnu&, WM_KEYUP, VK_DOWN, 0&)
  Next i&
End If
Call PostMessage(pmnu&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(pmnu&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.x, CurPos.Y)
End Sub

Public Sub mail_readnew()
Call aol_runmenubynum(1&, 0&, False)
End Sub

Public Sub mail_mcenter()
Call aol_runmenubynum(3&, 1&, False)
End Sub

Public Sub mail_readnew2()
Call aol_runmenubynum(3&, 2&, False)
End Sub

Public Sub mail_writenew()
Call aol_runmenubynum(3&, 3&, False)
End Sub

Public Sub mail_readold()
Call aol_runmenubynum(3&, 4&, False)
End Sub

Public Sub mail_readsent()
Call aol_runmenubynum(3&, 5&, False)
End Sub

Public Sub mail_addybook()
Call aol_runmenubynum(3&, 6&, False)
End Sub

Public Sub mail_viewprefs()
Call aol_runmenubynum(3&, 7&, False)
End Sub

Public Sub mail_controls()
Call aol_runmenubynum(3&, 8&, False)
End Sub

Public Sub mail_mextras()
Call aol_runmenubynum(3&, 9&, False)
End Sub

Public Sub mail_buildflash()
Call aol_runmenubynum(3&, 10&, False)
End Sub

Public Sub mail_runflash()
Call aol_runmenubynum(3&, 11&, False)
End Sub

Public Sub mail_readsaved()
Call aol_runmenubynum(3&, 13&, True)
End Sub

Public Sub mail_readwaiting()
Call aol_runmenubynum(3&, 14&, True)
End Sub

Public Sub mail_readcopies()
Call aol_runmenubynum(3&, 15&, True)
End Sub

Public Function chat_gettext() As String
Dim Room As Long, richcntl As Long
 If chat_findchat() = 0 Then
  chat_gettext$ = "chat room not found..."
 Else
  Room& = chat_findchat()
  richcntl& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
  chat_gettext$ = win_gettext(richcntl&)
 End If
End Function

Public Function im_gettext() As String
Dim theim As Long, richcntl As Long
 If aol_findroom() = 0 Then
  im_gettext$ = "chat room not found..."
 Else
  theim& = aol_findim()
  richcntl& = FindWindowEx(theim&, 0&, "RICHCNTL", vbNullString)
  im_gettext$ = win_gettext(richcntl&)
 End If
End Function

Public Sub win_closewin(thewin As Long)
 Call SendMessageByNum(thewin&, WM_CLOSE, 0, 0)
End Sub

Function misc_percent(comp As Integer, tot As Variant) As Integer
Dim thepercent As Long
 On Error Resume Next
 thepercent& = Int(comp / tot * 100)
End Function

Public Sub form_centertop(TheForm As Form)
 TheForm.Left = (Screen.Width - TheForm.Width) / 2
 TheForm.Top = (Screen.Height - TheForm.Height) / (Screen.Height)
End Sub

Public Function chat_getcount() As Long
Dim theroom As Long, AOLList As Long, thecount As Long
 theroom& = aol_findroom()
 AOLList& = FindWindowEx(theroom&, 0&, "_AOL_Listbox", vbNullString)
 thecount& = SendMessage(AOLList&, LB_GETCOUNT, 0&, 0&)
 chat_getcount& = thecount&
End Function

Public Sub win_hide(thewin As Long)
 Call ShowWindow(thewin&, SW_HIDE)
End Sub

Public Function Mail_CountNew() As Long
Dim Box As Long, AOLTabPage As Long, AOLTabControl As Long, AOLTree As Long
Box& = aol_findmailbox()
 If Box& = 0& Then Call aol_runmenubynum(1&, 0&, False)
  Do:
   DoEvents
   Box& = aol_findmailbox()
  Loop Until Box& <> 0
   DoEvents: DoEvents
 AOLTabControl& = FindWindowEx(Box&, 0&, "_AOL_TabControl", vbNullString)
 AOLTabPage& = FindWindowEx(AOLTabControl&, 0&, "_AOL_TabPage", vbNullString)
 AOLTree& = FindWindowEx(AOLTabPage&, 0&, "_AOL_Tree", vbNullString)
Mail_CountNew& = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
End Function

Function misc_randomnumber(max As Integer) As Integer
Dim a As Integer
 Randomize
 a = Int(finished * Rnd) + 1
 misc_randomnumber = a
End Function

Public Function mail_countsent() As Long
Dim Box As Long, AOLTabPage As Long, AOLTabControl As Long, AOLTree As Long
Box& = aol_findmailbox()
 If Box& = 0& Then Call aol_runmenubynum(1&, 0&, False)
  Do:
   DoEvents
   Box& = aol_findmailbox()
  Loop Until Box& <> 0
   DoEvents: DoEvents
 AOLTabControl& = FindWindowEx(Box&, 0&, "_AOL_TabControl", vbNullString)
 AOLTabPage& = FindWindowEx(AOLTabControl&, 0&, "_AOL_TabPage", vbNullString)
 AOLTabPage& = FindWindowEx(AOLTabControl&, AOLTabPage&, "_AOL_TabPage", vbNullString)
 AOLTabPage& = FindWindowEx(AOLTabControl&, AOLTabPage&, "_AOL_TabPage", vbNullString)
 AOLTree& = FindWindowEx(AOLTabPage&, 0&, "_AOL_Tree", vbNullString)
mail_countsent& = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function mail_countold() As Long
Dim Box As Long, AOLTabPage As Long, AOLTabControl As Long, AOLTree As Long
Box& = aol_findmailbox()
 If Box& = 0& Then Call mail_readold
  Do:
   DoEvents
   Box& = aol_findmailbox()
  Loop Until Box& <> 0
  DoEvents: DoEvents
 AOLTabControl& = FindWindowEx(Box&, 0&, "_AOL_TabControl", vbNullString)
 AOLTabPage& = FindWindowEx(AOLTabControl&, 0&, "_AOL_TabPage", vbNullString)
 AOLTabPage& = FindWindowEx(AOLTabControl&, AOLTabPage&, "_AOL_TabPage", vbNullString)
 AOLTree& = FindWindowEx(AOLTabPage&, 0&, "_AOL_Tree", vbNullString)
mail_countold& = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
End Function

Public Sub win_normalize(thewin As Long)
 Call ShowWindow(thewin&, SW_NORMAL)
End Sub

Public Sub win_show(thewin As Long)
 Call ShowWindow(thewin&, SW_SHOW)
End Sub

Public Sub win_minimize(thewin As Long)
 Call ShowWindow(thewin&, WM_MINIMIZE)
End Sub

Public Sub aol_killwait()
Dim AOLModal As Long, AOLModalVis As Long, aolicon As Long
Call aol_runmenubystring("&About America Online")
Do:
 DoEvents
 AOLModal& = FindWindow("_AOL_Modal", vbNullString)
 AOLModalVis& = IsWindowVisible(AOLModal&)
 Call pause(0.1)
Loop Until AOLModVis& = 1&
aolicon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(aolicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aol_runshortcut(theshortcut As Long)
Call aol_runmenubynum(7&, 4&, True, theshortcut& + 1&)
End Sub

Public Sub aol_signoff()
 Call aol_runmenubystring("&Sign Off")
End Sub

Public Sub aim_accept()
 If aim_findaccept() = 0& Then Exit Sub
 Call ClickIcon(FindWindowEx(aim_findaccept(), 0&, "_AOL_Icon", vbNullString))
End Sub

Public Function aim_findaccept() As Long
Dim aol As Long, mdi As Long, child As Long
Dim aoledit As Long, AOLStatic As Long, aolicon As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
aoledit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
aolicon& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
If aoledit& <> 0 And AOLStatic& <> 0 And aolicon& <> 0 And InStr(win_gettext(AOLStatic&), "Would you like to accept?") Then
   aim_findaccept& = child&
  Exit Function
Else
 Do:
  DoEvents
   child& = FindWindowEx(mdi&, aolchild&, "AOL Child", vbNullString)
   aoledit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
   AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
   aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
   aolicon& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
   aolicon& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
   aolicon& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
    If aoledit& <> 0 And AOLStatic& <> 0 And aolicon& <> 0 And InStr(win_gettext(AOLStatic&), "Would you like to accept?") Then
     aim_accept& = child&
    Exit Function
 Loop Until child& = 0&
aim_findaccept& = child&
End Function

Public Function chat_getname() As String
Dim theroom As Long
 theroom& = aol_findroom()
  If theroom& = 0& Then
   chat_getname$ = "room not found"
  Else
   chat_getname$ = win_gettext(theroom&)
  End If
End Function

Public Function aol_getcaption() As String
Dim aol As Long
 aol& = FindWindow("AOL Frame25", vbNullString)
  If aol& = 0& Then
   aol_getcaption$ = "aol not found"
  Else
   aol_getcaption$ = win_gettext(aol&)
  End If
End Function

Public Function combo_getcount(thecombo As Long) As Long
 combo_getcount& = SendMessage(thecombo&, CB_GETCOUNT, 0&, 0&)
End Function

Public Function txt_linecount(thetext As String) As Long
Dim Current As String, currentt As String, a As Long, Count As Long
If Len(thetext$) = 0& Then txt_linecount& = 0&: Exit Function
For a& = 1 To Len(thetext$)
 Current$ = Mid(thetext$, a&, 1&)
 currentt$ = Mid(thetext$, 1&, a&)
 If Current$ = Chr(13) Then Count& = Val(Count&) + 1&
Next a&
If Mid(thetext$, Len(thetext$), 1&) <> Chr(10) And Mid(thetext$, Len(thetext$), 1&) <> Chr(13) Then
 Count& = Val(Count&) + 1&
 txt_linecount& = Count&
End If
End Function

Public Function list_getcount(thelist As Long) As Long
 list_getcount& = SendMessage(thelist&, LB_GETCOUNT, 0&, 0&)
End Function

Public Sub win_maximize(thewindow As Long)
 Call ShowWindow(thewindow&, SW_MAXIMIZE)
End Sub

Sub form_center(TheForm As Form)
 TheForm.Top = (Screen.Height * 0.85) / 2 - TheForm.Height / 2
 TheForm.Left = Screen.Width / 2 - TheForm.Width / 2
End Sub

Public Sub chat_link(theurl As String, thename As String)
 aol_sendtext ("< a href=" & Chr(34) & "" & theurl$ & "" & Chr(34) & ">"" & thetext$ & ""</a>")
End Sub

Sub aol_addtempsn(thesn As String, replacedsn As String, aolpath As String)
Dim aolpaths As String
Screen.MousePointer = 11
Static m0226 As String * 40000, l9E68 As Long, l9E6A As Long
Dim l9E6C As Integer, l9E6E As Integer, l9E70 As Variant, l9E74 As Integer
 If UCase$(Trim$(thesn$)) = replacedsn$ Then: Exit Sub
  On Error GoTo ItsOver
ItsOver:
 Screen.MousePointer = 0
 Exit Sub
  If Len(thesn$) < 7 Then: Exit Sub
   replacedsn$ = replacedsn$ + String$(Len(thesn$) - 7, " ")
   Let aolpaths$ = (aolpath & "\idb\main.idx")
   Open aolpaths$ For Binary As #1
   l9E68& = 1
   l9E6A& = LOF(1)
    While l9E68& < l9E6A&
     m0226 = String$(16384, Chr$(0))
     Get #1, l9E68&, m0226
      While InStr(UCase$(m0226), UCase$(thesn$)) <> 0
       Mid$(m0226, InStr(UCase$(m0226), UCase$(thesn$))) = replacedsn$
      Wend
      Put #1, l9E68&, m0226
      l9E68& = l9E68& + 40000
     Wend
    Seek #1, Len(thesn$)
    l9E68& = Len(thesn$)
     While l9E68& < l9E6A&
     m0226 = String$(16384, " ")
     Get #1, l9E68&, m0226
      While InStr(UCase$(m0226), UCase$(thesn$)) <> 0
       Mid$(m0226, InStr(UCase$(m0226), UCase$(thesn$))) = replacedsn$
      Wend
     Put #1, l9E68&, m0226
     l9E68& = l9E68& + 16384
    Wend
   Close #1
  Screen.MousePointer = 0
 Resume Next
End Sub

Sub aol_addnewuser(thesn As String, aolpath As String)
Dim thesn2 As String, pathidx As String
Screen.MousePointer = 11
Static m0226 As String * 40000, l9E68 As Long, l9E6A As Long
Dim l9E6C As Integer, l9E6E As Integer, l9E70 As Variant, l9E74 As Integer
 If UCase$(Trim$(thesn$)) = "NEWUSER" Then: Exit Sub
  On Error GoTo ItsOver
ItsOver:
  Screen.MousePointer = 0
  Exit Sub
   If Len(thesn$) < 7 Then: Exit Sub
    thesn2$ = "NewUser" + String$(Len(thesn$) - 7, " ")
    Let pathidx$ = (aolpath & "\idb\main.idx")
    Open pathidx$ For Binary As #1
    l9E68& = 1
    l9E6A& = LOF(1)
     While l9E68& < l9E6A&
      m0226 = String$(40000, Chr$(0))
      Get #1, l9E68&, m0226
       While InStr(UCase$(m0226), UCase$(thesn$)) <> 0
         Mid$(m0226, InStr(UCase$(m0226), UCase$(thesn$))) = thesn2$
       Wend
      Put #1, l9E68&, m0226
      l9E68& = l9E68& + 40000
     Wend
    Seek #1, Len(thesn$)
    l9E68& = Len(thesn$)
     While l9E68& < l9E6A&
      m0226 = String$(40000, Chr$(0))
      Get #1, l9E68&, m0226
       While InStr(UCase$(m0226), UCase$(thesn$)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(thesn$))) = thesn2$
        Wend
       Put #1, l9E68&, m0226
       l9E68& = l9E68& + 40000
      Wend
    Close #1
    Screen.MousePointer = 0
   Resume Next
End Sub

Public Sub file_decomprotect(thepath As String, AppName As String)
Dim thefile As String, c As String
On Error Resume Next
If file_isexisting(thepath$) = False Then Exit Sub Else:
 thefile = FreeFile
 Open thepath$ For Binary As #thefile
 Cat = "."
 Seek #thefile, 25
 Put #thefile, , c
 Close #1
  If Err Then: Exit Sub
  'MsgBox "file protected"
End Sub

Public Sub pc_shutdown()
Dim EWX_SHUTDOWN, Msg As Long
 Msg = MsgBox("are you sure?", vbYesNo Or vbQuestion)
 If Msg = vbNo Then Exit Sub Else: Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub

Public Sub profile_get(thesn As String)
Dim aol As Long, mdi As Long, child As Long, getpro As Long
Dim aoledit As Long, aolicon As Long, Error As Long, pro As Long
Dim errorwin As Long, obk As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
 Call aol_opengetpro
  Do:
   DoEvents
   getpro& = FindWindowEx(mdi&, 0&, "AOL Child", "Get a Member's Profile")
   aoledit& = FindWindowEx(getpro&, 0&, "_AOL_Edit", vbNullString)
   aolicon& = FindWindowEx(getpro&, 0&, "_AOL_Icon", vbNullString)
  Loop Until getpro& <> 0& And aoledit& <> 0 And aolicon& <> 0
  DoEvents
   Call win_setfocus(getpro&)
   Call win_settext(aoledit&, thesn$)
   Call win_clickicon(aolicon&)
    Do:
     DoEvents
     errorwin& = FindWindow("#32770", "America Online")
     obk& = FindWindowEx(errorwin&, 0&, "Button", "OK")
     pro& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Profile")
    Loop Until errorwin& <> 0 And obk& <> 0 Or pro& <> 0
     If errorwin& Then
      Call win_setfocus(errorwin&)
      Call win_clickicon(obk&)
      Call win_closewin(getpro&)
      Exit Sub
     Else
       Call win_closewin(getpro&)
      Exit Sub
    End If
  Exit Sub
End Sub

Public Sub win_settext(thewindow As Long, thetext As String)
 Call SendMessageByString(thewindow&, WM_SETTEXT, 0&, thetext$)
End Sub

Public Sub win_changecaption(thwindow As Long, thecaption As String)
 Call SendMessageByString(thewindow&, WM_SETTEXT, 0&, thecaption$)
End Sub

Public Sub win_hitenter(thewin As Long)
 Call SendMessage(thewin&, VK_RETURN, 0&, 0&)
End Sub

Public Sub win_clickicon(theicon As Long)
 Call PostMessage(theicon&, WM_LBUTTONDOWN, 0&, 0&)
 Call PostMessage(theicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub win_setfocus(thewin As Long)
 Call SetFocus(thewin&)
End Sub

Public Sub im_sendim(thesn As String, themsg As String)
Dim aol As Long, mdi As Long, aoledit As Long, richcntl As Long
Dim aolicon As Long, errorwin As Long, obk As Long, theim As Long
Dim i As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
 Call aol_openim
  Do:
   DoEvents
   theim& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
   aoledit& = FindWindowEx(theim&, 0&, "_AOL_Edit", vbNullString)
   richcntl& = FindWindowEx(theim&, 0&, "RICHCNTL", vbNullString)
   aolicon& = FindWindowEx(theim&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 8&
     aolicon& = FindWindowEx(theim&, aolicon&, "_AOL_Icon", vbNullString)
    Next i&
  Loop Until theim& <> 0& And aoledit& <> 0 And aolicon& <> 0 And richcntl& <> 0
  DoEvents
   Call win_setfocus(theim&)
   Call win_settext(aoledit&, thesn$)
   Call win_settext(richcntl&, themsg$)
   Call win_clickicon(aolicon&)
    Do:
     DoEvents
     errorwin& = FindWindow("#32770", "America Online")
     obk& = FindWindowEx(errorwin&, 0&, "Button", "OK")
    Loop Until errorwin& <> 0 And obk& <> 0 Or theim& = 0
     If errorwin& Then
      Call win_setfocus(errorwin&)
      Call win_clickicon(obk&)
      Call win_closewin(theim&)
      Exit Sub
     Else
      Exit Sub
    End If
  Exit Sub
End Sub

Public Function tos_chatphrase1() As String
Dim num As Integer
num = misc_randomnumber(20)
 If num = 1 Then: tos_chatphrase1$ = "º¯`v´¯¯) PhrostByte By: Progee (¯¯`v´¯º"
 If num = 2 Then: tos_chatphrase1$ = ".·´¯`·-  gøthíc nightmâres by másta  ­·´¯`·" & Chr(13) & "·._.--   aøl 4.o punt tools · loaded ---._.·"
 If num = 3 Then: tos_chatphrase1$ = "· úpr mácro stùdio · másta ·"
 If num = 4 Then: tos_chatphrase1$ = "-=·Sting Anti Punta 2.o Loaded·=-" & Chr(13) & "-=·MaDe By SaBrE·=-"
 If num = 5 Then: tos_chatphrase1$ = "(¯\_ GøDZîLLa³·º _/¯)" & Chr(13) & "(¯\_ ßy ÇoLd _/¯)"
 If num = 6 Then: tos_chatphrase1$ = "¢º°¤÷®ÍP§ 2øøø÷¤°º¢" & Chr(13) & "¢º°¤÷£øÃdÊD÷¤°º¢"
 If num = 7 Then: tos_chatphrase1$ = "-•(`(`·•Fate Zero v¹ Loaded•·´)´)•-"
 If num = 8 Then: tos_chatphrase1$ = "•·.·´).·÷•[ Outlaw Mass Mailer by Twiztid"
 If num = 9 Then: tos_chatphrase1$ = "(¯`·.····÷• ärméñïå¹ · kðkô" & Chr(13) & "(¯`·.····÷• îøâdèd"
 If num = 10 Then: tos_chatphrase1$ = "•·._.·´¯`·>AoL 4.0 TooLz By: X GeNuS X" & Chr(13) & "•·._.·´¯`·>Status: LoaDeD" & Chr(13) & "•·._.·´¯`·>Ya'll BeTTa NoT MeSS WiT ThiS NiG!"
 If num = 11 Then: tos_chatphrase1$ = "^····÷• James Bond Toolz Ver .007" & Chr(13) & "^····÷• By: Saßan"
 If num = 12 Then: tos_chatphrase1$ = "(¯`•Prophecy²·° Loaded"
 If num = 13 Then: tos_chatphrase1$ = "···÷••(¯`·._ CoRn Fader _.·´¯)••÷···" & Chr(13) & "···÷••(¯`·._Created by :::PooP:::_.·´¯)••÷···"
 If num = 14 Then: tos_chatphrase1$ = "Blue Ice Punter¹ For AOL 4.0" & Chr(13) & "By STaNK"
 If num = 15 Then: tos_chatphrase1$ = "¤-----==America Onfire Platinum" & Chr(13) & "¤-----==Created (²›y Fatal Error"
 If num = 16 Then: tos_chatphrase1$ = "¤¤†³¹¹º†¤¤ SANNMEN †oºLz ¤¤†³¹¹º†¤¤" & Chr(13) & "¤¤†³¹¹º†¤¤    By:Má§†é®MinÐ    ¤¤†³¹¹º†¤¤" & Chr(13) & "¤¤†³¹¹º†¤¤ LOADED ¤¤†³¹¹º†¤¤"
 If num = 17 Then: tos_chatphrase1$ = "··¤÷×(Rapier Bronze)×÷¤··" & Chr(13) & "··¤÷×(By Excalibur)×÷¤··" & Chr(13) & "··¤÷×(Works for 3.0 and 4.0!!!)×÷¤··"
 If num = 18 Then: tos_chatphrase1$ = "<-==(`(` Icy Hot 2.0 For AOL 4.0 ')')==->" & Chr(13) & "<-==(`(` Loaded ')')==->"
 If num = 19 Then: tos_chatphrase1$ = "(\›•‹ Im Backfire KiLLer ›•‹/)" & Chr(13) & "(\›•‹ By:phire Status:Loaded ›•‹/)" & Chr(13) & "(\›•‹ Im Backfire KiLLer ›•‹/)"
 If num = 20 Then: tos_chatphrase1$ = "[_.·´¯° Indian Invasion Punter Loaded °¯`·._]"""
End Function


