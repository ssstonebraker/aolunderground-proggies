Attribute VB_Name = "Utils³"
Option Explicit
'July 30, 2000
'hey all
'this is naïve again. this is an update of my original bas's, Utils¹ and Utils²
'it has a large number of new subs, and some are getting the axe
'sorry bout any errors and old bad style, this bas was first started over 2 years ago
'i did compile it and i found all the errors there and fixed them, so this bas can now be compiled
'also note the option explicit up there, all vars in this bas are defined! lol.
'sorry i left that out in the early vers, but i hadn't been proggin long

'this update has 201 total subs with 72 new subs and many revised ones.
'i tried to remember to mark the new subs, but i know i forgot some. see if you notice any i didnt mark.
'nearly 80% of the subs in the bas were changed a little bit. it needed it. this bas
'should run much better now!
'check out my superfast CountLines sub and the SaveUnlimitedLists/LoadUnlimitedLists,
'i think they are the best and most useful

'as far as i know, the code in this bas is 99.5% original, i think there is one sub by monkegod in here, any i copied is marked
' (there is very litle)
'i coded it from the declarations on
'it has been a product of 2+ years work. i think the original utils bas was made about a year
'after i began proggin, 'more than 2+ years ago, i dont know the exact date. i have no problem with people
'using subs from this bas in their own work, but you must give me credit. thats the way things should work.

'im very proud of this bas, so if you use it or  if you like it contact me
'aim: lazy naive
'email: george@george.cx
'feel free to this file, just email me and give me the url.
'peace-
'naïve

'Peace To Unity!
'---------------------------------------
'old intro text -
'---------------------------------------
' Bas By Naïve
' Version: 3.01
'
' This Bas Is Full Of Stuff To Help Yah
' Out, Utility Stuff...
'
' Give me credit if ya
' bite my codes
'
' Peace Out,
' -= Naïve =-


Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function Gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const MOVE = &HA1
Public Const LB_SETHORIZONTALEXTENT = &H194

Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type


Public Type POINTAPI
        x As Long
        y As Long
End Type


Function Anagram(str As String) As String
'*new sub!
'scrambles a word. completely random. this actually took a while.

On Error Resume Next

Dim x As Integer
Dim NewWord As String
Dim rn As Integer
Dim e1 As String
Dim e2 As String
Dim e3 As String


For x = 1 To Len(str) * 4 '  this should be enough, if its not scrambling well enough, then increase this value ;)
eek:
    rn = Rnd * Len(str) + 1
    If rn = 1 Then GoTo eek
    
    e1 = Mid(str, rn, 1)
    e2 = Left(str, rn - 1)
    e3 = Right(str, Len(str) - rn)
    
If Len(e1 & e2 & e3) > Len(str) Then
GoTo eek
End If
    str = e1 & e2 & e3
Next x

Anagram = str

End Function

Function ArrayToString(arrayx() As Variant) As String
'*new sub!
'this is ez, but if you dont know split/join it helps alot
'adds each item in the array to a string, with a newline between each item
ArrayToString = Join(arrayx(), Chr(13) & Chr(10))
End Function

Public Sub CD_Close()
'*new sub!
'closes the cd drive
    Call MciSendString("set CDAudio door closed", vbNullString, 0&, 0&)
End Sub
Public Sub CD_Open()
'*new sub!
'closes the cd drive
    Call MciSendString("set CDAudio door open", vbNullString, 0&, 0&)
End Sub

Sub Bas_ContactAndWebsite()
' If you want to get ahold of me, my email is:
' george@george.cx
' or visit my site at www.george.cx

' Later ,
' Naïve

End Sub

Sub Bas_Greetz()
'Greetz To:
'
' JaGGeD, TaRBaS, Dragon, SouL, Kim, Mystical
' BMXBee, Nusance, Step, SeeN, RaouL, OtisP,
' Druso, WolfMagicX, and everyone at Monolithic

'added in utils³:
'flex, spook, nolia, kid, wgf, ace, whitey, dert, spiny, cenes
'scooby, jedi, bis, spot, tiz, cali, taco, raider, hvrp
'phonik, gavin, blue torpedo, everyone at sektor and project-x
'all my jungalistic people
'and all the rest in unity!!!!!!
'-peace-


'
' Thanks To Anyone Who Posts or Uses My Files!
'

End Sub

Sub Bas_MailMe()
' Email me at naivexistence@iname.com
End Sub




Function ChangeCase(str As String) As String
'*new sub!
'changes the case of every letter in a string
Dim x As Integer
Dim GetChr As String
Dim newstr As String

For x = 1 To Len(str)
    GetChr = Mid(str, x, 1)
    If IsUppercase(GetChr) = True Then
        GetChr = LCase(GetChr)
    Else
        GetChr = UCase(GetChr)
    End If
newstr = newstr & GetChr
Next x

ChangeCase = newstr
End Function

Function ComputePercent(Part As Integer, Whole As Integer) As Integer
'*new sub!
'this is in case you didnt know how to do it / for quickness

ComputePercent = Round(Part / Whole, 0)


End Function

Function HTMLSwitchBold(str As String) As String
'*new sub
'this sub is really useful...and a helluva lot harder to code then you'd think!


Dim iSpot As Integer
Dim iSpot2 As Integer
Dim sTest As String
Dim AddTag As Boolean

sTest = Left(str, 3)
If sTest = "<b>" Or sTest = "</b" Then
    AddTag = False
Else
    AddTag = True
End If


iSpot = 0
Sartre:
iSpot = InStr(iSpot + 1, str, "<")
sTest = Mid(str, iSpot + 1, 2)
If sTest = "b>" Then
    str = Left(str, iSpot) & "/" & Mid(str, iSpot + 1)
ElseIf sTest = "/b" Then
    str = Left(str, iSpot) & Mid(str, iSpot + 2)
End If

If iSpot > 0 And iSpot < Len(str) Then GoTo Sartre

iSpot = InStr(1, str, "<b>")
iSpot2 = InStr(1, str, "</b>")

If AddTag = False Then GoTo Rehab

If iSpot > iSpot2 Then
    str = "<b>" & str
Else
    str = "</b>" & str
End If

Rehab:

HTMLSwitchBold = str
End Function

Function HTMLSwitchItalic(str As String) As String
'*new sub
'this sub is really useful...and a helluva lot harder to code then you'd think!


Dim iSpot As Integer
Dim iSpot2 As Integer
Dim sTest As String
Dim AddTag As Boolean

sTest = Left(str, 3)
If sTest = "<i>" Or sTest = "</i" Then
    AddTag = False
Else
    AddTag = True
End If


iSpot = 0
Sartre:
iSpot = InStr(iSpot + 1, str, "<")
sTest = Mid(str, iSpot + 1, 2)
If sTest = "i>" Then
    str = Left(str, iSpot) & "/" & Mid(str, iSpot + 1)
ElseIf sTest = "/i" Then
    str = Left(str, iSpot) & Mid(str, iSpot + 2)
End If

If iSpot > 0 And iSpot < Len(str) Then GoTo Sartre

iSpot = InStr(1, str, "<i>")
iSpot2 = InStr(1, str, "</i>")

If AddTag = False Then GoTo Rehab

If iSpot > iSpot2 Then
    str = "<i>" & str
Else
    str = "</i>" & str
End If

Rehab:

HTMLSwitchItalic = str
End Function
Function HTMLSwitchUnderline(str As String) As String
'*new sub
'this sub is really useful...and a helluva lot harder to code then you'd think!


Dim iSpot As Integer
Dim iSpot2 As Integer
Dim sTest As String
Dim AddTag As Boolean

sTest = Left(str, 3)
If sTest = "<u>" Or sTest = "</u" Then
    AddTag = False
Else
    AddTag = True
End If


iSpot = 0
Sartre:
iSpot = InStr(iSpot + 1, str, "<")
sTest = Mid(str, iSpot + 1, 2)
If sTest = "u>" Then
    str = Left(str, iSpot) & "/" & Mid(str, iSpot + 1)
ElseIf sTest = "/u" Then
    str = Left(str, iSpot) & Mid(str, iSpot + 2)
End If

If iSpot > 0 And iSpot < Len(str) Then GoTo Sartre

iSpot = InStr(1, str, "<u>")
iSpot2 = InStr(1, str, "</u>")

If AddTag = False Then GoTo Rehab

If iSpot > iSpot2 Then
    str = "<u>" & str
Else
    str = "</u>" & str
End If

Rehab:

HTMLSwitchUnderline = str
End Function


Function HTMLToggleBold(str As String) As String
'*new sub
'makes it alternate between chars


str = Replace(str, "<b>", "")
str = Replace(str, "</b>", "")

Dim sOne As String
Dim sNew As String
Dim bOn As Boolean
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, 1)
    If bOn Then
        sOne = "<b>" & sOne
    Else
        sOne = "</b>" & sOne
    End If
    sNew = sNew + sOne
Next x
   
HTMLToggleBold = sNew
End Function

Function HTMLToggleItalic(str As String) As String
'*new sub
'makes it alternate between chars


str = Replace(str, "<i>", "")
str = Replace(str, "</i>", "")

Dim sOne As String
Dim sNew As String
Dim bOn As Boolean
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, 1)
    If bOn Then
        sOne = "<i>" & sOne
    Else
        sOne = "</i>" & sOne
    End If
    sNew = sNew + sOne
Next x
HTMLToggleItalic = sNew
End Function

Function HTMLToggleUnderline(str As String) As String
'*new sub
'makes it alternate between chars


str = Replace(str, "<u>", "")
str = Replace(str, "</u>", "")

Dim sOne As String
Dim sNew As String
Dim bOn As Boolean
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, 1)
    If bOn Then
        sOne = "<u>" & sOne
    Else
        sOne = "</u>" & sOne
    End If
    sNew = sNew + sOne
Next x
   
HTMLToggleUnderline = sNew
End Function


Function HTMLToggleStrike(str As String) As String
'*new sub
'makes it alternate between chars


str = Replace(str, "<strike>", "")
str = Replace(str, "</strike>", "")

Dim sOne As String
Dim sNew As String
Dim bOn As Boolean
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, 1)
    If bOn Then
        sOne = "<strike>" & sOne
    Else
        sOne = "</strike>" & sOne
    End If
    sNew = sNew + sOne
Next x
   
HTMLToggleStrike = sNew
End Function


Function HTMLWavy(str As String) As String
'*new sub!
'this makes the text look wavy, youve seen this before

'you may want to u comment these lines to kill returns, your call
'str = Replace(str, Chr(13), "")
'str = Replace(str, Chr(10), "")

Dim sOne As String
Dim sNew As String
Dim iState As Integer
Dim sTag As String
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, x)
    iState = iState + 1
    If iState > 3 Then iState = 0
    Select Case iState
        Case 0
            sTag = "<sup>str</sup>"
        Case 1
            sTag = "str"
        Case 2
            sTag = "<sub>str</sub>"
        Case 3
            sTag = "str"
    End Select
    sOne = Replace(sTag, "str", sOne)
    sNew = sNew & sOne
Next x
HTMLWavy = sNew
End Function

Function HTMLKillWavy(str) As String
'*new sub
'i always like to include a way to undo what i do ;)

str = Replace(str, "<sub>", "")
str = Replace(str, "</sub>", "")
str = Replace(str, "<sup>", "")
str = Replace(str, "</sup>", "")
HTMLKillWavy = str
End Function


Function HTMLKillFormat(str) As String
'*new sub
'i always like to include a way to undo what i do ;)
str = HTMLKillWavy(str)


str = Replace(str, "<b>", "")
str = Replace(str, "</b>", "")
str = Replace(str, "<u>", "")
str = Replace(str, "</u>", "")
str = Replace(str, "<i>", "")
str = Replace(str, "</i>", "")
str = Replace(str, "<strike>", "")
str = Replace(str, "</strike>", "")
HTMLKillFormat = str
End Function



Sub Label_ScrollString(str As String, lab As Label)
Dim iLines As Integer
Dim sCaption As String
Dim x As Integer

sCaption = lab.Caption
If Right(str, 2) <> Chr(13) & Chr(10) Then str = str & Chr(13) & Chr(10)

iLines = LineCount(str)

Dim aLines() As String
ReDim aLines(iLines)

aLines = String_AddLinesToArray(str)
For x = 0 To iLines - 1
    lab.Caption = aLines(x)
    TimeOutX (1)
Next x
lab.Caption = sCaption
End Sub

Sub List_AddPrefix(prfx As String, lst As ListBox)
'new sub! adds a prefix to every item in a listbox!
Dim x As Integer

For x = 0 To lst.ListCount - 1
    lst.List(x) = prfx & lst.List(x)
Next x

End Sub

Sub List_AddSuffix(sfx As String, lst As ListBox)
'new sub! adds a suffix to every item in a listbox!
Dim x As Integer

For x = 0 To lst.ListCount - 1
    lst.List(x) = lst.List(x) & sfx
Next x

End Sub

Function ListToArray(lst As ListBox) As String()
Dim AX() As String
ReDim AX(lst.ListCount - 1)

Dim x As Integer

For x = 0 To lst.ListCount - 1
    AX(x) = lst.List(x)
Next x

ListToArray = AX
End Function


Sub PicBox_ScrollString(str As String, pBox As PictureBox)
Dim iLines As Integer
Dim sCaption As String

If Right(str, 2) <> Chr(13) & Chr(10) Then str = str & Chr(13) & Chr(10)

iLines = LineCount(str)

Dim aLines() As String
ReDim aLines(iLines)

aLines = String_AddLinesToArray(str)
Dim x As Integer

For x = 0 To iLines - 1
    pBox.Cls
    pBox.Print aLines(x)
    TimeOutX (1)
Next x
pBox.Cls
End Sub
Function List_CombineTwo(ListA As ListBox, ListB As ListBox, ToList As ListBox)
'*new sub!
'adds ListA and ListB to ToList!

If ToList.Name <> ListA.Name And ToList.Name <> ListB.Name Then
    ToList.Clear
End If

Dim x As Integer

If ToList.Name <> ListA.Name Then
    For x = 0 To ListA.ListCount - 1
        ToList.AddItem ListA.List(x)
    Next x
End If

If ToList.Name <> ListB.Name Then
    For x = 0 To ListB.ListCount - 1
        ToList.AddItem ListB.List(x)
    Next x
End If

End Function

Function List_CompleteWord(PartStr As String, ListLook As ListBox) As String
'*new sub
'looks thru a list to complete a word.
'this is the kind of thing that a lot of aol x'ers do :P


Dim x As Integer

For x = 0 To ListLook.ListCount - 1
    If InStr(1, ListLook.List(x), PartStr) Then
        List_CompleteWord = ListLook.List(x)
        Exit Function
    End If
Next x

        
    
End Function

Function List_Search(LookFor As String, InList As ListBox, Optional IsCaseSensitive As Boolean) As String
'*new sub!
'looks thru a list and returns all of the items that match the criteria
'reference: in the instr function and in general, vbTextCompare means not CaseSensitive, and vbBinary Compare is case-sensitive :P
'
Dim sFound As String
Dim x As Integer



For x = 0 To InList.ListCount - 1
    If IsCaseSensitive = False Then
        If InStr(1, InList.List(x), LookFor, vbTextCompare) > 0 Then sFound = sFound & ", " & InList.List(x)
    Else
        If InStr(1, InList.List(x), LookFor, vbBinaryCompare) > 0 Then sFound = sFound & ", " & InList.List(x)
    End If
Next x

If sFound <> "" Then List_Search = Mid(sFound, 3)
   
        
End Function


Function NumChar(str As String, LookFor As String)
'*new sub!
'count the number of times the lookfor appears in the str!

NumChar = Len(str) - Len(Replace(str, LookFor, ""))
End Function

Sub PicBox_ShowFonts(pBox As PictureBox)
'*new sub!
'writes the name of every font, in that font on the picturebox
Dim x As Integer
Dim sFont As String

For x = 0 To Screen.FontCount - 1
    sFont = Screen.Fonts(x)
    pBox.Font = sFont
    pBox.Print sFont
    DoEvents
Next x



End Sub

Function RandomCaps(str As String)
'*new sub!
'every letter is randomly upper or lowercase!

Dim bCap As Boolean
Dim sOne As String
Dim sNew As String
Dim x As Integer
Dim iRN As Integer


For x = 1 To Len(str)
    iRN = RandomNumber(1, 2)
    bCap = ReadBoolStr(CStr(iRN))
    sOne = Mid(str, x, 1)
    If bCap Then
        sOne = UCase(sOne)
    Else
        sOne = LCase(sOne)
    End If
    sNew = sNew & sOne
Next x
RandomCaps = sNew


End Function

Function RandomDecoration(str As String) As String
'*new sub!
'for each cahr in the string, it gives it an html format look (bold, italic, underline, or nothing)

Dim sBold As String
Dim sItalic As String
Dim sUnder As String
Dim sConv As String
Dim sNew As String
Dim sOne As String

sBold = "<b>str</b>"
sItalic = "<i>str</i>"
sUnder = "<u>str</u>"
Dim x As Integer

For x = 1 To Len(str)
    Randomize Timer
    Dim rn As Integer
    rn = Rnd * 4 + 1
    If rn = 0 Then rn = 1
    If rn = 5 Then rn = 4

    sOne = Mid(str, x, 1)
    
    Select Case rn
        Case 1
            sConv = sBold
        Case 2
            sConv = sItalic
        Case 3
            sConv = sUnder
        Case 4
            sConv = sOne
    End Select
    
    sNew = sNew & Replace(sConv, "str", sOne)
Next x
RandomDecoration = sNew
End Function

Function RandomWave(str As String) As String
'*new sub!
'for each char in the string, it gives it a different script point (superscrpit, subscript, normal)

Dim sUp As String
Dim sDown As String
Dim sConv As String
Dim sNew As String
Dim sOne As String

sUp = "<sup>str</sup>"
sDown = "<sub>str</sub>"

Dim x As Integer

For x = 1 To Len(str)
    Randomize Timer
    Dim rn As Integer
    rn = Rnd * 3 + 1
    If rn = 0 Then rn = 1
    If rn = 4 Then rn = 4

    sOne = Mid(str, x, 1)
    
    Select Case rn
        Case 1
            sConv = sUp
        Case 2
            sConv = sDown
        Case 3
            sConv = sOne
    End Select
    
    sNew = sNew & Replace(sConv, "str", sOne)
Next x
RandomWave = sNew
End Function

Function RandomHTMLFont(str As String) As String
'*new sub!
'for each char in the string, it gives it a different font! uses my gneric list :P
'this ends up looking pretty neato, like a friggin ransom note or something

Dim sConv As String
Dim sNew As String
Dim sOne As String
Dim sUp As String, sDown As String
sUp = "<sup>str</sup>"
sDown = "<sub>str</sub>"

Dim x As Integer

For x = 1 To Len(str)
    sConv = "<font face=""" & RandomFont & """>str</font>"
    
    sOne = Mid(str, x, 1)
    
    sNew = sNew & Replace(sConv, "str", sOne)
Next x
RandomHTMLFont = sNew
End Function


Function String_RemoveLCase(str As String) As String
'*new sub
'removes all the lowercase chars. i had a program do this once, so here you go

Dim sNew As String
Dim sOne As String
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, x)
    If IsUppercase(sOne) = False Then sNew = sNew & sOne
Next x

String_RemoveLCase = sNew
End Function

Function String_RemoveUCase(str As String) As String
'*new sub
'removes all the uppercase chars. i had a program do this once, so here you go

Dim sNew As String
Dim sOne As String
Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, x)
    If IsUppercase(sOne) = True Then sNew = sNew & sOne
Next x

String_RemoveUCase = sNew
End Function


Sub SaveUnlimitedLists(TheLists() As ListBox, Path As String)
'*new sub!
'this sub is extremely useful
'just pass it an array of listboxes and it will save all of them to the file you pass it. this is one of my favorite subs int he bas

Dim TotalNumOfLists As Integer
TotalNumOfLists = UBound(TheLists)
    
    Dim SaveList As Long
    On Error Resume Next
        
Dim LNum As Integer
Dim LIndex As Integer
Open Path$ For Output As #1
    
For LNum = 0 To TotalNumOfLists - 1
    
    For LIndex = 0 To TheLists(LNum).ListCount - 1
        Print #1, TheLists(LNum).List(LIndex)
    Next LIndex

    Print #1, "*&*&* " & LNum

Next LNum

Close #1
End Sub

Sub LoadUnlimitedLists(TheLists() As ListBox, Path As String)
'*new sub!
'this is really cool - if you pass it an array of listboxes (TheLists()), and the path it will load them all. As many as you want.

Dim TotalNumOfLists As Integer
TotalNumOfLists = UBound(TheLists)

Dim SaveList As Long

On Error Resume Next
        
Dim LNum As Integer
Dim LoadString As String * 20000
Dim LIndex As Integer

Open Path$ For Input As #1
TheLists(0).Clear

For LNum = 0 To TotalNumOfLists - 1
    
    While Not EOF(1)
        Input #1, LoadString
        DoEvents
        If Left(LoadString, 6) = "*&*&* " Then
            LIndex = CInt(Mid(LoadString, 7)) + 1
            TheLists(LIndex).Clear
        Else
            TheLists(LIndex).AddItem LoadString
        End If
    Wend
    Close #1
    
Next LNum

Close #1
End Sub


Function String_AddLinesToArray(str As String) As String()
'*new sub!
'this will go thru and add every line to an array and return it!
'just make sure you dump the return val into an array! :P
'Ex:
'Dim LineArray () as String
'LineArray = String_AddLinesToArray(MyString)
'get it? cool. also this is a good way to see how you return an array from a function :)

str = Replace(str, Chr(13) & Chr(10), Chr(13))
String_AddLinesToArray = Split(str, Chr(13))

'now you can get any line from your string easily!

End Function

Function String_AddCharsToArray(str As String) As String()
'*new sub!
'this will go thru and add every char to an array and return it!
'just make sure you dump the return val into an array! :P
'Ex:
'Dim LineArray () as String
'LineArray = String_AddCharsToArray(MyString)

Dim arrayx() As String
ReDim arrayx(Len(str))

Dim x As Integer

For x = 1 To Len(str)
    arrayx(x) = Mid(str, x, 1)
Next x



String_AddCharsToArray = arrayx

'now you can get any char from your string easily!

End Function


Sub ArrayCopy(oldCopy(), newCopy())
'*new sub!
'i just put this is here because i didnt know this was possible, and i figured you might now either
newCopy = oldCopy
End Sub

Function String_FirstLine(str As String) As String
Dim iSpot As Integer

iSpot = InStr(1, str, Chr(13) & Chr(10))
String_FirstLine = Left(str, iSpot - 1)
End Function

Function String_LastLine(str As String) As String
Dim iSpot As Integer

If Right(str, 2) = Chr(13) & Chr(10) Then str = Left(str, Len(str) - 2)

iSpot = InStrRev(1, str, Chr(13) & Chr(10))

String_LastLine = Mid(str, iSpot + 2)
End Function


Function String_Load(Path As String) As String
'*new sub!
'simple - just loads the text from a file
Dim TempString As String
On Error Resume Next 'you just kind of have to use this in file ops. otherwise avoid it, it screws up debugging :)
Open Path$ For Input As #1
    TempString = Input(LOF(1), #1)
Close #1

String_Load = TempString
End Function

Sub String_Save(str As String, Path As String)
'*new sub
'Simple but really useful
On Error Resume Next
Open Path$ For Output As #1  'you just kind of have to use this in file ops. otherwise avoid it, it screws up debugging :)
    Print #1, str$
Close #1
End Sub

Function String_KillHTML(str As String) As String
'*newsub!
'this is real simple, removes anything inbetween "<" and ">" so it may overkill it. sorry :P

Dim sOne As String
Dim nNew As String
Dim InTag As Boolean
Dim x As Integer
Dim sNew As String

For x = 1 To Len(str)
    sOne = Mid(str, x, 1)
    If sOne = "<" Then InTag = True
    If InTag = False Then sNew = sNew & sOne
    If sOne = ">" Then InTag = False
Next x
String_KillHTML = sNew

End Function

Function String_GetHTML(str As String) As String
'*newsub!
'this is real simple, saves anything in between "<" and ">" so it should pull out the HTML tags! ;)

Dim sOne As String
Dim nNew As String
Dim InTag As Boolean
Dim x As Integer
Dim sNew As String

For x = 1 To Len(str)
    sOne = Mid(str, x, 1)
    If sOne = "<" Then InTag = True
    If InTag = True Then sNew = sNew & sOne
    If sOne = ">" Then InTag = False
Next x
String_GetHTML = sNew

End Function


Function String_SplitLettersAndNumbers(str As String) As String
'*new sub!
'moves all the etters to the back of the str!
' ex:
'r33to becomes rto33

Dim sLets As String
Dim sNums As String
Dim sOne As String

Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, x)
    If IsNumeric(sOne) = True Then
        sNums = sNums & sOne
    Else
        sLets = sLets & sOne
    End If
Next x
String_SplitLettersAndNumbers = sLets & sNums
End Function
Function String_ExtractNumbers(str As String) As Integer
'*new sub!
'pulls the numbers out of a string and returns them as an integer


Dim sNums As String
Dim sOne As String

Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, x)
    If IsNumeric(sOne) = True Then
        sNums = sNums & sOne
    End If
Next x
String_ExtractNumbers = CInt(sNums)
End Function

Function String_ExtractLetters(str As String) As String
'*new sub!
'pulls the letters out of a string and returns them as a string


Dim sLets As String
Dim sOne As String

Dim x As Integer

For x = 1 To Len(str)
    sOne = GetChar(str, x)
    If IsNumeric(sOne) = False Then
        sLets = sLets & sOne
    End If
Next x
String_ExtractLetters = sLets
End Function


Function ToggleCase(str As String) As String
'*new sub!
'changes the case on every other letter in a string
Dim x As Integer
Dim GetChr As String
Dim newstr As String
Dim WasCap As Boolean
For x = 1 To Len(str)
    GetChr = Mid(str, x, 1)
    If WasCap = True Then
        GetChr = LCase(GetChr)
    Else
        GetChr = UCase(GetChr)
    End If
    WasCap = Not WasCap
newstr = newstr & GetChr
Next x
ToggleCase = newstr
End Function



Function DivisionRemainder(Total As Integer, DivideBy As Integer) As Integer
Do While Total > DivideBy
Total = Total - DivideBy
Loop
DivisionRemainder = Total
End Function

Function FileExists(Path As String) As Boolean
Dim e As String
e = DiR(Path)
If e <> "" Then
    FileExists = True
Else
    FileExists = False
End If
End Function

Function FindLastChar(str As String, TheChar As String)
'finds the last instance of a char,
'most peeps dont know abou InStrRev
'i think its vb6 only
FindLastChar = InStrRev(str, TheChar)
End Function

Sub Form_CenterAt(Frm As Form, x As Single, y As Single)
'*new sub

Frm.MOVE x - Frm.Width / 2, y - Frm.Height / 2


End Sub

Sub Form_ExitPackUp(Frm As Form)
'*new sub
'i like this  a lot
'it makes the caption first suck in, then the width, then the height. see it in action



Dim OC As String
Dim OH As Integer
Dim OW As Integer

OC = Frm.Caption

Dim x As Integer

For x = Len(Frm.Caption) To 1 Step -1
    Frm.Caption = Left(OC, x)
    TimeOutX (0.01)
Next x
Frm.Caption = ""


Do While Frm.Width > 31 And Frm.Width <> OW
    OW = Frm.Width
    Frm.Width = Frm.Width - 50
    DoEvents
Loop

Do While Frm.Height > 31 And Frm.Height <> OH
    OH = Frm.Height
    Frm.Height = Frm.Height - 50
    DoEvents
Loop
    


End



End Sub


Sub Form_Pulse(Frm As Form, Optional NuMPuLSeS As Integer)
If NuMPuLSeS = 0 Then NuMPuLSeS = 5
Dim x As Integer, y As Integer, w As Integer

On Error Resume Next
' &H000000FF& = Red
Frm.BackColor = zGetRGB(&HFF&).Red
x = zGetRGB(&HFF&).Red
y = x
y = y - 10
w = 10
x = 255
y = x
Do While NuMPuLSeS > 0
Do While y > 100
Frm.BackColor = RGB(y - 10, 0, 0)
TimeOutX (1E-18)
y = y - 15
Loop
Do While y < 255
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = y + 15
Loop
NuMPuLSeS = NuMPuLSeS - 1
Loop
End Sub
Sub Form_Beat(Frm As Form, Optional NuMPuLSeS As Integer, Optional BeatSpeed As Integer)
Dim y As Integer

If NuMPuLSeS = 0 Then NuMPuLSeS = 5
If BeatSpeed = 0 Then BeatSpeed = 7
If BeatSpeed > 10 Then BeatSpeed = 10
On Error Resume Next
y = 255
Do While NuMPuLSeS > 0
Do While y < 255
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = y + 15
Loop
Do While y > 100
Frm.BackColor = RGB(y - 10, 0, 0)
TimeOutX (0.00000000001)
y = y - 15
Loop

NuMPuLSeS = NuMPuLSeS - 1
TimeOutX (BeatSpeed)
Loop

End Sub

Function CountLetters(TestStr As String) As Integer
'*new sub!
Dim x As Integer, TheCount As Integer

For x = 1 To Len(TestStr)
    If IsNumeric(Mid(TestStr, x, 1)) = False Then TheCount = TheCount + 1
Next x

CountLetters = TheCount

End Function

Function CountNumbers(TestStr As String) As Integer
'*new sub!

Dim x As Integer, TheCount As Integer

For x = 1 To Len(TestStr)
    If IsNumeric(Mid(TestStr, x, 1)) = True Then TheCount = TheCount + 1
Next x

CountNumbers = TheCount

End Function


Function DecorateText(TextStr As String, BoldOn As Boolean, UnderlineOn As Boolean, ItalicOn As Boolean, StrikeOn As Boolean, CapsOn As Boolean)
'*new sub!

'does a helluva lot of text manipultion and can be easily extended :)
'makes the first char of each word the above options

Dim LastSpace As Boolean
Dim newstr As String
Dim AddBeg As String, AddEnd As String

If BoldOn = True Then
    AddBeg = AddBeg & "<b>"
    AddEnd = AddEnd & "</b>"
End If

If UnderlineOn = True Then
    AddBeg = AddBeg & "<u>"
    AddEnd = AddEnd & "</u>"
End If

If StrikeOn = True Then
    AddBeg = AddBeg & "<strike>"
    AddEnd = AddEnd & "</strike>"
End If

If ItalicOn = True Then
    AddBeg = AddBeg & "<i>"
    AddEnd = AddEnd & "</i>"
End If

If Left(TextStr, 1) <> " " Then
    If CapsOn = True Then newstr = AddBeg & UCase(Left(TextStr, 1)) & AddEnd
    If CapsOn = False Then newstr = AddBeg & Left(TextStr, 1) & AddEnd
    LastSpace = False
End If

Dim x As Integer, CharStr As String

For x = 2 To Len(TextStr)
        CharStr = Mid(TextStr, x, 1)
        
        If CapsOn = True And LastSpace = True Then CharStr = UCase(CharStr)
        
        If CharStr <> " " And LastSpace = True Then
            newstr = newstr & AddBeg & CharStr & AddEnd
            LastSpace = False
        Else
            newstr = newstr & CharStr
        End If
        
        If CharStr = " " Then LastSpace = True
Next x

DecorateText = newstr

End Function


Sub Form_Slide(Frm As Form)
'*new sub
'this is really cool, watch your form slide across the screen



Dim grow As Single
Dim x As Single, y As Single
Dim OW As Single, OH As Single
Dim G1 As Single, G2 As Single
Dim rw As Single, rh As Single




OW = Frm.Width
OH = Frm.Height
rw = Frm.Width
rh = Frm.Height


x = Frm.Left + Frm.Width / 2
y = Frm.Top + Frm.Height / 2

Randomize Timer
grow = Rnd * 100
grow = grow + 300
G1 = grow

Dim d As Integer

For d = 1 To grow
    Frm.Width = OW + d
    Form_CenterAt Frm, x + d, y
    DoEvents
Next d


OW = Frm.Width
OH = Frm.Height
x = Frm.Left + Frm.Width / 2
y = Frm.Top + Frm.Height / 2



Randomize Timer
grow = Rnd * 100
grow = grow + 300
G2 = grow
 
For d = 1 To grow
    Frm.Height = OH + d
    Form_CenterAt Frm, x, y + d
    DoEvents
Next d

Dim s1 As Boolean

Do While G1 > 1
    If s1 Then
    Frm.Left = Frm.Left + 1
    Else
    Frm.Width = Frm.Width - 2
    End If
    s1 = Not s1
    G1 = G1 - 1
Loop

Do While G2 > 1
    If s1 Then
    Frm.Top = Frm.Top + 1
    Else
    Frm.Height = Frm.Height - 2
    End If
    s1 = Not s1
    G2 = G2 - 1
Loop


End Sub

Sub Form_SmartSize(Frm As Form)
'new in utils³
'this sub rocks...it will resize the form to fit the controls, a novel idea..

Dim x As Control
Dim maxleft As Single, maxtop As Single, Maxright As Single, Maxbottom As Single


For Each x In Frm.Controls
        If x.Left + x.Width > Maxright Then Maxright = x.Left + x.Width
    If x.Top + x.Height > Maxbottom Then Maxbottom = x.Top + x.Height
Next x

Frm.Width = Maxright + 50 + (Frm.Width - Frm.ScaleWidth)
Frm.Height = Maxbottom + 50 + (Frm.Height - Frm.ScaleHeight)

End Sub

Sub Form_SpellCaption(Frm As Form, CaptionStr As String)
'*new sub!
Dim x As Integer

For x = 1 To Len(CaptionStr)
    Frm.Caption = Right(CaptionStr, x)
    TimeOutX 0.001
Next x

End Sub


Function GetOption(Frm As Form) As String
'*new sub!
'cool as hell
'will return the caption of the option button currently selected. i think this is wildly useful
'note wont work if the options are in a control like a frame or picturebox, nothing i can do. :(

On Error Resume Next

Dim x As OptionButton
For Each x In Frm.Controls
    If x.Value = True Then
        GetOption = x.Caption: Exit Function
    End If
Next x
GetOption = "No Option Selected Or None Exist"
End Function

Function GetChecks(Frm As Form) As String
'*new sub!
'very cool
'will return a string with all the active checkboxes
'note wont work if the options are in a control like a frame or picturebox, nothing i can do. :(
On Error Resume Next

Dim sRet As String


Dim x As CheckBox
For Each x In Frm.Controls
    If x.Value = 1 Then
        sRet = sRet & ", " & x.Caption
    End If
Next x

If sRet = "" Then
GetChecks = "No Checkboxes Selected Or None Exist"
Else
GetChecks = Mid(sRet, 3)
End If
End Function


Sub HideControls(Frm As Form, Optional DoNotHide As Control)
'*new sub!


'fairly obvious
'if you give a control as DoNotHide then it, get this, will stay visible

On Error Resume Next
Dim x As Control
For Each x In Frm.Controls
    If x.Name <> DoNotHide.Name Then x.Visible = False
Next x

End Sub


Function IsUppercase(str As String) As Boolean
If UCase(str) = str Then
    IsUppercase = True
Else
    IsUppercase = False
End If
End Function

Sub Label_TempCap(lab As Label, TempCap As String)
'new sub!
'sets the labels text to something, then sets it back to the origianl after 5 secs.
Dim OC As String
OC = lab.Caption
lab.Caption = TempCap
TimeOutX (5)
lab.Caption = OC
End Sub

Function LineCount(CountStr As String) As Integer
'new sub
'hehe. this is better and faster than any other way! damn im proud of figuring this one out!
LineCount = Len(CountStr) - Len(Replace(CountStr, Chr(13), ""))
End Function

Sub List_AddFiles(lst As ListBox, Path As String, Optional ExtFilter As String)
'new sub!

'add all the files of the given type in the dir to a list
'ex : Call List_AddFiles(List1, "C:\windows", "*.dll")
If ExtFilter = "" Then ExtFilter = "*.*"

If Right(Path, 1) <> "\" Then Path = Path & "\"

Dim FindFile As String

FindFile = DiR(Path & ExtFilter)

Do Until FindFile = ""
    lst.AddItem FindFile
    FindFile = DiR
Loop

End Sub



Function MySite()
'*new sub! :P
OpenWebsite "http://www.george.cx"
End Function

Sub OpenWebsite(URL As String)
'*new sub!
'opens a website in the default browser!

On Error Resume Next
Call Shell("c:\windows\command\start.exe " & URL)
End Sub

Sub QueryExit()
'*new sub!
'asks the user if he/se wants to quit, if so then it quits

Dim e As VbMsgBoxResult
e = MsgBox("Are you sure that you want to exit?", vbYesNo, App.Title)
If e = vbYes Then Form_UnloadAll: End
End Sub

Function RandomNumber(Minimum As Integer, Maximum As Integer)
'*new sub!
'ez but convienient
Randomize Timer
Dim rn As String
rn = Rnd * (Maximum - Minimum)
RandomNumber = rn + Minimum

End Function

Function RandomSite() As String
'*new sub!

'This will allow someone to visit one of the website i like a  lot
'this is off the top of my head! not near a complete list!
'if they wanted to do that. it opens the site and returns the url
Dim sitearray() As String
Dim rn As String
Dim ws As String
Dim sites1 As String

sites1 = "www.george.cx|www.wheresgeorge.com|www.cybertears.net|www.redrival.com/wgf|www.mrdoobie.com|www.hecklers.com|unity2k.cjb.net|www.humor.com|www.cruel.com|www.epitaph.com|www.dj-tiesto.com|www.djspooky.com|www.okayplayer.com|www.brain-damage.net|www.dajoker.net|www.abbot3000.com|www.vertek.net|www.djshadow.com|www.grooveinjun.com"
sitearray = Split(sites1, "|")

Randomize Timer
rn = Rnd * UBound(sitearray)
ws = sitearray(rn)

OpenWebsite ws
RandomSite = ws

End Function

Function RandomSysFont()
'new sub!

Randomize Timer

Dim rn As Integer

rn = Round((Rnd * Screen.FontCount) - 1, 0)
RandomSysFont = Screen.Fonts(rn)

End Function

Sub ShowControls(Frm As Form, Optional DoNotShow As Control)
'*new sub!
'fairly obvious
'if you give a control as DoNotHide then it, get this, will stay visible

On Error Resume Next
Dim x As Control
For Each x In Frm.Controls
    If x.Name <> DoNotShow.Name Then x.Visible = True
Next x

End Sub

Function NormalizeString(OldStr As String) As String
'*new sub!

'cuts out commas, new lines, nulls, and lowercases it. i found myself doing this alot. :)
Dim str1 As String
str1 = LCase(OldStr)
str1 = Replace(str1, ",", "")
str1 = Replace(str1, Chr(0), "")
str1 = Replace(str1, Chr(13) & Chr(10), "")
NormalizeString = str1
End Function

Function ReadBoolStr(Boolstr As String) As Boolean
'*new sub!

' i got tired of coding for common  synonyms for true and false, so this interprets it for you :)

If Boolstr = "on" Then ReadBoolStr = True
If Boolstr = "off" Then ReadBoolStr = False

If Boolstr = "yes" Then ReadBoolStr = True
If Boolstr = "no" Then ReadBoolStr = False

If Boolstr = "1" Then ReadBoolStr = True
If Boolstr = "0" Then ReadBoolStr = False

If Boolstr = "true" Then ReadBoolStr = True
If Boolstr = "false" Then ReadBoolStr = False

If Boolstr = "enable" Then ReadBoolStr = True
If Boolstr = "disable" Then ReadBoolStr = False

If Boolstr = "enabled" Then ReadBoolStr = True
If Boolstr = "disabled" Then ReadBoolStr = False

End Function


Sub Form_QuickBeat(Frm As Form, Optional NuMPuLSeS As Integer, Optional BeatSpeed As Integer)
If NuMPuLSeS = 0 Then NuMPuLSeS = 5
If BeatSpeed = 0 Then BeatSpeed = 7
If BeatSpeed > 10 Then BeatSpeed = 10
On Error Resume Next
Dim y As Integer

y = 255
Do While NuMPuLSeS > 0
Do While y < 255
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = y + 15
Loop
Do While y > 100
Frm.BackColor = RGB(y - 10, 0, 0)
TimeOutX (0.00000000001)
y = y - 15
Loop
NuMPuLSeS = NuMPuLSeS - 1
Loop

End Sub


Sub Form_HeartBeat(Frm As Form, Optional NuMPuLSeS As Integer, Optional BeatSpeed As Integer)
If NuMPuLSeS = 0 Then NuMPuLSeS = 5
If BeatSpeed = 0 Then BeatSpeed = 7
If BeatSpeed > 10 Then BeatSpeed = 10
On Error Resume Next
Dim y As Integer

y = 255
Do While NuMPuLSeS > 0
y = 235
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = 240
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = 245
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = 255
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)

y = 255
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = 250
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = 245
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
y = 240
Frm.BackColor = RGB(y, 0, 0)
TimeOutX (1E-18)
Frm.BackColor = RGB(0, 0, 0)

NuMPuLSeS = NuMPuLSeS - 1
TimeOutX (BeatSpeed)
Loop

End Sub

Sub Form_ExitÐàMNéÐ(Frm As Form)
' This sub is for ÐàMNéÐ cuz he asked
' me to do some exit subs for him
' but anyone can use it
Form_HideAllControls Frm
Form_KrazeFade Frm
TimeOutX (1)
Form_RedWarp Frm, 1
Form_QuickBeat Frm, 1
Unload Frm
End Sub

Sub Form_SizeRight(Frm As Form)
On Error Resume Next
Dim maxx As Long
Dim maxy As Long

Dim x As Control

For Each x In Frm.Controls
If x = 1 Then Frm.ScaleMode = x.ScaleMode
    If maxx < x.Left + x.Width Then maxx = x.Left + x.Width
    If maxy < x.Top + x.Height Then maxy = x.Left + x.Width
Next x
Frm.Left = maxx + 100
Frm.Width = maxy + 100
End Sub

Sub Form_WarpGreen(Frm As Form)
Call Form_Warp(Frm, &HFF&, &HFFFF&)
End Sub

Sub Form_Warp(Frm As Form, StartColor, EndColor)
    ' modified form monkegod's fadeform sub
    
    Dim sb As Integer
    Dim sg As Integer
    Dim sr As Integer
    Dim eb As Integer
    Dim eg As Integer
    Dim er As Integer
    Dim OldRGB As Long
    On Error Resume Next
    sb = zGetRGB(StartColor).Blue
    sg = zGetRGB(StartColor).Green
    sr = zGetRGB(StartColor).Red
    eb = zGetRGB(EndColor).Blue
    eg = zGetRGB(EndColor).Green
    er = zGetRGB(EndColor).Red
    Dim x, MooK As Integer
    For x = 1 To 255
        If OldRGB <> RGB((er - sr) / 255 * x, (eg - sg) / 50 * x, (eb - sb) / 50 * x) Then
        Frm.BackColor = RGB((er - sr) / 255 * x, (eg - sg) / 50 * x, (eb - sb) / 50 * x)
        OldRGB = RGB((er - sr) / 255 * x, (eg - sg) / 50 * x, (eb - sb) / 50 * x)
        Else
        Exit For
        End If
        TimeOutX (0.01)
    Next x
End Sub



Sub Form_ExitChillin(Frm As Form)
Form_HideAllControls Frm
Form_Flash Frm, 7
Form_FadeToBlack Frm, False
TimeOutX (1)
Frm.BackColor = &HFFFFFF
TimeOutX (0.1)
Unload Frm
End Sub

Sub Form_Spiral(Frm As Form)
'edited because it was wack.....
Dim x As Integer

For x = 1 To 5
Form_TopLeft Frm
TimeOutX (0.1)
Form_TopRight Frm
TimeOutX (0.1)
Form_BottomRight Frm
TimeOutX (0.1)
Form_BottomLeft Frm
TimeOutX (0.1)
Next x

Form_Center Frm
For x = 1 To 5
Frm.Visible = False
TimeOutX 0.1
Frm.Visible = True
Next x


End Sub

Sub Form_TopLeft(Frm As Form)
Frm.Top = 0
Frm.Left = 0
End Sub
Sub Form_TopCenter(Frm As Form)
Form_Center Frm
Frm.Top = 0
End Sub

Sub Form_TopRight(Frm As Form)
Frm.Top = 0
Frm.Left = Screen.Width - Frm.Width
End Sub
Sub Form_BottomLeft(Frm As Form)
Frm.Top = Screen.Height - Frm.Height
Frm.Left = 0
End Sub
Sub Form_BottomRight(Frm As Form)
Frm.Top = Screen.Height - Frm.Height
Frm.Left = Screen.Width - Frm.Width
End Sub

Sub Form_HideAllControls(Frm As Form)
On Error Resume Next
Dim Crl As Control
For Each Crl In Frm.Controls
    Crl.Visible = False
Next Crl
End Sub



Sub Form_ShowAllControls(Frm As Form)
On Error Resume Next
Dim Crl As Control
For Each Crl In Frm.Controls
    Crl.Visible = True
Next Crl
End Sub



Sub Form_WildEntry(Frm As Form)
Frm.Top = 0
Frm.Left = 0
TimeOutX (0.2)
Frm.Left = Screen.Width - Frm.Width
TimeOutX (0.2)
Frm.Top = Screen.Height - Frm.Height
TimeOutX (0.2)
Frm.Left = 0
TimeOutX (0.2)
Form_Center Frm
Form_Flash Frm, 4

End Sub

Sub Help_CheckBoxesInListBoxes()
' If you want your listbox to have a
' check box, set the style property to 1

End Sub

Sub Hide_Lists(Frm As Form)
On Error Resume Next
Dim x
For Each x In Frm.Controls
    x.Visible = False
Next x
End Sub

Sub Label_Center(lab As Label)
lab.Top = ((lab.Parent.ScaleHeight / 2) - (lab.Height / 2))
lab.Left = ((lab.Parent.ScaleWidth / 2) - (lab.Width / 2))
End Sub

Sub Label_FX(lab As Label, message As String)
Label_Center lab
Label_Spell message, lab
Label_RedWarp lab
End Sub

Sub Label_Grow(lbl As Label)
Do While lbl.FontSize < 48
lbl.FontSize = lbl.FontSize + 2
TimeOutX (0.01)
Loop
End Sub
Sub Label_Shrink(lbl As Label)
Do While lbl.FontSize > 6
lbl.FontSize = lbl.FontSize - 2
TimeOutX (0.01)
Loop
End Sub

Sub Label_Spell(Word As String, lab As Label)
Dim x As Integer

For x = 1 To Len(Word)
lab.Caption = Left$(Word, x)
TimeOutX 0.1
Next x
End Sub


Sub List_AddFontSizes(lst)
Dim x As Integer

For x = 6 To 32 Step 2
lst.AddItem x
Next x
lst.AddItem "48"

lst.AddItem "36"
lst.AddItem "72"
End Sub

Sub List_HScrollBar(lst)
'yes! give a a horizontal scrollbar! cool!
Dim nRet As Long
Dim nNewWidth As Integer
nNewWidth = lst.Width + 1 'new width in pixels
nRet = SendMessage(lst.hwnd, LB_SETHORIZONTALEXTENT, nNewWidth, ByVal 0&)
End Sub

Function MultOf(StartAt As Long, MultOfInt As Integer, numdown As Integer) As Integer
MultOf = StartAt + (MultOfInt * numdown)
End Function

Function IsMultOf(SmallNumber As Integer, BigNumber As Integer) As Boolean
'new sub!
'returns a boolean value for if SmallNumber is a multiole of BigNumber
If DivisionRemainder(BigNumber, SmallNumber) > 0 Then
    IsMultOf = False
Else
    IsMultOf = True
End If
End Function


Sub PicBox_ScrollList(pb As PictureBox, lst As ListBox)
'edited for utils³
'umm, could be usefull ?
pb.Cls
Dim C As Integer
C = 0

Dim x As Integer

For x = 0 To lst.ListCount - 1
pb.Cls
pb.Print lst.List(x)
TimeOutX (1)
C = C + 1
Next x
End Sub
Function RandomFont() As String
'i know this leaves out fonts but its fast and old
'use RandomSysFont for the real thing

Dim FontArray(98)
FontArray(1) = "Abadi MT Condensed"
FontArray(2) = "Abadi MT Condensed Light"
FontArray(3) = "Arial"
FontArray(4) = "Arial Black"
FontArray(5) = "Arial Narrow"
FontArray(6) = "Arial Rounded MT Bold"
FontArray(7) = "Beesknees ITC"
FontArray(8) = "Book Antiqua"
FontArray(9) = "Bookman Old Style"
FontArray(10) = "Bradley Hand ITC"
FontArray(11) = "Brush Script MT"
FontArray(12) = "Calisto MT"
FontArray(13) = "Century Gothic"
FontArray(14) = "Century Schoolbook"
FontArray(15) = "Comic Sans MS"
FontArray(16) = "Copperplate Gothic Bold"
FontArray(17) = "Copperplate Gothic Light"
FontArray(18) = "Courier"
FontArray(19) = "Courier New"
FontArray(20) = "Curlz MT"
FontArray(22) = "Elephant"
FontArray(23) = "Engravers MT"
FontArray(24) = "Eras Bold ITC"
FontArray(25) = "Eras Demi ITC"
FontArray(26) = "Eras Light ITC"
FontArray(27) = "Eras Medium ITC"
FontArray(28) = "Eras Ultra ITC"
FontArray(29) = "Felix Titling"
FontArray(30) = "Fixedsys"
FontArray(31) = "Forte"
FontArray(32) = "Franklin Gothic Book"
FontArray(33) = "Franklin Gothic Demi"
FontArray(34) = "Franklin Gothic Demi Cond"
FontArray(35) = "Franklin Gothic Heavy"
FontArray(36) = "Franklin Gothic Medium"
FontArray(37) = "Franklin Gothic Medium Cond"
FontArray(38) = "French Script MT"
FontArray(39) = "Garamond"
FontArray(40) = "Georgia"
FontArray(41) = "Gill Sans MT"
FontArray(42) = "Gill Sans MT Condensed"
FontArray(43) = "Gill Sans MT Ext Condensed Bold"
FontArray(44) = "Gill Sans Ultra Bold"
FontArray(45) = "Gill Sans Ultra Bold Condensed"
FontArray(46) = "Gloucester MT Extra Condensed"
FontArray(47) = "Goudy Old Style"
FontArray(48) = "Haettenschweiler"
FontArray(50) = "Impact"
FontArray(51) = "Imprint MT Shadow"
FontArray(52) = "Juice ITC"
FontArray(54) = "Lucida Console"
FontArray(55) = "Lucida Handwriting"
FontArray(56) = "Lucida Sans"
FontArray(57) = "Lucida Sans Typewriter"
FontArray(58) = "Lucida Sans Unicode"
FontArray(59) = "Maiandra GD"
FontArray(60) = "Marlett"
FontArray(61) = "Matisse ITC"
FontArray(62) = "Modern"
FontArray(63) = "Monotype Sorts"
FontArray(65) = "MS LineDraw"
FontArray(66) = "MS Sans Serif"
FontArray(67) = "MS Serif"
FontArray(68) = "News Gothic MT"
FontArray(69) = "OCR A Extended"
FontArray(70) = "Palace Script MT"
FontArray(71) = "Perpetua"
FontArray(72) = "Perpetua Titling MT"
FontArray(73) = "PrestigeFixed"
FontArray(74) = "Rage Italic"
FontArray(75) = "Rockwell"
FontArray(76) = "Rockwell Condensed"
FontArray(77) = "Rockwell Extra Bold"
FontArray(78) = "Script MT Bold"
FontArray(79) = "Small Fonts"
FontArray(80) = "Snap ITC"
FontArray(81) = "Symbol"
FontArray(82) = "Symbol"
FontArray(83) = "Tahoma"
FontArray(84) = "Tempus Sans ITC"
FontArray(85) = "Terminal"
FontArray(86) = "Times New Roman"
FontArray(87) = "Trebuchet MS"
FontArray(88) = "Tw Cen MT"
FontArray(89) = "Tw Cen MT Condensed"
FontArray(90) = "Tw Cen MT Condensed Extra Bold"
FontArray(91) = "Verdana"
FontArray(92) = "Viner Hand ITC"
FontArray(93) = "Webdings"
FontArray(94) = "Westminster"
FontArray(95) = "Wide Latin"
FontArray(96) = "Wingdings"
FontArray(97) = "Wingdings 2"
FontArray(98) = "Wingdings 3"
Dim v As Integer
v = Int((98 * Rnd) + 1)
RandomFont = FontArray(v)
If RandomFont = "" Then RandomFont = "Arial"
End Function

Function RichTextToHTML(rtb)

'------------------------------------------------------
'  * Not all proprties of the text are converted to
'  * HTML code. Look for that in a later version
'------------------------------------------------------



'------------------------------------------------------
'Variable Declarations
'------------------------------------------------------
Dim FinalString, CharString As String
Dim CurrentChar, Font, FontString, ColorString, PropString As String
Dim Go As Integer
Dim LastCharColor As Long
Dim LastCharFont As String
Dim LastCharBold, LastCharItalic, LastCharUnderline, LastCharStrike As Boolean

LastCharBold = False
LastCharItalic = False
LastCharUnderline = False
LastCharStrike = False

For Go = 0 To Len(rtb.Text)
PropString = ""
FontString = ""
ColorString = ""

rtb.SelStart = Go
rtb.SelLength = 1
CurrentChar = rtb.SelText

Dim nl As String

nl = Chr(13) + Chr(10)
If CurrentChar = nl Then GoTo SkipIt
If CurrentChar = " " Then GoTo JustFont


'------------------------------------------------------
'Test for bold, italic and underline properties of text
'And create a string to set these
'------------------------------------------------------


If rtb.SelBold = True And LastCharBold = False Then
PropString = PropString & "<B>"
LastCharBold = True
End If
If rtb.SelBold = False And LastCharBold = True Then
PropString = PropString & "</B>"
LastCharBold = False
End If


If rtb.SelItalic = True And LastCharItalic = False Then
PropString = PropString & "<I>"
LastCharItalic = True
End If
If rtb.SelItalic = False And LastCharItalic = True Then
PropString = PropString & "</I>"
LastCharItalic = False
End If


If rtb.SelUnderline = True And LastCharUnderline = False Then
PropString = PropString & "<U>"
LastCharUnderline = True
End If
If rtb.SelUnderline = False And LastCharUnderline = True Then
PropString = PropString & "</U>"
LastCharUnderline = False
End If


If rtb.SelStrikeThru = True And LastCharStrike = False Then
PropString = PropString & "<strike>"
LastCharStrike = True
End If
If rtb.SelUnderline = False And LastCharStrike = True Then
PropString = PropString & "</strike>"
LastCharStrike = False
End If

'------------------------------------------------------
'Form the string to set the color
'------------------------------------------------------
If LastCharColor <> rtb.SelColor Then
ColorString = "<Font color=#" & rtb.SelColor & ">"
LastCharColor = rtb.SelColor
End If

JustFont:
'------------------------------------------------------
'Form the string to change the font
'------------------------------------------------------

If LastCharFont <> rtb.SelFontName Then
FontString = "<font face=""" & rtb.SelFontName & """>"
LastCharFont = rtb.SelFontName
End If


'------------------------------------------------------
'Form the final HTML string for this charachter
'And add it to total string
'------------------------------------------------------
CharString = PropString & ColorString & FontString & CurrentChar
FinalString = FinalString & CharString

'------------------------------------------------------
'Repeat for the next charachter!
'------------------------------------------------------
SkipIt:

Next Go
RichTextToHTML = FinalString
End Function
Sub Show_Lists(Frm As Form)
On Error Resume Next
Dim x
For Each x In Frm.Controls
    x.Visible = True
Next x
End Sub
Sub Show_ComboBoxes(Frm As Form)
On Error Resume Next
Dim x As ComboBox
For Each x In Frm.Controls
    x.Visible = True
Next x
End Sub
Sub Hide_ComboBoxes(Frm As Form)
On Error Resume Next
Dim x As ComboBox
For Each x In Frm.Controls
    x.Visible = False
Next x
End Sub

Sub Hide_CheckBoxes(Frm As Form)
On Error Resume Next
Dim x As CheckBox
For Each x In Frm.Controls
    x.Visible = False
Next x
End Sub

Sub Hide_OptionButtons(Frm As Form)
On Error Resume Next
Dim x As OptionButton
For Each x In Frm.Controls
    x.Visible = False
Next x
End Sub


Sub Show_OptionButtons(Frm As Form)
On Error Resume Next
Dim x As OptionButton
For Each x In Frm.Controls
    x.Visible = True
Next x
End Sub



Sub Show_Forms(Frm As Form)
' *** This will only
' *** show the loaded forms

On Error Resume Next
Dim x As Form
For Each x In Frm.Controls
    x.Visible = True
Next x
End Sub
Sub Hide_Forms(Frm As Form)
On Error Resume Next
Dim x As Form
For Each x In Frm.Controls
    x.Visible = False
Next x
End Sub




Sub Show_CheckBoxes(Frm As Form)
On Error Resume Next
Dim x As CheckBox
For Each x In Frm.Controls
    x.Visible = True
Next x
End Sub


Function IsEven(str As Integer) As Boolean
Dim y As Integer
y = str / 2
If InStr(1, y, ".") Then
IsEven = False
Else
IsEven = True
End If
End Function

Sub List_Save(DiR, lst)
Dim SaveList As Long
On Error Resume Next
Open DiR For Output As #1
For SaveList& = 0 To lst.ListCount - 1
Print #1, lst.List(SaveList&)
Next SaveList&
Close #1
End Sub
Function List_IsInList(SearchFer As String, lst) As Boolean
On Error Resume Next
If LCase(lst.List(0)) = LCase(SearchFer) Then
    List_IsInList = True
    Exit Function
End If
Dim x As Integer

For x = 0 To lst.ListCount
    
    If LCase(lst.List(x)) = LCase(SearchFer) Then
    List_IsInList = True
    Exit Function
    End If
    
Next x
List_IsInList = False

End Function


Sub List_Load(DiR, lst)
Dim strin As String
On Error Resume Next
Open DiR For Input As #1
While Not EOF(1)
Input #1, strin$
DoEvents
lst.AddItem strin$
Wend
Close #1
Exit Sub
End Sub

Function Filter(Text$, charfind$, charchange$)
'i left this here in case you dont have vb6
'btw, get it...keep up with the vers, its worth it

Dim ReplaceText As String, TheChar As String, TheChars As String

If InStr(Text$, charfind$) = 0 Then
ReplaceText = Text$
Exit Function
End If

Dim ReplaceThis As Integer

For ReplaceThis = 1 To Len(Text$)
TheChar$ = Mid(Text$, ReplaceThis, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = charfind$ Then
TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1) + charchange$
End If
Next ReplaceThis
Filter = TheChars$
End Function


Sub Midi_Play(file$)
Dim z
z = mciExecute("play " & file$)
End Sub





Sub Form_FadeToBlack(Frm As Form, Optional ShouldExit As Boolean)

Frm.BackColor = 16777215
TimeOutX (0.02)
Frm.BackColor = 14737632
TimeOutX (0.02)
Frm.BackColor = 12632256
TimeOutX (0.02)
Frm.BackColor = 8421504
TimeOutX (0.02)
Frm.BackColor = 4210752
TimeOutX (0.02)
Frm.BackColor = 0
If ShouldExit = True Then Unload Frm
End Sub

Sub Form_Flash(Frm As Form, Optional NumFlashes As Integer)
On Error Resume Next
Dim x As Integer
Dim y As String
y = Frm.BackColor
If y = 0 Then y = &H0&

x = NumFlashes
If x = 0 Then x = 100

Do While x > 0
Frm.BackColor = &HFFFFFF
TimeOutX (0.01)
Frm.BackColor = &H0&
TimeOutX (0.01)
x = x - 1
Loop
Frm.BackColor = y
End Sub
Sub Form_GloryHoleFade(Frm As Form)
' This sub is an adaptaion from another
' sub. There was no name, so I can't
' give kredit

Frm.AutoRedraw = True
Frm.Cls
Dim cx, cy, i
    Frm.ScaleMode = 3
    cx = Frm.ScaleWidth \ 2
    cy = Frm.ScaleHeight \ 2

Frm.DrawWidth = 2
For i = 0 To 255
Frm.Circle (cx, cy), i, RGB(255 - i, 0, 0)  'Red to Black
Next i


End Sub

Sub Form_ExitTopLeft(Frm As Form)

Dim OldH As Integer
On Error Resume Next
Do While Frm.Top + Frm.Height > 0
With Frm
    .Top = .Top - 50
    .Left = .Left - 50
End With
DoEvents
Loop
Unload Frm
End Sub
Sub Form_ExitTopRight(Frm As Form)

Dim OldH As Integer
On Error Resume Next
Do While Frm.Top + Frm.Height > 0
With Frm
    .Top = .Top - 50
    .Left = .Left + 50
End With
DoEvents
Loop
Unload Frm
End Sub

Sub Form_ExitBottomLeft(Frm As Form)


Dim OldH As Integer
On Error Resume Next
Do While Frm.Top < Screen.Height
With Frm
    .Top = .Top + 50
    .Left = .Left - 50
End With
DoEvents
Loop

Unload Frm
End Sub

Sub Form_ExitBottomRight(Frm As Form)

Dim OldH As Integer
On Error Resume Next
Do While Frm.Top < Screen.Height
With Frm
    .Top = .Top + 50
    .Left = .Left + 50
End With
DoEvents
Loop

Unload Frm
End Sub


Sub Form_ExitShrinkTopLeft(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Height = .Height - 50
    .Width = .Width - 50
    End With
TimeOutX (0.01)
Loop
Unload Frm
End Sub
Sub Form_ExitShrinkTopRight(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Height = .Height - 50
    .Width = .Width - 50
    .Left = .Left + 50
End With

Loop
Unload Frm
End Sub

Sub Form_ExitShrinkBottomRight(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Top = .Top + 50
    .Height = .Height - 50
    .Width = .Width - 50
    .Left = .Left + 50
End With
TimeOutX (0.01)
Loop
Unload Frm
End Sub
Sub Form_ExitShrinkBottomLeft(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Top = .Top + 50
    .Height = .Height - 50
    .Width = .Width - 50
    
End With
TimeOutX (0.01)
Loop
Unload Frm
End Sub




Sub Form_Implode(Frm As Form)
Dim OldH As Integer
On Error Resume Next
Do While OldH <> Frm.Height
OldH = Frm.Height
With Frm
    .Top = .Top + 40
    .Height = .Height - 80
    .Width = .Width - 80
    .Left = .Left + 40
End With
TimeOutX (0.001)

Loop
Unload Frm
End Sub

Sub Form_Explode(Frm As Form)
Form_Center Frm
Dim x, y As Integer
x = Frm.Height
y = Frm.Width

With Frm
    .Height = 50
    .Width = 50
End With
Dim OldH As Integer
On Error Resume Next
Do
    OldH = Frm.Height
    If Frm.Width < y Then
        Frm.Width = Frm.Width + 80
    End If
Frm.Top = Frm.Top - 40
    If Frm.Height < x Then
        Frm.Height = Frm.Height + 80
    End If
Frm.Left = Frm.Left - 40
Form_Center Frm
TimeOutX (0.001)

Loop Until x <= Frm.Height Or y <= Frm.Width
Frm.Height = x
Frm.Width = y

End Sub




Sub Form_KrazeFade(Frm)
' This sub is an adaptaion from another
' sub. There was no name, so I can't
' give kredit if you use it or optimize it
' give us credit (Naïve, Crystal Dawn)

Frm.AutoRedraw = True
Frm.Cls
Dim cx, cy, F, F1, F2, i
Frm.AutoRedraw = True
Frm.Cls
    Frm.ScaleMode = 3
    cx = Frm.ScaleWidth / 2
    cy = Frm.ScaleHeight / 2
Dim DrawWidth As Integer

DrawWidth = 2
For i = 255 To 0 Step -2
F = i / 255
F1 = 1 - F: F2 = 1 + F
Frm.ForeColor = RGB(i, 0, i)
Frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF
Next i
    Frm.ScaleMode = 3
    cx = Frm.ScaleWidth / 2
    cy = Frm.ScaleHeight / 2
    Frm.DrawWidth = 2

For i = 0 To 255
Frm.ForeColor = RGB(i, 0, i)
F = i / 255
F1 = 1 - F: F2 = 1 + F
        
Frm.Line (cx * F1, cy)-(cx, cy * F1)
Frm.Line -(cx * F2, cy)
Frm.Line -(cx, cy * F2)
Frm.Line -(cx * F1, cy)
Next i
End Sub


Sub Form_SuperWarp(Frm As Form, Optional NumTrips As Integer)
'changes the colors alot. kind of nice.

Dim oldbk As String
oldbk = Frm.BackColor

If NumTrips = 0 Then NumTrips = 5
Do While NumTrips > 0
Frm.BackColor = 16711680
TimeOutX (0.1)
Frm.BackColor = 12582912
TimeOutX (0.1)
Frm.BackColor = 8388608
TimeOutX (0.1)
Frm.BackColor = 8388736
TimeOutX (0.1)
Frm.BackColor = 12583104
TimeOutX (0.1)
Frm.BackColor = 12583104
TimeOutX (0.1)
Frm.BackColor = 12583104
TimeOutX (0.1)
Frm.BackColor = 16761087
TimeOutX (0.1)
Frm.BackColor = 16777215
TimeOutX (0.1)
Frm.BackColor = 14737632
TimeOutX (0.1)
Frm.BackColor = 12632256
TimeOutX (0.1)
Frm.BackColor = 8421504
TimeOutX (0.1)
Frm.BackColor = 4210752
TimeOutX (0.1)
Frm.BackColor = 128
TimeOutX (0.1)
Frm.BackColor = 192
TimeOutX (0.1)
Frm.BackColor = 255
TimeOutX (0.1)
Frm.BackColor = 8421631
TimeOutX (0.1)
Frm.BackColor = 12632319
TimeOutX (0.1)
Frm.BackColor = 12640511
TimeOutX (0.1)
Frm.BackColor = 8438015
TimeOutX (0.1)
Frm.BackColor = 3313
TimeOutX (0.1)
Frm.BackColor = 16576
NumTrips = NumTrips - 1
Loop
Frm.BackColor = oldbk
End Sub

Function Integer_Increment(IntX As String)
If TestIfInteger(IntX) = True Then
IntX = IntX + 1
Integer_Increment = IntX
Else
Integer_Increment = IntX
End If
End Function

Sub List_KillDupes2(lst)
' This is the old style
' if th other one gives ya probs
' use this, but its a lot slower
On Error GoTo Kraze
Dim OldInd As Integer, lucid As String

Dim x As Integer
Dim y As Integer

OldInd = lst.ListIndex

For x = 0 To lst.ListCount - 1
For y = 0 To lst.ListCount - 1
If y < lst.ListCount Then
lst.ListIndex = y
End If
lucid = lst.Text
If x < lst.ListCount Then
lst.ListIndex = x
Else
Exit Sub
End If
If y = x Then GoTo Kraze
If y < lst.ListCount Then
If lucid = lst.Text Then lst.RemoveItem (y)
End If
Kraze:
Next y
Next x
If OldInd > lst.ListCount Then
lst.ListIndex = OldInd
End If
End Sub
Sub List_KillDupes(lst)
On Error Resume Next
Dim x As Integer
Dim lucid As String
Dim y As Integer
Dim Kraze As String

For x = 1 To lst.ListCount + 2
    lucid = lst.List(x)
    For y = 1 To lst.ListCount + 2
    Kraze = lst.List(y)
    If Kraze = lucid Then
    If x > y Then
    lst.RemoveItem y
    ElseIf x < y Then
    lst.RemoveItem y
    End If
    End If
    Next y
    Next x
lst.ListIndex = lst.ListCount - 1
If lst.ListCount > 0 Then
Do While lst.List(0) = lst.List(lst.ListCount - 1)
lst.RemoveItem lst.ListCount - 1
Loop
End If
End Sub


Function Form_GetActiveControl() As String
' Put this in incase ya don't know

' returns true if the control that called it is active
' else it returns the nam of the active control
Form_GetActiveControl = Screen.ActiveControl
End Function
Sub Form_MoveForm(Frm As Form)
ReleaseCapture
Dim ret As Long


ret& = SendMessage(Frm.hwnd, &HA1, 2, 0)
End Sub



Function List_ListToString(lsd)
Dim XXX As String
Dim nl As String
nl = Chr(13) + Chr(10)
XXX = ""
Dim x As Integer
For x = 1 To lsd.ListCount
lsd.ListIndex = 0
XXX = XXX & lsd.Text & nl
Next x
List_ListToString = XXX
End Function

Function List_ListToStringNumbered(lsd)
Dim XXX As String
Dim nl As String
nl = Chr(13) + Chr(10)
XXX = ""
Dim x As Integer
For x = 1 To lsd.ListCount
lsd.ListIndex = 0
XXX = XXX & lsd.ListIndex + 1 & ") " & lsd.Text & nl
Next x
List_ListToStringNumbered = XXX
End Function


Sub StatusBar_SetText(StatBar, PanelNumber, NewCaption)
    StatBar.Panels(PanelNumber).Text = NewCaption
End Sub


Sub Fun_99Bottles()
'added the cancel option :), thank me someday. :P
Dim x As Integer
Dim e
x = 99
Do While x > 0
e = MsgBox(x & " bottles of beer on the wall, " & x & " bottles of beer. Take on down, pass it around, " & x - 1 & " bottles of beer on the wall.", vbOKCancel, "So much beer!")
If e = vbCancel Then GoTo Rehab
x = x - 1
Loop
Exit Sub

Rehab:
Call MsgBox("You gave up with " & x & " bottles of beer on the wall. Thats sad man.", , "You're an Assclown      >:0")

End Sub

Sub Fun_Kids()
' never ending. kinda fun
Do
Call MsgBox("It's " & Time & " do you know where your  children are?", vbQuestion, "Warning:")
TimeOutX (1)
Loop
End Sub


Function GetChar(str, Letter)
GetChar = Mid$(str, Letter, 1)
End Function

Function String_Increment(str)
On Error Resume Next
If IsNumeric(str) = True Then
    String_Increment = str + 1
    Exit Function
End If

Dim TestChr As String, newstr As String, x As Integer

For x = 1 To Len(str)
    TestChr = Mid(str, x, 1)
    If IsNumeric(TestChr) = True Then
        newstr = newstr & CStr(CInt(TestChr) + 1)
    Else
        newstr = newstr & TestChr
    End If
Next x

End Function
Function String_Reverse(str As String)
'i know there is a vb6 function to do this now, but this bas was started before vb6,
'there are some w/o vb6 ;)

Dim x, z As Integer
Dim y As String
For x = 0 To Len(str) - 1
    z = Len(str) - x
    y = y & Mid$(str, z, 1)
Next x
String_Reverse = y
End Function

Function String_SpaceWithChar(str As String, TheChar)
' This inserts th char between every charachter
Dim x, y As Integer
Dim TempString As String
y = 1
For x = 1 To Len(str)
    TempString = TempString & Mid$(str, y, 1) & TheChar
    y = y + 1
Next x
String_SpaceWithChar = TempString
End Function

Function Text_Reverse(str As String)
On Error GoTo Error
Dim words As Integer
Dim rt As String

For words = Len(str) To 1 Step -1
rt = rt & Mid(str, words, 1)
Next words
Text_Reverse = rt
Exit Function
Error:
Err = 1
End Function

Function TextBox_Filter(txt As TextBox, StringToFind As String, StringToReplace As String)
txt.Text = Filter(txt.Text, StringToFind, StringToReplace)
End Function


Sub List_Filter(LstX, StringToFind As String)
' This will remove any item without the
' String you pass to it

On Error Resume Next
LstX.ListIndex = 0

Dim x As Integer

For x = 0 To LstX.ListCount - 1
If LstX.List(x) <> StringToFind Then
    LstX.RemoveItem (x)
End If
Next x
End Sub
Sub List_RemoveByFilter(LstX, StringToFind As String)
' This will remove any item with the
' String you pass to it

On Error Resume Next
LstX.ListIndex = 0
Dim x As Integer
For x = 1 To LstX.ListCount
If InStr(1, StringToFind, LstX.Text) = 0 Then LstX.RemoveItem LstX.ListIndex
LstX.ListIndex = LstX.ListIndex + 1
Next x
End Sub




Sub Form_UnloadAll()
Dim Frm As Form
For Each Frm In Forms
Unload Frm
Next Frm
End Sub

Sub Form_HideAll()
Dim Frm As Form
For Each Frm In Forms
Frm.Visible = False
Next Frm
End Sub

Sub Form_ShowAll()
Dim Frm As Form
For Each Frm In Forms
Frm.Visible = True
Next Frm

End Sub


Function HTML_Bold(str As String)
HTML_Bold = "<B>" & str & "</B>"
End Function
Function HTML_Italic(str As String)
HTML_Italic = "<I>" & str & "</I>"
End Function
Function HTML_Underline(str As String)
HTML_Underline = "<U>" & str & "</U>"
End Function

Sub List_AsciiChars(lsd)
Dim x As Integer
For x = 33 To 223
lsd.AddItem Chr(x)
Next x
End Sub

Function TestIfInteger(Chk) As Boolean
' Tells you if a string is an integer
' I know this is like the vb function
' IsNumeric, but i wanted to find
' how it was done

Dim y As Integer
On Error GoTo Nope
y = Chk / 2
TestIfInteger = True
Exit Function
Nope:
TestIfInteger = False

End Function
Function HTML_Blue(str As String)
HTML_Blue = "<FONT COLOR=""#0000FF"">" & str
End Function

Function HTML_Red(str As String)
HTML_Red = "<FONT COLOR=""#FF0000"">" & str
End Function
Function HTML_Green(str As String)
HTML_Green = "<FONT COLOR=""#008000"">" & str
End Function


Function HTML_Black(str As String)
HTML_Black = "<FONT COLOR=""#FFFFFF"">" & str
End Function

Function HTML_Yellow(str As String)
HTML_Yellow = "<FONT COLOR=""&H0000FFFF&"">" & str
End Function


Function HTML_Purple(str As String)
HTML_Purple = "<FONT COLOR=""&H00C000C0&"">" & str
End Function



Function HTML_White(str As String)
HTML_White = "<FONT COLOR=""&H00FFFFFF&"">" & str
End Function




Sub TextBox_Bold(txt As TextBox)
If txt.Font.Bold = True Then txt.Font.Bold = False
If txt.Font.Bold = False Then txt.Font.Bold = True
End Sub


Sub TextBox_Spell(txt As TextBox, Word As String, Optional Speed As Integer)
Dim tIm
' Pick a Speed and pass it
'-------------------------
' 1 = very slow
' 2 = slow
' 3 = normal
' 4 = fast
' 5 = very fast
' 6 = Wild

If Speed = 0 Then tIm = 1
If Speed = 1 Then tIm = 2
If Speed = 2 Then tIm = 1
If Speed = 3 Then tIm = 0.5
If Speed = 4 Then tIm = 0.25
If Speed = 5 Then tIm = 0.1
If Speed = 6 Then tIm = 0.01


Dim ha As Integer, cnt As Integer
txt = ""
ha = Len(Word)
cnt = 1
Do While cnt <= ha
txt.Text = txt.Text & GetChar(Word, cnt)
TimeOutX tIm
cnt = cnt + 1
Loop
End Sub


Sub TextBox_Italic(txt As TextBox)
If txt.Font.Italic = True Then txt.Font.Italic = False
If txt.Font.Italic = False Then txt.Font.Italic = True
End Sub

Sub TextBox_Underline(txt As TextBox)
If txt.Font.Underline = True Then txt.Font.Underline = False
If txt.Font.Underline = False Then txt.Font.Underline = True
End Sub





Sub Form_ExitDown(ThaForm As Form)
Do
ThaForm.Top = Trim(str(Int(ThaForm.Top) + 300))
DoEvents
Loop Until ThaForm.Top > 7200
End Sub



Sub Form_ExitLeft(ThaForm As Form)
Do
ThaForm.Left = Trim(str(Int(ThaForm.Left) - 300))
DoEvents
Loop Until ThaForm.Left < -6300
End Sub



Sub Form_ExitRight(ThaForm As Form)
Do
ThaForm.Left = Trim(str(Int(ThaForm.Left) + 300))
DoEvents
Loop Until ThaForm.Left > 9600
End Sub



Sub Form_ExitUp(ThaForm As Form)
Do
ThaForm.Top = Trim(str(Int(ThaForm.Top) - 300))
DoEvents
Loop Until ThaForm.Top < -4500
End Sub


Sub Label_FlyInOut(lab As Label)
' This makes the Label's Text look
' Like it's flying in and out

lab.Visible = True
lab.FontSize = 1
TimeOutX (0.15)
lab.FontSize = 2
TimeOutX (0.15)
lab.FontSize = 4
TimeOutX (0.15)
lab.FontSize = 6
TimeOutX (0.15)
lab.FontSize = 8
TimeOutX (0.15)
lab.FontSize = 10
TimeOutX (0.15)
lab.FontSize = 12
TimeOutX (0.15)
lab.FontSize = 14
TimeOutX (0.15)
lab.FontSize = 16
TimeOutX (0.15)
lab.FontSize = 18
TimeOutX (0.15)
lab.FontSize = 20
TimeOutX (0.15)
lab.FontSize = 22
TimeOutX (0.15)
lab.FontSize = 24
TimeOutX (0.15)
lab.FontSize = 26
TimeOutX (0.15)
lab.FontSize = 24
TimeOutX (0.15)
lab.FontSize = 22
TimeOutX (0.15)
lab.FontSize = 20
TimeOutX (0.15)
lab.FontSize = 18
TimeOutX (0.15)
lab.FontSize = 16
TimeOutX (0.15)
lab.FontSize = 14
TimeOutX (0.15)
lab.FontSize = 12
TimeOutX (0.15)
lab.FontSize = 10
TimeOutX (0.15)
lab.FontSize = 8
TimeOutX (0.15)
lab.FontSize = 6
TimeOutX (0.15)
lab.FontSize = 4
TimeOutX (0.15)
lab.FontSize = 2
TimeOutX (0.15)
lab.FontSize = 1
lab.Visible = False

End Sub
Sub Label_RedWarp(lbl As Label)
' Fades Label From Red To Black And back
' Call it in a loop to make it go on
' Forever


lbl.ForeColor = 4210752
TimeOutX (0.2)
lbl.ForeColor = 128
TimeOutX (0.2)
lbl.ForeColor = 192
TimeOutX (0.2)
lbl.ForeColor = 255
TimeOutX (0.2)
lbl.ForeColor = 8421631
TimeOutX (0.2)
lbl.ForeColor = 12632319
TimeOutX (0.2)
lbl.ForeColor = 8421631
TimeOutX (0.2)
lbl.ForeColor = 255
TimeOutX (0.2)
lbl.ForeColor = 192
TimeOutX (0.2)
lbl.ForeColor = 128
TimeOutX (0.2)
lbl.ForeColor = 4210752
TimeOutX (0.2)

End Sub

Sub Label_Type(lab As Label)
Dim vString As String
Dim vCount As Integer
Dim vLength As Integer
vString = lab.Caption
vCount = 1
vLength = Len(vString)
Do Until vCount = vLength
lab.Caption = Mid$(vString, vCount, 1)
TimeOutX (1)
vCount = vCount + 1
Loop


End Sub


Sub Form_RedWarp(lbl As Form, Optional NumWarps As Integer)
' Fades Label From Red To Black And back
' Call it in a loop to make it go on
' Forever
If NumWarps = 0 Then NumWarps = 5
Do While NumWarps > 0
lbl.BackColor = 4210752
TimeOutX (0.1)
lbl.BackColor = 128
TimeOutX (0.1)
lbl.BackColor = 192
TimeOutX (0.1)
lbl.BackColor = 255
TimeOutX (0.1)
lbl.BackColor = 8421631
TimeOutX (0.1)
lbl.BackColor = 12632319
TimeOutX (0.1)
lbl.BackColor = 8421631
TimeOutX (0.1)
lbl.BackColor = 255
TimeOutX (0.1)
lbl.BackColor = 192
TimeOutX (0.1)
lbl.BackColor = 128
TimeOutX (0.1)
lbl.BackColor = 4210752
TimeOutX (0.1)
NumWarps = NumWarps - 1
Loop

End Sub



Sub TimeOutX(Duration)
Dim starttime As Long
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub


Sub Form_StayOnTop(theForm As Form)
Dim ret As Long
ret& = SetWindowPos(theForm.hwnd, -1, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub List_AddAsciiStrings(lst As ListBox)
' Use This To Make An Ascii Shop
' be sure to put Naïve (me) on credits, or your just an assclown...
lst.AddItem "[`·¸]"
lst.AddItem "[¸·´]"
lst.AddItem "(•("
lst.AddItem ")•)"
lst.AddItem "(•)"
lst.AddItem ")•("
lst.AddItem "`·´`·´"
lst.AddItem "¸.-~·*'˜¨¯"
lst.AddItem "¨˜'*·~-.¸"
lst.AddItem "¸·`¯¨˜'*·~-.¸"
lst.AddItem "¨ '¹i|¡,¡|i¹'¨"
lst.AddItem "¨‘¨ˆ‘°ª·--·³°ˆ‘ˆ "
lst.AddItem "¨ˆ‘°·-.,,.-·³°¨"
lst.AddItem "·]¦[·"
lst.AddItem "•,¸,.·´¯`·.,¸,•"
lst.AddItem "· .¸·´\"
lst.AddItem "/`·¸. ·"
lst.AddItem "··¤÷×("
lst.AddItem ")×÷¤··"
lst.AddItem "«··÷×)·("
lst.AddItem ")·(×÷··»"
lst.AddItem "(¯`·¸ ["
lst.AddItem "]¸·´¯)"
lst.AddItem "«·÷·¦["
lst.AddItem "]¦·÷·»"
lst.AddItem "•···×"
lst.AddItem "×···•"
lst.AddItem "(`¨˜°º²³³™º˜¨"
lst.AddItem "˜°º™³³²º°˜¨´)"
lst.AddItem "«÷±·´)"
lst.AddItem "(`·±÷»"
lst.AddItem "˜¨°ºðØ"
lst.AddItem "Øðº°¨˜"
lst.AddItem "¸.´)(`·•||["
lst.AddItem "]||•·´)(` .¸"
lst.AddItem "[`·¸][`·÷<l["
lst.AddItem "]l>÷·´][¸·´]"
lst.AddItem "«·÷´'°ºðØ"
lst.AddItem "Øðº°'`÷·»"
lst.AddItem "ðº°¨˜°ºð"
lst.AddItem "ðº°¨˜°ºð"
lst.AddItem "«·÷´)´)"
lst.AddItem "(`(`÷·»"
lst.AddItem "(`·¸_¸.·´)"
lst.AddItem "(`·¸_¸.·´)"
lst.AddItem "(¸.·´)"
lst.AddItem "(¸.·´)"
lst.AddItem "‹›í¦ì‹›"
lst.AddItem "•÷±/)"
lst.AddItem "(\±÷•"
lst.AddItem "(¯`v"
lst.AddItem "v´¯)"
lst.AddItem "•v^÷·•"
lst.AddItem "•·÷^v•"
lst.AddItem "(`·~^v÷"
lst.AddItem "÷v^~·´)"
lst.AddItem "*·._.•÷"
lst.AddItem "÷•._.·*"
lst.AddItem "(¯`·._ (¯`·._"
lst.AddItem "_.·´¯)_.·´¯)"
lst.AddItem "(¯`·._"
lst.AddItem "_.·´¯)"
lst.AddItem "«­.·´¯`•"
lst.AddItem "•´¯`·.-»"
lst.AddItem "‹v÷"
lst.AddItem "÷v›"
lst.AddItem "¸.´)(`·["
lst.AddItem "]·´)(` .¸"
lst.AddItem "·÷±‡±"
lst.AddItem "±‡±÷·"
lst.AddItem "´)·(`"
lst.AddItem "´)·(`"
lst.AddItem "¨˜°ºð"
lst.AddItem "ðº°˜¨"
lst.AddItem "···¤›"
lst.AddItem "‹¤···"
lst.AddItem "«··•[]¤)"
lst.AddItem "(¤[]•··»"
lst.AddItem "···÷±¦)"
lst.AddItem "(¦±÷···"
lst.AddItem "•···¤»"
lst.AddItem "«¤···•"
lst.AddItem "–··•)"
lst.AddItem "(•··–"
lst.AddItem "¤···¤(•)"
lst.AddItem "(•)¤···¤"
lst.AddItem "···÷±|¤"
lst.AddItem "¤|±÷···"
lst.AddItem "()··¤"
lst.AddItem "¤··()"
lst.AddItem "€¨'·.,•"
lst.AddItem "•,.·'¨€"
lst.AddItem "¤··»"
lst.AddItem "«··¤"
lst.AddItem "«¤•···/"
lst.AddItem "\···•¤»"
End Sub

Sub INI_Write(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub


Function INI_Read(AppName, KeyName As String, FileName As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INI_Read = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function


Sub List_CopyListToList(ListA As ListBox, ListB As ListBox)
'revised in ver 3 for speed. it was bad before :)
On Error Resume Next
ListB.Clear
Dim x As Integer
For x = 0 To ListA.ListCount - 1
    ListB.AddItem ListA.List(x)
Next x
End Sub


Sub List_Fonts(lst)
Dim x As Integer
For x = 1 To Screen.FontCount
    lst.AddItem Screen.Fonts(x)
Next x
End Sub
Sub List_RemoveMultiple(lsd)
On Error Resume Next
Dim x As Integer
For x = 1 To lsd.SelCount
lsd.RemoveItem lsd.ListIndex
Next x
End Sub


Sub List_MakeIndexList(ListFrom, ListTo)
On Error Resume Next

ListTo.Clear
ListFrom.ListIndex = 0
Dim x As Integer
Dim nl As String

For x = 0 To ListFrom.ListCount
    ListTo.AddItem ListFrom.ListIndex
    If ListFrom.ListIndex < ListFrom.ListCount + 1 Then
        ListFrom.ListIndex = ListFrom.ListIndex + 1
    End If
Next x
ListTo.RemoveItem ListTo.ListCount - 1
If ListFrom.ListCount <> ListTo.ListCount Then

nl = Chr(13) & Chr(10)
Dim x9 As VbMsgBoxResult

End If
End Sub
Sub List_SetIndexToString(acid As String, ThaBox)
On Error Resume Next
ThaBox.ListIndex = 0
Dim x As Integer
For x = 1 To ThaBox.ListCount
    If ThaBox.Text = acid Then
    ThaBox.ListIndex = x - 1
    Exit Sub
    End If
Next x
End Sub


Function TrimSpaces(str As String)
'recoded! it needed it!
TrimSpaces = Replace(str, " ", "")
End Function
Function TrimString(str As String, totrim As String)
'recoded! it needed it!
TrimString = Replace(str, totrim, "")
End Function


Sub Help_SendKeys()
' A lotta peeps hate sendkys.
' some times they're good though
'
' Special Keys:
'BACKSPACE:       {BACKSPACE}, {BS}, or {BKSP}
'BREAK:           {BREAK}
'CAPS LOCK:       {CAPSLOCK}
'DEL or DELETE:   {DELETE} or {DEL}
'DOWN ARROW:      {DOWN}
'END:             {END}
'ENTER:           {ENTER}or ~
'ESC:             {ESC}
'HELP:            {HELP}
'HOME:            {HOME}
'INSERT:          {INSERT} or {INS}
'LEFT ARROW:      {LEFT}
'NUM LOCK:        {NUMLOCK}
'PAGE DOWN:       {PGDN}
'PAGE UP:         {PGUP}
'PRINT SCREEN:    {PRTSC}
'RIGHT ARROW:     {RIGHT}
'SCROLL LOCK:     {SCROLLLOCK}
'TAB:             {TAB}
'UP ARROW:        {UP}
'F1:              {F1}
'F2:              {F2}
'F3:              {F3}
'F4:              {F4}
'F5:              {F5}
'F6:              {F6}
'F7:              {F7}
'F8:              {F8}
'F9:              {F9}
'F10:             {F10}
'F11:             {F11}
'F12:             {F12}

'More Special Keys:
'SHIFT:   +
'CTRL:    ^
'ALT:     %
'
' Example: SendKeys"%(m)"
' that does alt + m
' the rest is ez
End Sub

Function Web_DipsVBWorld()
Web_DipsVBWorld = "http://come.to/Dipsvbworld"
End Function


Function Web_KnK()
'hehe. the old fun said www.nwozone.com/knk
Web_KnK = "http://www.knk4life.com"
End Function



Sub Form_Center(Frm As Form)
Dim a As Integer
Dim b As Integer
a% = (Screen.Width - Frm.Width) / 2
b% = (Screen.Height - Frm.Height) / 2
Frm.MOVE a%, b%

End Sub

Function zGetRGB(ByVal CVal As Long) As COLORRGB
  ' sub by monkegod
  zGetRGB.Blue = Int(CVal / 65536)
  zGetRGB.Green = Int((CVal - (65536 * zGetRGB.Blue)) / 256)
  zGetRGB.Red = CVal - (65536 * zGetRGB.Blue + 256 * zGetRGB.Green)
End Function



