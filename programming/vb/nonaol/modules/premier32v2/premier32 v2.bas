Attribute VB_Name = "Premier32"
'Premier Thirty Two Bit Module By Galen Grover
'Representing Trumedia Designs 2000
'Zero percent coded for AOL
'I will be glad to help you with any of my coding; Email me at gdgrover@truman.navy.mil
'Http://trumedia.gyrate.org
Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const CB_SHOWDROPDOWN = &H14F

Public Const EM_UNDO = &HC7

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_SETHORIZONTALEXTENT = &H194

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Const SC_SCREENSAVE = &HF140

Public Const SPI_SCREENSAVERRUNNING = 97

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const MAX_PATH = 260

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Sub Array_AddPrefix(TheArray() As String, Prefix As String)

'Adds a string to the end of each item in an array.

'Arguments...
'TheArray(): The array you wish to add the suffix to
'Prefix: The string you wish to add to the beginning of each item in the array

'Example...
'info: This numbers the array [nArray]

'Dim i as integer

'For i% = lbound(nArray) to ubound(nArray)

'Array_AddPrefix nArray(), i% & ". "

'Next i%

Dim i As Integer

For i% = LBound(TheArray) To UBound(TheArray) 'Add prefix to each item in array

    TheArray(i%) = Prefix$ & TheArray(i%)
    
Next i%

End Sub
Public Sub Array_AddSuffix(TheArray() As String, Suffix As String)

'Adds a string to the end of each item in an array.

'Arguments...
'TheArray(): The array you wish to add the suffix to
'Suffix: The string you wish to add to the end of each item in the array

'Example...
'info: Adds the string "@hotmail.com" to each of the items in the array
'      Not actual hotmail accounts

'Dim sArray(1 to 3) as String
'sArray(1) = "premier"
'sArray(2) = "ernpyr"
'sArray(3) = "deomega"

'Array_AddSuffix sArray(), "@hotmail.com"

Dim i As Integer

For i% = LBound(TheArray) To UBound(TheArray) 'Add suffix to each item in array

    TheArray(i%) = TheArray(i%) & Suffix$
    
Next i%

End Sub
Public Function Array_InstrSearch(TheArray() As String, FindWhat As String, Optional ReturnWhat As Integer, Optional SeparationChars As String, Optional MatchCase As Boolean) As String

'This function is a little tough. Here's a breakdown of what it does.
'1. Searches through an array
'2. Finds all array items that contain the search string [FindWhat]
'3. Returns a string containing these items separated by given characters [SeparationChars]
'4. If optional arguments are omitted defaults will be chosen

'Arguments...
'TheArray(): Array to search
'FindWhat: String to find
'ReturnWhat[optional]: Decides what information to return
'   1: Returns what the matching array items equal [default]
'   2: Returns the index of matching array items
'SeparationChars[optional]: What to separate all matching items with [default: " | "]
'CaseSensitive[optional]: Match the case or not [default: false]

'Example...
'note: This will search through all the songs and find all array items
'      that contain the search string "nofx"
'      Last three optional arguments all omitted so they are set to defaults

'Dim Songs as string
'Dim sArray(1 To 5)

'sArray(1) = "Kids of the K Hole - NOFX"
'sArray(2) = "Johnny Appleseed - NOFX"
'sArray(3) = "Alien - PennyWise"
'sArray(4) = "Know Your Enemy - Rage Against The Machine"
'sArray(5) = "Motorcycle Driveby - Third Eye Blind"

'Songs$ = Array_InstrSearch(sArray(), "nofx")
'MsgBox Songs$

'The message box would say "Kids of the K Hole - NOFX | Johny Appleseed - NOFX"

Dim Serch As Integer
Dim FoundM As String
Dim iCase As Integer

If ReturnWhat = 0 Then iCase% = 1
If SeparationChars$ = "" Then SeparationChars$ = " | "
If MatchCase = 0 Then MatchCase = False

Select Case iCase%

    Case 2
    
        For Serch% = LBound(TheArray) To UBound(TheArray)

            If MatchCase = True Then If InStr(TheArray(Serch%), FindWhat$) > 0 Then FoundM$ = FoundM$ & Serch% & SeparationChars$
            If MatchCase = False Then If InStr(LCase$(TheArray(Serch%)), LCase$(FindWhat$)) > 0 Then FoundM$ = FoundM$ & Serch% & SeparationChars$
    
        Next Serch%

        If FoundM$ = "" Then FoundM$ = 0 'If not found set function to 0

        Array_InstrSearch = Left(FoundM$, Len(FoundM$) - Len(SeparationChars$))

    Case 1
    
        For Serch% = LBound(TheArray) To UBound(TheArray)

            If MatchCase = True Then If InStr(TheArray(Serch%), FindWhat$) > 0 Then FoundM$ = FoundM$ & TheArray(Serch%) & SeparationChars$
            If MatchCase = False Then If InStr(LCase$(TheArray(Serch%)), LCase$(FindWhat$)) > 0 Then FoundM$ = FoundM$ & TheArray(Serch%) & SeparationChars$
    
        Next Serch%

        If FoundM$ = "" Then FoundM$ = 0 'If not found set function to 0

        Array_InstrSearch = Left(FoundM$, Len(FoundM$) - Len(SeparationChars$))
    
End Select

End Function
Public Sub Array_Load(TheArray() As String, lFile As String)

'Loads a file into an array

'Arguments...
'TheArray(): Array to load file into
'lFile: File to load into array

'Example...
'note: poo.text contains
'      one
'      two
'      three

'Dim lArray() As String
'dim x as integer
'dim aText as string

'Array_Load lArray(), "c:\poo.txt"

'For x% = LBound(lArray) To UBound(lArray)
'aText$ = aText$ & " " & lArray(x%)
'Next x%

'MsgBox aText$

'The message box would contain "one two three"

Dim freenumber
Dim ArrayItem As String
Dim aItem As Integer
Dim Lines As Integer

Lines% = String_LineCount(String_Load(lFile$))

ReDim TheArray(1 To Lines%)

aItem% = 1  'Initialize variable

If File_Validity(lFile$, 3) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile 'Set variable to freefile

Open lFile$ For Input As #freenumber

    While Not EOF(freenumber) 'Stop when end of file is reached

        Input #freenumber, ArrayItem$ 'Input all lines from text
        DoEvents
        TheArray(aItem%) = ArrayItem$ 'Add item to array
        aItem% = aItem% + 1
    
    Wend
    
Close #freenumber

End Sub
Public Sub Array_Save(TheArray() As String, sFile As String)

'Saves an array to a file.

'Arguments...
'TheArray(): The array you wish to save
'sFile: The destination you wish to save the file to

'Example...
'info: Simply saves the array

'Dim sArray(1 to 3) as String
'sArray(1) = "a"
'sArray(2) = "b"
'sArray(3) = "c"

'Array_Save sArray(), "C:\windows\desktop\array.txt"

Dim i As Integer
Dim freenumber

File_CheckReadOnly sFile$, 1

If File_Validity(sFile$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile

Open sFile$ For Output As #freenumber 'Open file

    For i% = LBound(TheArray) To UBound(TheArray)  'Go through the array
    
        Print #freenumber, TheArray(i%) 'Save each item of array
        
    Next i%
    
Close #freenumber

End Sub
Public Function Array_Search(TheArray() As String, FindWhat As String, Optional SeparationChars As String, Optional MatchCase As Boolean) As String

'This function is very similar to the is. Here's a breakdown of what it does.
'1. Searches through an array
'2. Finds all array items that contain the search string [FindWhat]
'3. Returns a string containing these items separated by given characters [SeparationChars]
'4. If optional arguments are omitted defaults will be chosen

'Arguments...
'TheArray(): Array to search
'FindWhat: String to find
'SeparationChars[optional]: What to separate all matching items with [default: " | "]
'CaseSensitive[optional]: Match the case or not [default: do not match case]

'Example...
'note: This will search through all the songs and find all array items
'      that equal the search string "applesauce"
'      Last three optional arguments all omitted so they are set to defaults

'Dim Songs as string
'Dim sArray(1 To 5)

'sArray(1) = "applesauce"
'sArray(2) = "carrots"
'sArray(3) = "corn"
'sArray(4) = "applesauce"
'sArray(5) = "peas"

'Songs$ = Array_Search(sArray(), "applesauce")
'MsgBox Songs$

'The message box would say "1 | 4"

If SeparationChars$ = "" Then SeparationChars$ = " | "
If MatchCase = 0 Then MatchCase = False
Dim Serch As Integer
Dim FoundM As String

For Serch% = LBound(TheArray) To UBound(TheArray)

    If MatchCase = True Then If TheArray(Serch%) = FindWhat$ Then FoundM$ = FoundM$ & Serch% & SeparationChars$
    If MatchCase = False Then If LCase$(TheArray(Serch%)) = LCase$(FindWhat$) Then FoundM$ = FoundM$ & Serch% & SeparationChars$
    
Next Serch%

If FoundM$ = "" Then FoundM$ = 0 'If not found set function to 0

Array_Search = Left(FoundM$, Len(FoundM$) - Len(SeparationChars$))

End Function
Public Sub Array_ToControl(TheArray() As String, Ctl As Control)

'This will transfer an array to a control.

'Arguments...
'TheArray(): The array you wish to add to the control [Ctl]
'Ctl: The control you want the array added to

'Example...
'info: This will add a, b, and c to the control

'Dim sArray(1 To 3) as String
'sArray(1) = "a"
'sArray(2) = "b"
'sArray(3) = "c"

'Array_ToControl sArray(), List1

Dim i As Integer

For i% = LBound(TheArray) To UBound(TheArray) 'Adds all items in array to control

    Ctl.AddItem TheArray(i%)
    
Next i%

End Sub
Public Function Colors_FixHex(hex As String)

'Fixes hex color codes by adding a 0 in front values containing only 1 character

'Example:

'hex(0) & hex(255) & hex(0) = 0ff0 which is not a valid html color
'However fixhex(hex(0)) & fixhex(hex(255)) & fixhex(hex(0)) = 00ff00

If Len(hex) = 1 Then

    Colors_FixHex = "0" & hex
    
Else

    Colors_FixHex = hex
    
End If

End Function
Public Function Colors_GetBlue(color As Long)
    
'Extracts blue value from a color

'Arguments...
'color: any rgb color

'Example...

'Dim vBlue as long

'vBlue& = Colors_GetBlue(pic.backcolor)

Dim blue As Long
     
blue& = color& \ 65536
    
Colors_GetBlue = blue&

End Function
Public Function Colors_GetGreen(color As Long)

'Extracts green value from a color

'Arguments...
'color: any rgb color

'Example...

'Dim vGreen as long

'vGreen& = Colors_GetGreen(pic.backcolor)

Dim blue As Long
Dim green As Long

blue& = color& \ 65536
green& = (color& - blue& * 65536) \ 256
    
Colors_GetGreen = green&

End Function
Function Colors_GetRed(color As Long)
    
'Extracts red value from a color

'Arguments...
'color: any rgb color

'Example...

'Dim vRed as long

'vRed& = Colors_GetRed(pic.backcolor)

Dim blue As Long
Dim green As Long
Dim red As Long

blue& = color& \ 65536
green& = (color& - blue& * 65536) \ 256
red& = color& - blue& * 65536 - green& * 256

Colors_GetRed = red&

End Function
Public Function Colors_HtmlToRgb(html As Variant)

'Simply converts an html color to an rgb value

'Example...
'hscroll1.value = HtmlToRgb("ff")
'This would set the value of hscroll1 to 255

html = LCase$(html)

If html = "00" Then Colors_HtmlToRgb = 0
If html = "01" Then Colors_HtmlToRgb = 1
If html = "02" Then Colors_HtmlToRgb = 2
If html = "03" Then Colors_HtmlToRgb = 3
If html = "04" Then Colors_HtmlToRgb = 4
If html = "05" Then Colors_HtmlToRgb = 5
If html = "06" Then Colors_HtmlToRgb = 6
If html = "07" Then Colors_HtmlToRgb = 7
If html = "08" Then Colors_HtmlToRgb = 8
If html = "09" Then Colors_HtmlToRgb = 9
If html = "0a" Then Colors_HtmlToRgb = 10
If html = "0b" Then Colors_HtmlToRgb = 11
If html = "0c" Then Colors_HtmlToRgb = 12
If html = "0d" Then Colors_HtmlToRgb = 13
If html = "0e" Then Colors_HtmlToRgb = 14
If html = "0f" Then Colors_HtmlToRgb = 15
If html = "10" Then Colors_HtmlToRgb = 16
If html = "11" Then Colors_HtmlToRgb = 17
If html = "12" Then Colors_HtmlToRgb = 18
If html = "13" Then Colors_HtmlToRgb = 19
If html = "14" Then Colors_HtmlToRgb = 20
If html = "15" Then Colors_HtmlToRgb = 21
If html = "16" Then Colors_HtmlToRgb = 22
If html = "17" Then Colors_HtmlToRgb = 23
If html = "18" Then Colors_HtmlToRgb = 24
If html = "19" Then Colors_HtmlToRgb = 25
If html = "1a" Then Colors_HtmlToRgb = 26
If html = "1b" Then Colors_HtmlToRgb = 27
If html = "1c" Then Colors_HtmlToRgb = 28
If html = "1d" Then Colors_HtmlToRgb = 29
If html = "1e" Then Colors_HtmlToRgb = 30
If html = "1f" Then Colors_HtmlToRgb = 31
If html = "20" Then Colors_HtmlToRgb = 32
If html = "21" Then Colors_HtmlToRgb = 33
If html = "22" Then Colors_HtmlToRgb = 34
If html = "23" Then Colors_HtmlToRgb = 35
If html = "24" Then Colors_HtmlToRgb = 36
If html = "25" Then Colors_HtmlToRgb = 37
If html = "26" Then Colors_HtmlToRgb = 38
If html = "27" Then Colors_HtmlToRgb = 39
If html = "28" Then Colors_HtmlToRgb = 40
If html = "29" Then Colors_HtmlToRgb = 41
If html = "2a" Then Colors_HtmlToRgb = 42
If html = "2b" Then Colors_HtmlToRgb = 43
If html = "2c" Then Colors_HtmlToRgb = 44
If html = "2d" Then Colors_HtmlToRgb = 45
If html = "2e" Then Colors_HtmlToRgb = 46
If html = "2f" Then Colors_HtmlToRgb = 47
If html = "30" Then Colors_HtmlToRgb = 48
If html = "31" Then Colors_HtmlToRgb = 49
If html = "32" Then Colors_HtmlToRgb = 50
If html = "33" Then Colors_HtmlToRgb = 51
If html = "34" Then Colors_HtmlToRgb = 52
If html = "35" Then Colors_HtmlToRgb = 53
If html = "36" Then Colors_HtmlToRgb = 54
If html = "37" Then Colors_HtmlToRgb = 55
If html = "38" Then Colors_HtmlToRgb = 56
If html = "39" Then Colors_HtmlToRgb = 57
If html = "3a" Then Colors_HtmlToRgb = 58
If html = "3b" Then Colors_HtmlToRgb = 59
If html = "3c" Then Colors_HtmlToRgb = 60
If html = "3d" Then Colors_HtmlToRgb = 61
If html = "3e" Then Colors_HtmlToRgb = 62
If html = "3f" Then Colors_HtmlToRgb = 63
If html = "40" Then Colors_HtmlToRgb = 64
If html = "41" Then Colors_HtmlToRgb = 65
If html = "42" Then Colors_HtmlToRgb = 66
If html = "43" Then Colors_HtmlToRgb = 67
If html = "44" Then Colors_HtmlToRgb = 68
If html = "45" Then Colors_HtmlToRgb = 69
If html = "46" Then Colors_HtmlToRgb = 70
If html = "47" Then Colors_HtmlToRgb = 71
If html = "48" Then Colors_HtmlToRgb = 72
If html = "49" Then Colors_HtmlToRgb = 73
If html = "4a" Then Colors_HtmlToRgb = 74
If html = "4b" Then Colors_HtmlToRgb = 75
If html = "4c" Then Colors_HtmlToRgb = 76
If html = "4d" Then Colors_HtmlToRgb = 77
If html = "4e" Then Colors_HtmlToRgb = 78
If html = "4f" Then Colors_HtmlToRgb = 79
If html = "50" Then Colors_HtmlToRgb = 80
If html = "51" Then Colors_HtmlToRgb = 81
If html = "52" Then Colors_HtmlToRgb = 82
If html = "53" Then Colors_HtmlToRgb = 83
If html = "54" Then Colors_HtmlToRgb = 84
If html = "55" Then Colors_HtmlToRgb = 85
If html = "56" Then Colors_HtmlToRgb = 86
If html = "57" Then Colors_HtmlToRgb = 87
If html = "58" Then Colors_HtmlToRgb = 88
If html = "59" Then Colors_HtmlToRgb = 89
If html = "5a" Then Colors_HtmlToRgb = 90
If html = "5b" Then Colors_HtmlToRgb = 91
If html = "5c" Then Colors_HtmlToRgb = 92
If html = "5d" Then Colors_HtmlToRgb = 93
If html = "5e" Then Colors_HtmlToRgb = 94
If html = "5f" Then Colors_HtmlToRgb = 95
If html = "60" Then Colors_HtmlToRgb = 96
If html = "61" Then Colors_HtmlToRgb = 97
If html = "62" Then Colors_HtmlToRgb = 98
If html = "63" Then Colors_HtmlToRgb = 99
If html = "64" Then Colors_HtmlToRgb = 100
If html = "65" Then Colors_HtmlToRgb = 101
If html = "66" Then Colors_HtmlToRgb = 102
If html = "67" Then Colors_HtmlToRgb = 103
If html = "68" Then Colors_HtmlToRgb = 104
If html = "69" Then Colors_HtmlToRgb = 105
If html = "6a" Then Colors_HtmlToRgb = 106
If html = "6b" Then Colors_HtmlToRgb = 107
If html = "6c" Then Colors_HtmlToRgb = 108
If html = "6d" Then Colors_HtmlToRgb = 109
If html = "6e" Then Colors_HtmlToRgb = 110
If html = "6f" Then Colors_HtmlToRgb = 111
If html = "70" Then Colors_HtmlToRgb = 112
If html = "71" Then Colors_HtmlToRgb = 113
If html = "72" Then Colors_HtmlToRgb = 114
If html = "73" Then Colors_HtmlToRgb = 115
If html = "74" Then Colors_HtmlToRgb = 116
If html = "75" Then Colors_HtmlToRgb = 117
If html = "76" Then Colors_HtmlToRgb = 118
If html = "77" Then Colors_HtmlToRgb = 119
If html = "78" Then Colors_HtmlToRgb = 120
If html = "79" Then Colors_HtmlToRgb = 121
If html = "7a" Then Colors_HtmlToRgb = 122
If html = "7b" Then Colors_HtmlToRgb = 123
If html = "7c" Then Colors_HtmlToRgb = 124
If html = "7d" Then Colors_HtmlToRgb = 125
If html = "7e" Then Colors_HtmlToRgb = 126
If html = "7f" Then Colors_HtmlToRgb = 127
If html = "80" Then Colors_HtmlToRgb = 128
If html = "81" Then Colors_HtmlToRgb = 129
If html = "82" Then Colors_HtmlToRgb = 130
If html = "83" Then Colors_HtmlToRgb = 131
If html = "84" Then Colors_HtmlToRgb = 132
If html = "85" Then Colors_HtmlToRgb = 133
If html = "86" Then Colors_HtmlToRgb = 134
If html = "87" Then Colors_HtmlToRgb = 135
If html = "88" Then Colors_HtmlToRgb = 136
If html = "89" Then Colors_HtmlToRgb = 137
If html = "8a" Then Colors_HtmlToRgb = 138
If html = "8b" Then Colors_HtmlToRgb = 139
If html = "8c" Then Colors_HtmlToRgb = 140
If html = "8d" Then Colors_HtmlToRgb = 141
If html = "8e" Then Colors_HtmlToRgb = 142
If html = "8f" Then Colors_HtmlToRgb = 143
If html = "90" Then Colors_HtmlToRgb = 144
If html = "91" Then Colors_HtmlToRgb = 145
If html = "92" Then Colors_HtmlToRgb = 146
If html = "93" Then Colors_HtmlToRgb = 147
If html = "94" Then Colors_HtmlToRgb = 148
If html = "95" Then Colors_HtmlToRgb = 149
If html = "96" Then Colors_HtmlToRgb = 150
If html = "97" Then Colors_HtmlToRgb = 151
If html = "98" Then Colors_HtmlToRgb = 152
If html = "99" Then Colors_HtmlToRgb = 153
If html = "9a" Then Colors_HtmlToRgb = 154
If html = "9b" Then Colors_HtmlToRgb = 155
If html = "9c" Then Colors_HtmlToRgb = 156
If html = "9d" Then Colors_HtmlToRgb = 157
If html = "9e" Then Colors_HtmlToRgb = 158
If html = "9f" Then Colors_HtmlToRgb = 159
If html = "a0" Then Colors_HtmlToRgb = 160
If html = "a1" Then Colors_HtmlToRgb = 161
If html = "a2" Then Colors_HtmlToRgb = 162
If html = "a3" Then Colors_HtmlToRgb = 163
If html = "a4" Then Colors_HtmlToRgb = 164
If html = "a5" Then Colors_HtmlToRgb = 165
If html = "a6" Then Colors_HtmlToRgb = 166
If html = "a7" Then Colors_HtmlToRgb = 167
If html = "a8" Then Colors_HtmlToRgb = 168
If html = "a9" Then Colors_HtmlToRgb = 169
If html = "aa" Then Colors_HtmlToRgb = 170
If html = "ab" Then Colors_HtmlToRgb = 171
If html = "ac" Then Colors_HtmlToRgb = 172
If html = "ad" Then Colors_HtmlToRgb = 173
If html = "ae" Then Colors_HtmlToRgb = 174
If html = "af" Then Colors_HtmlToRgb = 175
If html = "b0" Then Colors_HtmlToRgb = 176
If html = "b1" Then Colors_HtmlToRgb = 177
If html = "b2" Then Colors_HtmlToRgb = 178
If html = "b3" Then Colors_HtmlToRgb = 179
If html = "b4" Then Colors_HtmlToRgb = 180
If html = "b5" Then Colors_HtmlToRgb = 181
If html = "b6" Then Colors_HtmlToRgb = 182
If html = "b7" Then Colors_HtmlToRgb = 183
If html = "b8" Then Colors_HtmlToRgb = 184
If html = "b9" Then Colors_HtmlToRgb = 185
If html = "ba" Then Colors_HtmlToRgb = 186
If html = "bb" Then Colors_HtmlToRgb = 187
If html = "bc" Then Colors_HtmlToRgb = 188
If html = "bd" Then Colors_HtmlToRgb = 189
If html = "be" Then Colors_HtmlToRgb = 190
If html = "bf" Then Colors_HtmlToRgb = 191
If html = "c0" Then Colors_HtmlToRgb = 192
If html = "c1" Then Colors_HtmlToRgb = 193
If html = "c2" Then Colors_HtmlToRgb = 194
If html = "c3" Then Colors_HtmlToRgb = 195
If html = "c4" Then Colors_HtmlToRgb = 196
If html = "c5" Then Colors_HtmlToRgb = 197
If html = "c6" Then Colors_HtmlToRgb = 198
If html = "c7" Then Colors_HtmlToRgb = 199
If html = "c8" Then Colors_HtmlToRgb = 200
If html = "c9" Then Colors_HtmlToRgb = 201
If html = "ca" Then Colors_HtmlToRgb = 202
If html = "cb" Then Colors_HtmlToRgb = 203
If html = "cc" Then Colors_HtmlToRgb = 204
If html = "cd" Then Colors_HtmlToRgb = 205
If html = "ce" Then Colors_HtmlToRgb = 206
If html = "cf" Then Colors_HtmlToRgb = 207
If html = "d0" Then Colors_HtmlToRgb = 208
If html = "d1" Then Colors_HtmlToRgb = 209
If html = "d2" Then Colors_HtmlToRgb = 210
If html = "d3" Then Colors_HtmlToRgb = 211
If html = "d4" Then Colors_HtmlToRgb = 212
If html = "d5" Then Colors_HtmlToRgb = 213
If html = "d6" Then Colors_HtmlToRgb = 214
If html = "d7" Then Colors_HtmlToRgb = 215
If html = "d8" Then Colors_HtmlToRgb = 216
If html = "d9" Then Colors_HtmlToRgb = 217
If html = "da" Then Colors_HtmlToRgb = 218
If html = "db" Then Colors_HtmlToRgb = 219
If html = "dc" Then Colors_HtmlToRgb = 220
If html = "dd" Then Colors_HtmlToRgb = 221
If html = "de" Then Colors_HtmlToRgb = 222
If html = "df" Then Colors_HtmlToRgb = 223
If html = "e0" Then Colors_HtmlToRgb = 224
If html = "e1" Then Colors_HtmlToRgb = 225
If html = "e2" Then Colors_HtmlToRgb = 226
If html = "e3" Then Colors_HtmlToRgb = 227
If html = "e4" Then Colors_HtmlToRgb = 228
If html = "e5" Then Colors_HtmlToRgb = 229
If html = "e6" Then Colors_HtmlToRgb = 230
If html = "e7" Then Colors_HtmlToRgb = 231
If html = "e8" Then Colors_HtmlToRgb = 232
If html = "e9" Then Colors_HtmlToRgb = 233
If html = "ea" Then Colors_HtmlToRgb = 234
If html = "eb" Then Colors_HtmlToRgb = 235
If html = "ec" Then Colors_HtmlToRgb = 236
If html = "ed" Then Colors_HtmlToRgb = 237
If html = "ee" Then Colors_HtmlToRgb = 238
If html = "ef" Then Colors_HtmlToRgb = 239
If html = "f0" Then Colors_HtmlToRgb = 240
If html = "f1" Then Colors_HtmlToRgb = 241
If html = "f2" Then Colors_HtmlToRgb = 242
If html = "f3" Then Colors_HtmlToRgb = 243
If html = "f4" Then Colors_HtmlToRgb = 244
If html = "f5" Then Colors_HtmlToRgb = 245
If html = "f6" Then Colors_HtmlToRgb = 246
If html = "f7" Then Colors_HtmlToRgb = 247
If html = "f8" Then Colors_HtmlToRgb = 248
If html = "f9" Then Colors_HtmlToRgb = 249
If html = "fa" Then Colors_HtmlToRgb = 250
If html = "fb" Then Colors_HtmlToRgb = 251
If html = "fc" Then Colors_HtmlToRgb = 252
If html = "fd" Then Colors_HtmlToRgb = 253
If html = "fe" Then Colors_HtmlToRgb = 254
If html = "ff" Then Colors_HtmlToRgb = 255

End Function
Public Sub ComboBox_DropDown(Combo As ComboBox)

'Drops down a combo

'Arguments...
'Combo: The combo you wish to drop down

'Example...
'ComboBox_DropDown Combo1

Dim mCombo As Long

mCombo& = SendMessage(Combo.hwnd, CB_SHOWDROPDOWN, True, 0)
Combo.SetFocus

End Sub
Public Sub ComboBox_RollUp(Combo As ComboBox)

'Drops down a combo

'Arguments...
'Combo: The combo you wish to drop down

'Example...
'ComboBox_DropDown Combo1

Dim mCombo As Long

mCombo& = SendMessage(Combo.hwnd, CB_SHOWDROPDOWN, False, 0)
Combo.SetFocus

End Sub
Public Function Control_Compare(Ctl1 As Control, Ctl2 As Control, MatchCase As Boolean)

'This will return the number of alike items two controls contain

Dim i As Integer
Dim i2 As Integer
Dim Comp As String
Dim DaCount As Integer

For i% = 0 To Ctl1.ListCount - 1
    
    Comp$ = Ctl1.List(i%)
    
    For i2% = 0 To Ctl2.ListCount - 1
            
        If MatchCase = True Then
            
            If Comp$ = Ctl2.List(i2%) Then DaCount% = DaCount% + 1: Exit For 'If item in ctl1 same as item in ctl2 then add 1 to variable
                
        Else
            
            If LCase$(Comp$) = LCase$(Ctl2.List(i2%)) Then DaCount% = DaCount% + 1: Exit For 'If item in ctl1 same as item in ctl2 then add 1 to variable
                
        End If
            
    Next i2%
    
Next i%
                
Control_Compare = DaCount%

End Function
Public Sub Control_Sort(Ctl As Control)

'You can't change the sort property of a listbox during runtime
'So I designed this sub
'The whole basis behind this sub is that > can be used to alphabetize

'Example...
'Control_Sort(list1)

Dim i As Integer
Dim X As Integer
Dim Temp As String

For i% = 0 To Ctl.ListCount - 2

    For X% = i% + 1 To Ctl.ListCount - 1
    
        If Ctl.List(i%) > Ctl.List(X%) Then 'Use <> to alphabetize
        
            Temp$ = Ctl.List(i%)
            Ctl.List(i%) = Ctl.List(X%)
            Ctl.List(X%) = Temp$
            
        End If
        
    Next X%
    
Next i%

End Sub
Public Sub Control_SizeToForm(Ctl As Control, frm As Form)

'This will resize a control to fit the form
'Put this in the resize procedure of the form

'Example...
'Control_SizeToForm text1

Ctl.Width = frm.ScaleWidth
Ctl.Height = frm.ScaleHeight

End Sub
Public Sub Control_Save(Ctl As Control, FullPath As String)

'Saves a control

'Example...
'Control_Save playlist, "c:\windows\desktop\playlist.txt"

Dim i As Integer
Dim freenumber

File_CheckReadOnly FullPath$, 1

If File_Validity(FullPath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile 'Set variable to free file

Open FullPath$ For Output As #freenumber 'Open file

    For i% = 0 To Ctl.ListCount 'Go through the whole list
    
        Print #freenumber, Ctl.List(i%) 'Save each line of the list
        
    Next i%
    
Close #freenumber

File_SetNormal FullPath$
    
End Sub
Public Sub Control_AppendSave(Ctl As Control, FullPath As String)

'Saves a control, but appends it to a file

'Example...
'Control_Save playlist, "c:\windows\desktop\playlist.txt"

Dim i As Integer
Dim freenumber

File_CheckReadOnly FullPath$, 1

If File_Validity(FullPath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile 'Set variable to free file

Open FullPath$ For Append As #freenumber 'Open file

    For i% = 0 To Ctl.ListCount 'Go through the whole list
    
        Print #freenumber, Ctl.List(i%) 'Save each line of the list
        
    Next i%
    
Close #freenumber

File_SetNormal FullPath$
    
End Sub
Public Sub Control_Load(Ctl As Control, FullPath As String)

'Loads a control

'Example...
'Control_Load list1, "c:\windows\desktop\pw.txt"

Dim freenumber
Dim ctlitem As String

If File_Validity(FullPath$, 3) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile 'Set variable to free file

Open FullPath$ For Input As #freenumber

    While Not EOF(freenumber) 'Stop when end of file is reached

        Input #freenumber, ctlitem$ 'Input all lines from text
        DoEvents
        Ctl.AddItem ctlitem$ 'Add item to control
    
    Wend
    
Close #freenumber

If Ctl.List(Ctl.ListCount - 1) = "" Then Ctl.ListIndex = Ctl.ListCount - 1: Ctl.RemoveItem Ctl.ListIndex 'If last item is "" then remove it

End Sub
Public Sub Control_Transfer(FromCtl As Control, ToCtl As Control, ClearToCtl As Boolean, LowCase As Boolean, ClearFromCtl As Boolean)

'This will copy a control to another control
'Lists, comboboxes and combinations

'Arguments...
'FromCtl: Control to take items form
'ToCtl: Control to add items to
'ClearCtl: Clear from control or not

'Example...
'Control_Transfer list1, combo1, true, true, true

Dim i As Integer

If FromCtl.ListCount = 0 Then Exit Sub 'If origin control is empty exit sub
If ClearToCtl = True Then ToCtl.Clear 'Clear target control if desired

For i% = 0 To FromCtl.ListCount - 1

    If LowCase = True Then 'Check case option
    
        ToCtl.AddItem LCase$(FromCtl.List(i%)) 'Add lowercase item to target control
        
    Else
    
        ToCtl.AddItem FromCtl.List(i%) 'Add item to target control
        
    End If
    
Next i%

outoffor:
If ClearFromCtl = True Then FromCtl.Clear: Exit Sub 'Clear origin control if desired

End Sub
Public Sub Control_AddSuffix(Ctl As ListBox, Suffix As String)

'Adds a string to the end of each item in a control

'Example...
'List_AddSuffix snlist, "@aol.com"

Dim X As Integer
Dim i As Integer
Dim ctlcount As Integer

ctlcount% = Ctl.ListCount

For X% = 0 To Ctl.ListCount - 1

    Ctl.AddItem Ctl.List(X%) & Suffix$
    
Next X%

For i% = 1 To ctlcount%

    Ctl.ListIndex = 0
    Ctl.RemoveItem Ctl.ListIndex
    
Next i%

End Sub
Public Sub Control_AddPrefix(Ctl As Control, Prefix As String)

'Adds a prefix to each item in a control

'Example...

'Dim i as integer
'For i% = 0 to list1.listcount-1
'Control_AddPrefix list1, i% & ". "
'Next i%

'This would number the items in the control

Dim X As Integer
Dim i As Integer
Dim ctlcount As Integer

ctlcount% = Ctl.ListCount

For X% = 0 To Ctl.ListCount - 1

    Ctl.AddItem Prefix & Ctl.List(X%)
    
Next X%

For i% = 1 To ctlcount%

    Ctl.ListIndex = 0
    Ctl.RemoveItem Ctl.ListIndex
    
Next i%

End Sub
Public Sub Control_EqualsSearch(Ctl As Control, SearchFor As String, CaseSensitive As Boolean)

'Will highlight the first instance of a string in a listbox
'Will set the text of a combo to the first instance of a string

'Arguments...
'Ctl: Control to search in
'SearchFor: What to search for in the control
'CaseSensitive: Match case or not

'Example...
'Control_EqualSearch list1, "eminem", true

Dim i As Integer

For i% = 0 To Ctl.ListCount - 1

    If CaseSensitive = True Then
    
        If Ctl.List(i%) = SearchFor$ Then Ctl.ListIndex = i%: Exit Sub
        
    Else
    
        If LCase$(Ctl.List(i%)) = LCase$(SearchFor$) Then Ctl.ListIndex = i%: Exit Sub
        
    End If
    
Next i%

MsgBox "Search string not found.", 64, "Search..."

End Sub
Public Sub Control_InstrSearch(Ctl As Control, SearchFor As String, CaseSensitive As Boolean)

'Will highlight the first item in a listbox containing the search string
'Will set the text of a combo to the first item containing the search string

'Arguments...
'Ctl: Control to search in
'SearchFor: What to search for in the control
'CaseSensitive: Match case or not

'Example...
'Control_EqualSearch list1, "emi", true

Dim i As Integer

For i% = 0 To Ctl.ListCount - 1

    If CaseSensitive = True Then
    
        If InStr(Ctl.List(i%), SearchFor$) <> 0 Then Ctl.ListIndex = i%: Exit Sub
        
    Else
    
        If InStr(LCase$(Ctl.List(i%)), LCase$(SearchFor$)) <> 0 Then Ctl.ListIndex = i%: Exit Sub
        
    End If
    
Next i%

MsgBox "Search string not found.", 64, "Search..."

End Sub
Public Sub Control_InstrSearchToList(SearchCtl As Control, SearchFor As String, Ctl2 As Control, Optional MatchCase As Boolean)

'Will search a list and add each item containing search string to a second list

'Arguments...
'SearchCtl: Control to search in
'SearchFor: What to search for in the control
'MatchCase: Match case or not

'Example...
'Control_InstrSearchToList list1, "emin", list2

Dim i As Integer

If MatchCase = True And InStr(Control_ToNumberedString(SearchCtl), SearchFor$) = False Then MsgBox "Search string not found.", 64, "Search..."
If MatchCase = False And InStr(LCase$(Control_ToNumberedString(SearchCtl)), LCase$(SearchFor$)) = False Then MsgBox "Search string not found.", 64, "Search..."

For i% = 0 To SearchCtl.ListCount - 1

    If MatchCase = False Then
    
        If InStr(LCase$(SearchCtl.List(i%)), LCase$(SearchFor$)) <> 0 Then Ctl2.AddItem SearchCtl.List(i%)
    
    Else
    
        If InStr(SearchCtl.List(i%), SearchFor$) <> 0 Then Ctl2.AddItem SearchCtl.List(i%)

    End If
    
Next i%

End Sub
Public Sub Control_DeleteDuplicates(Ctl As Control)

'Deletes duplicate instances in a listbox

Dim i As Integer
Dim X As Integer
Dim dp As String

For i% = 0 To Ctl.ListCount - 1
    dp$ = Ctl.List(i%)
    For X% = 0 To Ctl.ListCount - 1
        If LCase$(Ctl.List(X%)) Like LCase$(dp$) Then Ctl.ListIndex = X%: Ctl.RemoveItem Ctl.ListIndex
    Next X%
Next i%
    

End Sub
Public Function Control_InstrMatches(Ctl As Control, SearchFor As String, CaseSensitive As Boolean)

'Will return number of items in listbox containing the search string

'Arguments...
'Ctl: Control to search in
'SearchFor: What to search for in the control
'CaseSensitive: Match case or not

'Example...
'Msgbox "There are " & Control_InstrMatches(list1, "emin", false) " songs containing " & searchfor$ "." ,64, "info..."

Dim i As Integer
Dim count As Integer

For i% = 0 To Ctl.ListCount - 1

    If CaseSensitive = True Then
    
        If InStr(Ctl.List(i%), SearchFor$) <> 0 Then count% = count% + 1
        
    Else
    
        If InStr(LCase$(Ctl.List(i%)), LCase$(SearchFor$)) <> 0 Then count% = count% + 1
        
    End If
    
Next i%

Control_InstrMatches = count%

End Function
Public Function Control_EqualMatches(Ctl As Control, SearchFor As String, Optional CaseSensitive As Boolean)

'Will return number of items in listbox are equal to the search string

'Arguments...
'Ctl: Control to search in
'SearchFor: What to search for in the control
'CaseSensitive: Match case or not

'Example...
'Msgbox "There are " & Control_EqualMatches(list1, "emin", false) " songs containing " & searchfor$ "." ,64, "info..."

Dim i As Integer
Dim count As Integer

For i% = 0 To Ctl.ListCount - 1

    If CaseSensitive = True Then
    
        If Ctl.List(i%) = SearchFor$ Then count% = count% + 1
        
    Else
    
        If LCase$(Ctl.List(i%)) = LCase$(SearchFor$) Then count% = count% + 1
        
    End If
    
Next i%

Control_EqualMatches = count%

End Function
Public Function Control_ToNumberedString(Ctl As Control)

'This will take the items in a control and put them into a numbered string where
'each item is numbered according to its place in the list

'Example...
'Songlist$ = Control_ToNumberedString(list1)

Dim i As Integer
Dim numstr As String

For i% = 0 To Ctl.ListCount - 1

    numstr$ = numstr$ & i% + 1 & ".] " & Ctl.List(i%) & vbCrLf
    
Next i%

Control_ToNumberedString = numstr$

End Function
Public Sub Control_AddFonts(Ctrl As Control)

'Adds system fonts to a control

'Example...
'Control_AddFonts fontcombo

Dim i As Integer

For i% = 1 To Screen.FontCount 'Go through each font

    Ctrl.AddItem Screen.Fonts(i%) 'Add each font to the control
    
Next i%

End Sub
Public Sub Control_AddAsciis(Ctrl As Control)

'Adds ascii's to a control

'Example...
'Control_AddAsciis list1

Dim i As Integer

For i% = 33 To 255 'Go through all the asciis

    Ctrl.AddItem Chr$(i%) 'Add each ascii to the control
    
Next i%

End Sub
Public Function CP_CtrlAltDel(enabled As Boolean)

'Enables and disables ctl alt del function

'Examples...

'CP_CtrlAltDel true - will enable ctl alt del
'CP_CtrlAltDel false - will disable ctl alt del

Dim lReturn  As Long
Dim lBool As Long

If enabled = False Then lReturn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, lBool, vbNull)
If enabled = True Then lReturn = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, lBool, vbNull)

End Function
Function CP_DriveEmpty(DriveLetter As String) As Boolean

'Returns true if drive is empty, false if it is not empty

'Arguments...
'DriveLetter: Drive letter to check if empty

'Example...
'If CP_DriveEmpty("a") = True Then exit sub

Dim dirct As String

On Error Resume Next 'Continue because we need the error

dirct = Dir$(DriveLetter$ & ":\*.*")

If Err.Number = 52 Then '52 is empty drive error

    CP_DriveEmpty = True
    
Else

    CP_DriveEmpty = False
    
End If

End Function
Public Sub CP_ShowDesktop()

'Shows the desktop

'Example...
'CP_ShowDesktop

On Error Resume Next

CP_Run "C:\WINDOWS\SYSTEM\Show Desktop.scf"

End Sub
Public Sub CP_NoBeep()

'This will stop the textboxes from beeping
'Copy and paste the code into the keypress procedure of a textbox

Dim KeyCode

If KeyCode = 13 Then KeyCode = 0

End Sub
Public Function CP_TempPath() As String

'Returns computer's temp path

'Example...
'Textbox_Save text1, cp_temppath & "\temp.txt"

Dim strfldr As String
Dim lngrslt As Long

strfldr$ = String(MAX_PATH, 0)
lngrslt& = GetTempPath(MAX_PATH, strfldr)

If lngrslt& <> 0 Then

  CP_TempPath = Left(strfldr$, InStr(strfldr$, Chr(0)) - 1)
  
Else

  CP_TempPath = ""
  
End If

End Function
Public Sub CP_StandBy(frm As Form)

'Turns your computer on standby mode

'Exampl...
'CP_StandBy me

Call SendMessage(frm.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)

End Sub
Public Sub CP_Run(file As String)

'Opens anything you want

'Exampl...
'Run "c:\" will open up the c folder

Dim hwnd

Call ShellExecute(hwnd, "Open", file$, "", App.Path, 1)

End Sub
Public Function CP_SystemPath() As String

'Returns windows system directory

'Example...
'File_Cut "c:\windows\desktop\threed32.ocx", CP_SystemPath & "\Threed32.ocx"

Dim strfldr As String
Dim lngrslt As Long

strfldr = String(MAX_PATH, 0)
lngrslt = GetSystemDirectory(strfldr, MAX_PATH)

If lngrslt <> 0 Then

  CP_SystemPath = Left(strfldr, InStr(strfldr, Chr(0)) - 1) & "\"
  
Else

  CP_SystemPath = ""
  
End If

End Function
Public Sub Directory_Delete(drctry As String)

'Deletes Directory

'Example...
'Directory_Delete "C:\Windows\"

'Dont ever do that

If InStr(drctry, ":\") = 0 Then Exit Sub

RmDir drctry$

End Sub
Public Sub Directory_Create(drctry As String)

'Creates a directory

'Example...
'Directory_Create "C:\Windows\Desktop\My Program\"

On Error GoTo endit:

If InStr(drctry$, ":\") = 0 Then Exit Sub

MkDir drctry$

endit: Exit Sub

End Sub
Public Function Files_ToControl(Ctl As Control, Directory As String, Optional ExtFilter As String, Optional RemovExt As Boolean)

'Adds files in a given directoory matching extension filter to a control

'Arguments...
'Ctl: Control you wish to add the files to
'Directory: The directory that you want to search for files in
'ExtFilter: The extension filter [default:

Dim fFile As String

If ExtFilter$ = "" Then ExtFilter = "*.*"
If Right(Directory$, 1) <> "\" Then Directory$ = Directory$ & "\"

Select Case RemovExt

Case False

    fFile$ = Dir(Directory$ & ExtFilter)

    Do Until fFile$ = ""

        Ctl.AddItem fFile$
        fFile$ = Dir

    Loop
    
Case True

    fFile$ = Dir(Directory$ & ExtFilter)

    Do Until fFile$ = ""
    
        If InStr(fFile$, ".") = 0 Then GoTo noEXT:
        
        fFile$ = Left(fFile$, Len(fFile$) - (Len(fFile$) - InStr(fFile$, ".")) - 1)
                
noEXT:

        Ctl.AddItem fFile$
        fFile$ = Dir

    Loop
    
End Select

End Function
Public Function File_CheckReadOnly(FullPath As String, SetNormal As Boolean)

'This function will see if a file is readonly and if so will set the file to normal if desired

'Arguments...
'FullPath: Full path of file to check
'SetNormal: If the file is read only this decides whether or not to set it normal

'Example...
'CP_SaveDialog Me, "All Formats", "*.htm;*.txt", "Save As...", Label1
'File_CheckReadOnly Label1.caption, True
'TextBox_Save Text1, Label1.caption
        
Select Case SetNormal
    
    Case True
    
        If File_GetAttributes(FullPath, 2) = "1" Or _
        File_GetAttributes(FullPath, 2) = "3" Or _
        File_GetAttributes(FullPath, 2) = "5" Or _
        File_GetAttributes(FullPath, 2) = "33" Or _
        File_GetAttributes(FullPath, 2) = "35" Or _
        File_GetAttributes(FullPath, 2) = "37" Then
        
            File_CheckReadOnly = True
            
        Else
        
            File_CheckReadOnly = False
            
        End If
        
        File_SetNormal (FullPath$)
        
    Case False
        
        If File_GetAttributes(FullPath, 2) = "1" Or _
        File_GetAttributes(FullPath, 2) = "3" Or _
        File_GetAttributes(FullPath, 2) = "5" Or _
        File_GetAttributes(FullPath, 2) = "33" Or _
        File_GetAttributes(FullPath, 2) = "35" Or _
        File_GetAttributes(FullPath, 2) = "37" Then
        
            File_CheckReadOnly = True
        
        Else
        
            File_CheckReadOnly = False
        
        End If
    
End Select

End Function
Public Sub File_Copy(originfilepath As String, newfilepath As String)

'Copies file
'Using the built in sub is just as easy but this will check for file validity first

'Exampl...
'File_Copy "C:\My shit\butt sex.jpg", "C:\Homework\History\butt sex.jpg"

If File_Validity(originfilepath$, 1) = False Or File_Validity(newfilepath$, 3) = False Then Exit Sub

FileCopy originfilepath$, newfilepath$

End Sub
Public Sub File_Cut(originfilepath As String, newfilepath As String)

'Copies file and deletes original

'Example...
'File_Copy "C:\My shit\butt sex.jpg", "C:\Homework\History\butt sex.jpg"

If File_Validity(originfilepath$, 1) = False Or File_Validity(newfilepath$, 3) = False Then Exit Sub

FileCopy originfilepath$, newfilepath$
Kill originfilepath$

End Sub
Public Function File_GetAttributes(FileFullPath As String, Form)

'Gets the attributes of a file
'Case one will return the name of the attribute
'Case two will return the integer of the attribute
'0 = Normal | 1 = ReadOnly | 2 = Hidden | 4 = System | 32 = Archive and all combinations

Dim daattr As Integer

Select Case Form

    Case 1 'Return string
    
        If File_Validity(FileFullPath$, 3) = False Then Exit Function
        
        daattr% = GetAttr(FileFullPath$) 'Get integer
        
        If daattr% = 0 Then File_GetAttributes = "Normal"
        If daattr% = 1 Then File_GetAttributes = "ReadOnly"
        If daattr% = 2 Then File_GetAttributes = "Hidden"
        If daattr% = 3 Then File_GetAttributes = "ReadOnly and System"
        If daattr% = 4 Then File_GetAttributes = "System"
        If daattr% = 5 Then File_GetAttributes = "ReadOnly, Hidden and System"

        If daattr% = 32 Then File_GetAttributes = "Archive"
        If daattr% = 33 Then File_GetAttributes = "Archive and ReadOnly"
        If daattr% = 34 Then File_GetAttributes = "Archive and Hidden"
        If daattr% = 35 Then File_GetAttributes = "Archive and ReadOnly and Hidden"
        If daattr% = 36 Then File_GetAttributes = "Archive and System"
        If daattr% = 37 Then File_GetAttributes = "Archive, ReadOnly, Hidden and System"

    Case 2 'Return integer
        
        If File_Validity(FileFullPath$, 3) = False Then Exit Function
        
        File_GetAttributes = GetAttr(FileFullPath$)

End Select

End Function
Public Sub File_SetNormal(FileFullPath$)

'Sets file attribute to normal

If File_Validity(FileFullPath$, 3) = False Then Exit Sub

SetAttr FileFullPath$, vbNormal

End Sub
Public Sub File_SetReadOnly(FileFullPath$)

'Sets file attribute to read only

If File_Validity(FileFullPath$, 3) = False Then Exit Sub

SetAttr FileFullPath$, vbReadOnly

End Sub
Public Sub File_SetHidden(FileFullPath$)

'Sets file attribute to hidden

If File_Validity(FileFullPath$, 3) = False Then Exit Sub

SetAttr FileFullPath$, vbHidden

End Sub
Public Sub File_SetArchive(FileFullPath$)

'Sets file attribute to archive

If File_Validity(FileFullPath$, 3) = False Then Exit Sub

SetAttr FileFullPath$, vbArchive

End Sub
Public Sub File_SetSystem(FileFullPath$)

'Sets file attribute to read only

If File_Validity(FileFullPath$, 3) = False Then Exit Sub

SetAttr FileFullPath$, vbSystem

End Sub
Public Function File_GetDirectory(gFile As String) As String

'Returns the directory of a file given the full path

Dim i As Integer
Dim start As Integer

For i% = Len(gFile$) To 1 Step -1

If Mid(gFile$, i%, 1) = "\" Then File_GetDirectory = Left(gFile$, i%): Exit Function

Next i%

End Function
Public Sub File_TextAppend(strng As String, FullPath As String)

'ADDS to a text file already saved

'Arguments...
'Strng: String to append to file
'FullPath: Paht of file to append

'Example...
'File_TextAppend vbcrlf & "Freedom [RATM]", "C:\my shit\song list.txt"

Dim freenumber
Dim DaText As String

File_CheckReadOnly FullPath$, 1

If File_Validity(FullPath$, 1) = False Then Exit Sub 'Check for validity of file

freenumber = FreeFile
DaText$ = strng$

Open FullPath$ For Append As #freenumber 'Open the path to edit it

    Print #freenumber, DaText$ 'Add the string to the existing file

Close #freenumber 'Close file

End Sub
Public Function File_Validity(FileFullPath As String, TheCase) As Boolean

'This will check all aspects of file validity

'Arguments...
'FileFullPath: Full path of file to validate
'TheCase: Described below

'Case 1: File existance
'Case 2: File name validity
'Case 3: Both cases

'Example...

'Dim Myfile as string

'Myfile$ = "C:\windows\desktop\pwlist.txt"

'If File_Validity(Myfile$, 3) = true then

'    Text_Load Myfile, pwtext.Text

'Else

'    Msgbox "File does not exist" ,64,"error..."

'End If

Select Case TheCase

    Case 1 'Check existance of file
    
        On Error GoTo Done: 'On error goto label "done:"
        FileLen (FileFullPath$) 'Gets size of file, error if file doesn't exist, hence on error
        
Done:

        File_Validity = False: Exit Function 'File doesn't exist therefore fnction is false
        File_Validity = True
        
    Case 2 'Search filename for necessary and illegal characters
    
        If InStr(FileFullPath, ":\") = 0 Then File_Validity = False: Exit Function
        If InStr(Right(FileFullPath, 5), ".") = 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "?") <> 0 Then File_Validity = False:  Exit Function
        If InStr(FileFullPath, "*") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "<") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, ">") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, Chr(34)) <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "|") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "/") <> 0 Then File_Validity = False: Exit Function
        File_Validity = True

    Case 3 'Perform both cases
    
        On Error GoTo done2:
        
        FileLen (FileFullPath$)
        
        'Search filename for necessary and illegal characters
        If InStr(FileFullPath, ":\") = 0 Then File_Validity = False: Exit Function
        If InStr(Right(FileFullPath, 5), ".") = 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "?") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "*") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "<") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, ">") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, Chr(34)) <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "|") <> 0 Then File_Validity = False: Exit Function
        If InStr(FileFullPath, "/") <> 0 Then File_Validity = False: Exit Function
        
        File_Validity = True
        

End Select

File_Validity = True: Exit Function

done2:      File_Validity = False: Exit Function 'From case 3... file doesn't exist

End Function
Public Sub File_Delete(FileFullPath As String)

'Deletes file
'Careful with this one there little booger

'Example: File_Delete "C:\Windows\my porn\sister in shower.jpg"

If File_Validity(FileFullPath$, 1) = False Then Exit Sub 'Check file for existance

Kill FileFullPath$ 'Delete file

End Sub
Public Sub File_Rename(FileFullPath As String, newfilefullpath As String)

'Renames file
'Using the built in sub is just as easy but this will check for file validity first

'Example...
'File_Rename "C:\My shit\horse sex.jpg", "C:\My shit\term paper.txt"

If File_Validity(FileFullPath$, 1) = False Or File_Validity(newfilefullpath$, 3) = False Then Exit Sub

Name FileFullPath$ As newfilefullpath$

End Sub
Public Function File_GetFile(fPath As String) As String

'This functions returns the file name of a file given the full path

'Arguments...
'fPath: Path of file including file to extract file name from

'Example...
'msgbox File_GetFile("C:\win.ini")

'msgbox would contain "win.ini"

Dim i As Integer
Dim start As Integer

For i% = Len(fPath$) To 1 Step -1
    
    If Mid(fPath$, i%, 1) = "\" Then File_GetFile = Right(fPath$, Len(fPath$) - i%): Exit Function

Next i%

End Function
Public Sub Form_FullScreen(frm As Form)

'Usually used if forms borderstyle is set to none

With frm

    .Top = 0
    .Left = 0
    .Width = Screen.Width
    .Height = Screen.Height

End With

End Sub
Public Sub Form_Move(frm As Form)

'Will move a form
'Put this in the mousedown procedure

Call ReleaseCapture
Call SendMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)

End Sub
Public Sub Form_SizePosition(frm As Form, Left As Integer, Top As Integer, Optional Height As Integer, Optional Width As Integer)

'Set the position of the form

'Arguments...
'Frm: Form
'Left: New left of form
'Top: New top of form
'Height[if null height won't be changed]: New height of form
'Width[if null width won't be changed]: New width of form

'Example...
'Form_Position me, 0, 0

frm.Left = Left% 'Set left
frm.Top = Top% 'Set top
If Not Height = "" Then frm.Height = Height% 'Set height
If Not Width = "" Then frm.Width = Width% 'Set width

End Sub
Public Sub Form_SetTop(frm As Form)

'Form will always be topmost window
'Example: Form_SetTop me

Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)

End Sub
Public Sub Form_SetNotTop(frm As Form)

'Form will no longer be on top

'Example...
'Form_SetNotTop me

Call SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)

End Sub
Public Sub Form_Center(frm As Form)

'Centers a form
'Usually used in form_load

frm.Left = Screen.Width / 2 - frm.Width / 2 'Set middle of width to middle of screen width
frm.Top = Screen.Height / 2 - frm.Height / 2 'Set middle of height to middle of screen height

End Sub
Public Sub Form_CenterAt(frm As Form, X As Integer, Y As Integer)

'Centers a form at a given point on the screen

'Arguments...
'Frm: the form you wish to center
'X: X coordinate to center form at
'Y: Y coordinate to center form at

'Example...
'info: That example would center the form at the middle of the screen

'Form_CenterAt me, screen.width/2, screen.height/2

frm.Left = X - frm.Width / 2
frm.Top = Y - frm.Height / 2

End Sub
Public Function Ini_Read(PathofIni As String, Section As String, Key As String) As Variant

'Gets the info from a ini for a specific item

'[Section]
'Key=KeyValue

'Returns KeyValue

'Example...
'If Ini_Read("Options", "intro art", "c:\program files\trumedia designs\trumedia html editor.ini") = false then main.load


Dim buf As String

buf = String(750, Chr(0))
Key$ = LCase$(Key$)
Ini_Read = Left(buf$, GetPrivateProfileString(Section$, ByVal Key$, "", buf$, Len(buf$), PathofIni$))

End Function
Public Sub Ini_Write(PathofIni As String, Section As String, Key As String, KeyValue As String)

'Writes to an ini file

'PathofIni: ? figure it out

'[Section]
'Key=KeyValue

'Example...
'Ini_Write "Options", "intro art", "false", "c:\program files\trumedia designs\trumedia html editor.ini"

File_CheckReadOnly PathofIni$, 1
Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, PathofIni$)

End Sub
Public Sub Label_Spell(sLabel As Label, sString As String, Speed As String)

'This will spell a string to a labels caption

'Arguments...
'sLabel: Label you want to spell on
'sString: String you want spelled out

'Speeds:    [time in between in letter placed]
'       1[1.00 seconds]
'       2[0.90 seconds]
'       3[0.80 seconds]
'       4[0.70 seconds]
'       5[0.60 seconds]
'       6[0.50 seconds]
'       7[0.40 seconds]
'       8[0.30 seconds]
'       9[0.20 seconds]
'      10[0.10 seconds]

Dim i As Integer
Dim speeda As Single
Dim spell As String

If Number_Valid(Speed$) = False Then Exit Sub

If Speed = 1 Then speeda! = 1 'define speeds
If Speed = 2 Then speeda! = 0.9
If Speed = 3 Then speeda! = 0.8
If Speed = 4 Then speeda! = 0.7
If Speed = 5 Then speeda! = 0.6
If Speed = 6 Then speeda! = 0.5
If Speed = 7 Then speeda! = 0.4
If Speed = 8 Then speeda! = 0.3
If Speed = 9 Then speeda! = 0.2
If Speed = 10 Then speeda! = 0.1

For i% = 1 To Len(sString$)

    spell$ = Mid(sString$, i%, 1) 'Set variable to letter
    sLabel.Caption = sLabel.Caption & spell$ 'Add letter to textbox
    Program_Pause speeda! 'Timeout to  control thespeed

Next i%

End Sub
Public Sub List_RemoveSelected(Lst As ListBox)

'Removes all selected items from a listbox

Dim i As Integer

For i% = Lst.ListCount - 1 To 0 Step -1

    If Lst.Selected(i%) Then Lst.RemoveItem i% 'If list item is selected remove it
    
Next i

End Sub
Public Sub List_HScrollBar(Lst As ListBox)

'Gives a listbox horizontal a horizontal scrollbar

Dim DoIt As Long
Dim wID As Integer

wID% = Lst.Width + 1 'new width in pixels
DoIt& = SendMessage(Lst.hwnd, LB_SETHORIZONTALEXTENT, wID%, ByVal 0&)

End Sub
Public Sub List_CopySelectedToControl(Lst As ListBox, Ctl As Control)

'Copies all selected items from a listbox and adds them to a control

'Arguments...
'Lst: Listbox to copy items from
'Ctl: Control to add items to

'Example...
'List_CopySelectedToControl list1, combo1

Dim i As Integer

For i% = Lst.ListCount - 1 To 0 Step -1
    If Lst.Selected(i%) Then Ctl.AddItem Lst.List(i%) 'If list item is selected add it to control
    
Next i%

End Sub
Public Sub List_RemoveSelectedToControl(Lst As ListBox, Ctl As Control)

'Removes all selected items from a listbox and adds them to a control

'Arguments...
'Lst: Listbox to remove items from
'Ctl: Control to add items to

'Example...
'List_RemoveSelectedToControl list1, combo1

Dim i As Integer

For i% = Lst.ListCount - 1 To 0 Step -1

    If Lst.Selected(i%) Then Ctl.AddItem Lst.List(i%): Lst.RemoveItem i% 'If list item is selected add it to control and remove it

Next i%

End Sub

Public Function Math_Factorial(Nmbr As Integer)

'Returns the factorial of a number
'A factorial is defined as:
'n! = n(n-1)(n-2)......(1)

'Example...
'5! = 5(4)(3)(2)(1) = 120
'Therefore Math_Factorial(5) would return 120

Dim i As Integer
Dim Fact As Integer

Fact% = Nmbr% 'Set variable to number

For i% = Nmbr% - 1 To 1 Step -1 'Step down from Nmbr -1 to 1

    Fact% = Fact% * i% 'Multiply total by total - 1
    
Next i%

Math_Factorial = Fact%

End Function

Public Function Math_Quad(Aval As Single, Bval As Single, Cval As Single) As String

'Returns the computed quadratic equation of A, B, and C in a string form
'Quadratic Equation is (-B +- SquareRoot(B^2 - 4AC)) / 2A

'Breaking down the equation into parts
Dim negB As Single
Dim sqB As Single
Dim m4AC As Single
Dim m2A As Single

'Get the parts
negB! = -Bval!
sqB! = Bval! ^ 2
m4AC! = Aval! * Cval! * 4
m2A! = 2 * Aval!

'Final answer
Math_Quad = "(" & negB! & "+- sqr(" & sqB! - m4AC! & "))/" & m2A!

End Function
Public Sub Menu_RunByNumber(ProgramClassName As String, TopMenu As Long, SubMenu As Long)

'Runs any programs menu by a set of numbers
'0 is the first number in the top menu and sub menu

'Arguments...
'ProgramClassName: Class name of program. CAn be obtained with a window spy
'TopMenu: Top menu you want to run 0 is first top menu
'Submenu: Sub menu you want to run 0 is first sub menu

'Example...
'Menu_RunByNumber "Aol Frame25", 0, 0
'This will execute new in the file menu of aol

Dim ClassName As Long
Dim Menu1 As Long
Dim Menu2 As Long
Dim MenuId As Long

ClassName& = FindWindow(ProgramClassName$, vbNullString)
Menu1& = GetMenu(ClassName&)
Menu2& = GetSubMenu(Menu1&, TopMenu&)
MenuId& = GetMenuItemID(Menu2&, SubMenu&)

Call SendMessageLong(ClassName&, WM_COMMAND, MenuId&, 0&)

End Sub
Public Function Misc_Bingo() As String

'Why did i make this?

'Generates a random bingo number such as B14

Dim letter As String
Dim rBingo As Integer

rBingo% = Number_RandomCustom(1, 75)

If 1 <= rBingo% And rBingo% <= 15 Then letter$ = "B"
If 16 <= rBingo% And rBingo% <= 30 Then letter$ = "I"
If 31 <= rBingo% And rBingo% <= 45 Then letter$ = "N"
If 46 <= rBingo% And rBingo% <= 60 Then letter$ = "G"
If 61 <= rBingo% And rBingo% <= 75 Then letter$ = "O"


Misc_Bingo$ = letter$ & rBingo%

End Function
Public Sub Misc_RelationalCenter(Centerwhat As Variant, InRelationTo As Variant)

'Centers something in relation to something else... hmmm

'Arguments...
'CenterWhat: What ot center
'InRelationTo: What ot center it in relation to

'Example...
'Misc_RelationalCenter label1, picture1

'This will center label1 in picture1

Centerwhat.Left = InRelationTo.Width / 2 - Centerwhat.Width / 2 'Set middle of width to middle of inrelationto width
Centerwhat.Top = InRelationTo.Height / 2 - Centerwhat.Height / 2 'Set middle of height to middle of inrelationto height

End Sub
Public Sub Net_Email(EmailAddress As String)

'Sends email to given address

'Arguments...
'EmailAddress: Email address to send email to

'Example...
'Net_Email "klowde@netzero.net"

Dim hwnd As Long

ShellExecute hwnd&, "open", "mailto:" + EmailAddress$, vbNullString, vbNullString, 5

End Sub
Sub Net_Webpage(Address As String)

'Opens a web page

'Arguments...
'Address: Address of webpage to open

'Example...
'Net_Webpage "http://trumedia.gyrate.org"

Dim hwnd As Long

Call ShellExecute(hwnd&, "Open", Address$, "", App.Path, 1)

End Sub
Public Function Number_Valid(Var As Variant) As Boolean

'Finds if a string is a number

'Arguments...
'Var: variant to check if valid number

'Example...
'if Number_Valid(text1) = false then exit sub

Number_Valid = IsNumeric(Var)

End Function
Public Function Number_EvenOrOdd(Number As Integer) As String

'Returns if number is even or odd or decimal

'Arguments...
'Number: Number to decide even or odd

'Example...
'if Number_EvenorOdd(hscroll1.value) = even then goto continue
'exit sub
'continue:

Dim Operate As Single

If Number_Valid(Number%) = False Then Exit Function 'Exit if not a number

If InStr(Number%, ".") <> 0 Then Number_EvenOrOdd = "decimal": Exit Function 'Set function to decimal and exit

Operate! = Number% / 2 'Divide number by 2

If InStr(Operate!, ".") <> 0 Then 'If . is found number was odd
    
    Number_EvenOrOdd = "odd" 'So set function to odd
    
Else

    Number_EvenOrOdd = "even" 'Else number was even
    
End If

End Function
Public Function Number_RandomCustom(Low As Integer, High As Integer)

'Generates random number between and including low and high

'Example...
'Dim guess As Integer
'Dim correct As Integer
'Dim try
'correct% = Number_RandomCustom(1, 100)
'start:
'On Error Resume Next
'guess% = InputBox("Enter a number from 1 to 100", "Guessing game")
'If guess% = 0 Then Exit Sub
'If Number_Valid(guess%) = False Then MsgBox "Enter a number", vbCritical, "error...": GoTo start:
'If guess% = correct% Then MsgBox "Correct answer", 64, "Good Job!!!": Exit Sub
'try = MsgBox("Wrong Answer, Try again?", vbYesNo, "Try Again?")
'If try = vbYes Then GoTo start
'Exit Sub
        


Dim High2 As Integer
Dim Darnd As Integer

High2% = High% - Low% + 1 'Fix random so high and low ends are variables
Randomize 'Initialize random number generator
Darnd% = Int((Rnd * High2%) + Low%) 'Generate random number

Number_RandomCustom = Darnd%

End Function
Public Function Number_Remainder(DivideWhat As Variant, IntoWhat As Variant)

'Returns the remainder when two numbers are divided

'Arguments...
'DivideWhat: divisor
'IntoWhat: dividend

'Example: Number_Remainder(4, 10) returns 2

If Number_Valid(DivideWhat) Or Number_Valid(IntoWhat) = False Then Exit Function

Number_Remainder = IntoWhat Mod DivideWhat

End Function
Public Function Number_Percent(Done As Variant, Total As Variant) As Integer

'Returns percent

'Arguments...
'Done: How much is completed
'Total: Total to be done

'Example...
'label1.caption = Number_Percent(50, 100)

Dim X As Integer
Dim perc As Integer

Number_Percent = (Done / Total) * 100

End Function
Public Function Number_Random(High As Integer)

'Generates random number from 0 to high using custom version

'Arguments...
'High: High, low will be 0

'Example...
'picture1.backcolor = Number_Random(255)

Number_Random = Number_RandomCustom(High%, "0")

End Function
Public Function Number_UnDecimal(Number As Variant)

'Basically just removes the decimal from a number

If InStr(Number, ".") = 0 Then Number_UnDecimal = Number

Number_UnDecimal = Left(Number, InStr(Number, ".") - 1)


End Function
Public Sub Picturebox_RandomColor(Pic As PictureBox)

'This sub will generate a random backcolor for a picturebox

'Arguments...
'Pic: Picturebox to give random backcolor

'Example...
'Colors_RandomPictureboxColor(pic1)

Pic.BackColor = RGB(Number_Random(255), Number_Random(255), Number_Random(255))

End Sub
Public Function Password_Random(Number_Characters As Integer, Optional Different_Cases As String) As String

'Arguments...
'Number_Characters: Number of character to generate
'Differents cases[default: false]: whether or not string contains both cases
'Different cases will wither be "both" or "lcase"

Dim rPwArray As Integer
Dim rpw As String
Dim i As Integer

If Different_Cases = "" Then Different_Cases = False

Select Case LCase("Different_Cases$")

Case "both"

    Dim bPwArray(1 To 3) As Variant

    For i% = 1 To Number_Characters%

        bPwArray(1) = Number_RandomCustom(48, 57)
        bPwArray(2) = Number_RandomCustom(97, 122)
        bPwArray(3) = Number_RandomCustom(65, 90)

        rPwArray% = Number_RandomCustom(1, 3)

        rpw$ = rpw$ & Chr$(bPwArray(rPwArray%))

    Next i%

Case "lcase"

    Dim lPwArray(1 To 2) As Variant

    For i% = 1 To Number_Characters%

        lPwArray(1) = Number_RandomCustom(48, 57)
        lPwArray(2) = Number_RandomCustom(97, 122)

        rPwArray% = Number_RandomCustom(1, 2)

        rpw$ = rpw$ & Chr$(lPwArray(rPwArray%))

    Next i%

Case "ucase"

    Dim uPwArray(1 To 2) As Variant

    For i% = 1 To Number_Characters%

        uPwArray(1) = Number_RandomCustom(48, 57)
        uPwArray(2) = Number_RandomCustom(65, 90)

        rPwArray% = Number_RandomCustom(1, 2)

        rpw$ = rpw$ & Chr$(uPwArray(rPwArray%))

    Next i%

End Select

Password_Random$ = rpw$

End Function
Public Sub Picturebox_Fade(Colors() As Long, Pic2Fade As PictureBox)

'Fades a picturebox using an array of colors

'Arguments...
'Colors(): Array of colors described as below
'pic2fade: picture to fade

'Example...
'        dim colors(1 to 3) as long 'The second number is the number of colors
'        colors(1) = picture1.backcolor
'        colors(2) = picture3.backcolor
'        colors(3) = picture3.backcolor
'        Picturebox_Fade colors(), picture1

'Array can also be in rgb format
'Example...
'        dim colors(1 to 3) as long 'The second number is the number of colors
'        colors(1) = rgb(number_random(255), number_random(255), number_random(255))
'        colors(2) = rgb(number_random(255), number_random(255), number_random(255))
'        colors(3) = rgb(number_random(255), number_random(255), number_random(255))
'        Picturebox_Fade colors(), picture1


'IMPORTANT:  MAKE SURE TO DIM THE ARRAY AS LONG OR SUB WILL NOT WORK

Dim ArrayStart As Integer
Dim ArrayEnd As Integer
Dim GetColor As Integer
Dim Changes As Integer
Dim place As Integer
Dim Fading As Integer
Dim Lines As Integer

ArrayStart% = LBound(Colors) 'Get beginning of array
ArrayEnd% = UBound(Colors) 'Get end of array

ReDim Change(1 To ArrayEnd%, 1 To 3) As Single 'Set limits
ReDim IndivColor(ArrayStart To ArrayEnd% + 1, 1 To 3) As Single 'Set limits

For GetColor% = 1 To ArrayEnd% 'Get rgb values for each color

    IndivColor!(GetColor%, 1) = Colors_GetRed(Colors(GetColor%))
    IndivColor!(GetColor%, 2) = Colors_GetGreen(Colors(GetColor%))
    IndivColor!(GetColor%, 3) = Colors_GetBlue(Colors(GetColor%))

Next GetColor%

For Changes% = 1 To ArrayEnd% 'Get the changes needed to fade to next color

    Change!(Changes%, 1) = (IndivColor!(Changes% + 1, 1) - IndivColor!(Changes%, 1)) / (Pic2Fade.Width / (ArrayEnd% - 1))
    Change!(Changes%, 2) = (IndivColor!(Changes% + 1, 2) - IndivColor!(Changes%, 2)) / (Pic2Fade.Width / (ArrayEnd% - 1))
    Change!(Changes%, 3) = (IndivColor!(Changes% + 1, 3) - IndivColor!(Changes%, 3)) / (Pic2Fade.Width / (ArrayEnd% - 1))
    
Next Changes%

For Fading% = 1 To ArrayEnd% - 1

    For Lines% = (Fading% - 1) * (Pic2Fade.Width / (ArrayEnd% - 1)) To Fading% * (Pic2Fade.Width / (ArrayEnd% - 1))

        'Add line to picturebox with color
        Pic2Fade.Line (Lines%, 0)-(Lines%, Pic2Fade.Height), RGB(IndivColor!(Fading%, 1), IndivColor!(Fading%, 2), IndivColor!(Fading%, 3))
        
        'Add changes to rgb values to fade to next color
        IndivColor!(Fading%, 1) = IndivColor!(Fading%, 1) + Change!(Fading%, 1)
        IndivColor!(Fading%, 2) = IndivColor!(Fading%, 2) + Change!(Fading%, 2)
        IndivColor!(Fading%, 3) = IndivColor!(Fading%, 3) + Change!(Fading%, 3)
           
    Next Lines%

Next Fading%

End Sub
Public Sub Program_Pause(Duration As Single)

'Pauses the program for a given duration

'Arguments...
'Duration: Duration of pause

'Example...
'dim i as integer
'for i% = 1 to len(props$)
'    text1.seltext = mid(props$, i%, 1)
'next i%

Dim starttime

starttime = Timer 'Set variable to timer

Do While Timer - starttime < Duration!
    
    DoEvents

Loop

End Sub
Public Sub Program_Hide()

'Hides your program from the program list
'I use this for trick programs

Dim one As Long
Dim two As Long

one& = GetCurrentProcessId()
two& = RegisterServiceProcess(one&, RSP_SIMPLE_SERVICE)

End Sub
Public Sub Program_Unhide()

'Unhides your program from the program list
'I use this for trick programs

Dim one As Long
Dim two As Long

one& = GetCurrentProcessId()
two& = RegisterServiceProcess(one&, RSP_UNREGISTER_SERVICE)

End Sub
Public Sub Program_Exit()

'Simply an exit query
'Uncomment and put this whole code in the unload procedure of the form

'Dim response

'response = MsgBox("Do You want to exit?", vbYesNo, App.Title)

'If response = vbYes Then End
'If response = vbNo Then Cancel = 1: Exit Sub

End Sub
Public Sub Program_About(frm As Form, Caption As String, Optional Copyright As String)

'Professional looking about box for your program
'This will add the forms icon to the about box

'Arguments...
'Frm: Form
'Caption: Caption of about box
'Copyright: copyright

'Example...
'Program_AboutBox me, "About", "Trumedia Designs"

If VarType(Copyright) = vbString Then
    
    Call ShellAbout(frm.hwnd, Caption$, Copyright$, frm.Icon)

Else
    
    Call ShellAbout(frm.hwnd, Caption$, "", frm.Icon)

End If

End Sub
Public Function String_CharCount(strng As String, Char2Count As String, Optional MatchCase As Boolean) As Integer

'Counts the number of times Char2Count appears in Strng

'Arguments...
'Strng: String to count characters in
'Char2Count: Character or characters to count
'MatchCase[default: false]: Match case or not

'Example...
'dim x as string
'x$ = String_CharCount(text1.text, "@hotmail.com")

Dim X As Integer
Dim countc As Integer

If MatchCase = "" Then MatchCase = False

If InStr(strng$, Char2Count) = 0 Then Exit Function

For X% = 1 To Len(strng$)

    If MatchCase = True Then
    
        If Mid(strng$, X%, Len(Char2Count$)) = Char2Count$ Then countc% = countc% + 1
        
    ElseIf MatchCase = False Then
    
        If LCase$(Mid(strng$, X%, Len(Char2Count$))) = LCase$(Char2Count$) Then countc% = countc% + 1

    End If
    
Next X%

String_CharCount = countc%

End Function
Public Function String_Cryption(strng As String, Encrypt_Decrypt)

'This is an encryption algorithm
'What it does is it converts the character to it ascii integer,
'adds a constant, and gets the character of the new ascii integer
'So i guess you could call it kind of a shift encyption

'Arguments...
'Strng: String to encrypt
'Encrypt_Decrypt: String... either "encrypt" or "decrypt" minus quotes ofcourse

'Example...
'text1.text = String_Cryption(text1.text, 1)...this will encrypt the text in text1
'text1.text = String_Cryption(text1.text, 2)...this will decrypt the text in text1

Dim i As Integer
Dim char As String
Dim ascii As Integer
Dim newchar As String
Dim NewString As String

Select Case LCase$(Encrypt_Decrypt)
    
    Case "encrypt" 'Encrypt
    
        For i% = 1 To Len(strng$)
            
            char$ = Mid(strng$, i%, 1) 'Get the char
            ascii% = Asc(char$) + 10 'Change char to asc and add 64 to encrypt
            newchar$ = Chr$(ascii%) 'Change new asc to new character
            NewString$ = NewString$ & newchar$ 'Add new character to encrypted string
        
        Next i
        
    String_Cryption = NewString$

    Case "decrypt" 'Decrypt
        
        For i% = 1 To Len(strng$)
            
            char$ = Mid(strng$, i%, 1) 'Get the char
            ascii% = Asc(char$) - 10 'Change char to asc and subtract the 64 used to encrypt
            newchar$ = Chr$(ascii%) 'Change new asc back to new character
            NewString$ = NewString$ & newchar$ 'Add new character to decrypted string
        
        Next i
        
    String_Cryption = NewString$

End Select

End Function
Public Sub String_OpenAsWebPage(strng As String)

'Arguments...
'Strng: String to open as a webpage

'Example...
'String OpenAsWebPage text1.text

String_Save strng$, "c:\temp.html"
Net_Webpage "c:\temp.html"

'Uncomment these to kill the temporary file
'Program_Pause 3
'Kill "c:\temp.html"

End Sub
Public Function String_RandomLetter(Capitalize As Boolean)

'Generates random letter

'Arguments...
'Capitalize: Capitalize the letter or not

'Example...
'dim letter as string
'letter$ = String_RandomLetter(false)

Dim letter As Integer

If Capitalize = False Then String_RandomLetter = LCase$(Chr$(Number_RandomCustom(65, 90)))
If Capitalize = True Then String_RandomLetter = Chr$(Number_RandomCustom(65, 90))

End Function
Public Function String_Scramble(strng As String)

'This randomly scrambles a string

'Arguments...
'Strng: String to scramble

'Example...
'dim scramble as string
'scramble$ = String_Scramble(text1.text)

Dim Length As Integer
Dim part As String
Dim point As Integer
Dim checkrandom As String
Dim scrstr As String
Dim times As Integer

Length% = Len(strng$)
checkrandom$ = ","
times% = 0

Do

startagain:
    
    point% = Number_Random(Length%) + 1
    If point% > Length% Then GoTo startagain:
    If InStr(checkrandom$, "," & point% & ",") = 0 Then GoTo skip:

    GoTo startagain:

skip:
    
    checkrandom = checkrandom$ & point% & ","
    scrstr$ = scrstr$ & Mid(strng$, point%, 1)
    times% = times% + 1

Loop Until times% = Length%


String_Scramble = scrstr$

End Function
Public Function String_Fade(Colors() As Long, String2Fade As String) As String

'Fades a string using an array of colors

'Arguments...
'Colors(): Array of colors described as below
'String2Fade: String ot fade

'Example...
'        dim colors(1 to 3) as long the second number is the number of colors
'        colors(1) = picture1.backcolor
'        colors(2) = picture3.backcolor
'        colors(3) = picture3.backcolor
'        String_Save String_Fade(colors(), text1), "c:\windows\desktop\fade.html"

'Array can also be in RGB format
'Example...
'        dim colors(1 to 3) as long 'The second number is the number of colors
'        colors(1) = rgb(number_random(255), number_random(255), number_random(255))
'        colors(2) = rgb(number_random(255), number_random(255), number_random(255))
'        colors(3) = rgb(number_random(255), number_random(255), number_random(255))
'        String_Save String_Fade(colors(), text1), "c:\windows\desktop\fade.html"

'IMPORTANT: BE SURE TO DIM YOUR COLOR ARRAY AS LONG OR YOU WILL GET AN ERROR

Dim ArrayStart As Integer
Dim ArrayEnd As Integer
Dim GetColor As Integer
Dim Changes As Integer
Dim Fading As Integer
Dim Chars As Integer
Dim fadedstring As String
Dim TheMod As Integer
Dim Length As Integer

ArrayStart% = LBound(Colors) 'Get beginning of array
ArrayEnd% = UBound(Colors) 'Get end of array

If ArrayEnd% > Len(String2Fade$) Then Exit Function

ReDim Change(1 To ArrayEnd%, 1 To 3) As Single 'Set limits
ReDim IndivColor(ArrayStart To ArrayEnd% + 1, 1 To 3) As Integer 'Set limits

For GetColor% = 1 To ArrayEnd% 'Get rgb values for each color

    IndivColor(GetColor%, 1) = Colors_GetRed(Colors(GetColor%))
    IndivColor(GetColor%, 2) = Colors_GetGreen(Colors(GetColor%))
    IndivColor(GetColor%, 3) = Colors_GetBlue(Colors(GetColor%))

Next GetColor%

If InStr(Len(String2Fade$) / ArrayEnd%, ".") = True Then 'Get number of letters
    
    'Find remainder when length of string is divided by number of colors
    TheMod% = Len(String2Fade$) Mod ArrayEnd%
    'Subtract remainder so numbers will divide evenly
    Length% = Len(String2Fade$) - TheMod%
    
Else
    'If numbers already divide evenly set length to length of string
    Length% = Len(String2Fade$)
    
End If

For Changes% = 1 To (ArrayEnd%) 'Get the changes needed to fade to next color
    
    Change(Changes%, 1) = (IndivColor(Changes% + 1, 1) - IndivColor(Changes%, 1)) / (Length% / (ArrayEnd% - 1))
    Change(Changes%, 2) = (IndivColor(Changes% + 1, 2) - IndivColor(Changes%, 2)) / (Length% / (ArrayEnd% - 1))
    Change(Changes%, 3) = (IndivColor(Changes% + 1, 3) - IndivColor(Changes%, 3)) / (Length% / (ArrayEnd% - 1))
    
Next Changes%



For Fading% = 1 To ArrayEnd% - 1

    For Chars% = ((Fading% - 1) * (Length% / (ArrayEnd% - 1))) + 1 To Fading% * (Length% / (ArrayEnd% - 1))

        'Add to faded text string
        fadedstring$ = fadedstring$ & "<font color=" & Colors_FixHex(hex(IndivColor(Fading%, 1))) & Colors_FixHex(hex(IndivColor(Fading%, 2))) & Colors_FixHex(hex(IndivColor(Fading%, 3))) & ">" & Mid(String2Fade$, Chars%, 1)
                        
        'Add changes to rgb values to fade to next color
        IndivColor(Fading%, 1) = IndivColor(Fading%, 1) + Change(Fading%, 1)
        IndivColor(Fading%, 2) = IndivColor(Fading%, 2) + Change(Fading%, 2)
        IndivColor(Fading%, 3) = IndivColor(Fading%, 3) + Change(Fading%, 3)
           
    Next Chars%

Next Fading%

String_Fade = fadedstring$

End Function
Public Function String_FirstLine(strng As String)

'Returns first line of a string

'Arguemnts...
'Strng: String to get first line from

'Example...
'dim lastline as string
'firstline$ = String_FirstLine (text1.text)

Dim rspot As Integer

If InStr(strng$, Chr$(10)) = 0 Then String_FirstLine = strng$: Exit Function
rspot% = InStr(strng$, Chr$(10)) 'Find return
String_FirstLine = Left(strng$, rspot% - 2) 'Get everything to the left of the return not includiong return

End Function
Public Function String_LastLine(strng As String)

'Returns last line of a string

'Arguments...
'Strng: String to get last line from

'Example...
'dim lastline as string
'lastline$ = String_LastLine (text1.text)

Dim i As Integer
Dim rspot As Integer

For i% = Len(strng$) To 1 Step -1

    If Mid(strng$, i%, 1) = Chr$(10) Then String_LastLine = Right(strng$, Len(strng$) - i%): Exit Function
    
Next i%

End Function
Public Function String_LineCount(strng As String)

'This counts the amount of lines in a string

'Arguments...
'Strng: String to get line count of

'Example...
'dim linecount as string
'linecount$ = String_LineCount(text1)

Dim i As Integer
Dim count As Integer
Dim Look As String

If strng$ = "" Then Exit Function 'Exit if string is empty

count% = 1 'set count to one because it has to have atleast one line

For i% = 1 To Len(strng)
    
    Look$ = Mid(strng$, i%, 1) 'Search each character
    If Look$ = Chr$(13) Then count% = count% + 1 'If character is return[chr$(13)] add 1 to line count

Next i%

String_LineCount = count%

End Function
Public Function String_GetLine(strng As String, Line As Integer, Optional Return_Place As Integer)

'Gets a certain line from a string

'Arguments...
'Strng: String ot get line from
'Line: Line to get
'Return_Place:

'Example:
'dim daline as string
'dim returnP as string
'daline$ = String_Getline(mystring$, 4, returnP)

'returnP now equals the char position in the textbox of the first character of the line
'So now you can do something like this

'text1.selstart=returnp
'text1.sellength = len(ReturnedString$)
'text1.setfocus

'That will highlight the line you retrieved :)

On Error GoTo fixit:

Dim i As Integer
Dim X As Integer
Dim time As Integer

If Line% = 1 Then String_GetLine = String_FirstLine(strng$): Exit Function

fixed:
time% = 0
i% = 0
X% = 0

Do

    i% = InStr(i% + 1, strng$, Chr$(13) & Chr$(10))
    time% = time% + 1
    
Loop Until time% = Line% - 1

X% = InStr(i% + 1, strng$, Chr$(13) & Chr$(10))

String_GetLine = Mid(strng$, i% + 2, X% - i% - 2)
Return_Place% = i% + 1
Exit Function

fixit: strng$ = strng$ & Chr$(13) & Chr$(10): GoTo fixed:

End Function
Public Function String_Load(FullPath As String)

'Load a string

'Arguments...
'Fullpath: Path of file to load into string variable

'Example...
'dim songlist as string
'songlist$ = String_Load("C:\My shit\songlist.txt")

Dim DaText As String
Dim freenumber

If File_Validity(FullPath$, 3) = False Then Exit Function

freenumber = FreeFile

Open FullPath$ For Input As #freenumber

    String_Load = Input(LOF(freenumber), #freenumber)

Close #freenumber

End Function
Public Function String_Reverse(rString As String)

'Simply reverses a string

'Arguments...
'rString: The string to reverse

'Example...
'msgbox String_Reverse("premier")

'Return "reimerp"
'Basically useless

Dim i As Integer
Dim reverse As String

For i% = Len(rString$) To 1 Step -1 'Start at the end
    
    reverse$ = reverse$ & Mid(rString$, i%, 1) 'Get last charcter and add it to variable

Next i%

String_Reverse = reverse$

End Function
Public Sub String_Save(strng As String, FullPath As String)

'Saves a string instead of saving a texbox

'Arguments...
'Strng: String to save
'FullPath: Patht to save file to

'Example...
'String_Save List_ToNumberedString(songlistbox)

Dim freenumber

File_CheckReadOnly FullPath$, 1
If File_Validity(FullPath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile

Open FullPath$ For Output As #freenumber

    Print #freenumber, strng$ 'Print text to file

Close #freenumber

End Sub
Public Function String_Replace(rString As String, ReplaceWhat As String, ReplaceWith As String, Optional MatchCase As Boolean, Optional MessageBox As Boolean)

'This will replace every instance of a string within a string with another
'string and return a string with the replace string replaced
'AHAHAHA Sound confusing? -Good!

'Arguments...
'rString: The whole string
'ReplaceWhat: What string to replace
'ReplaceWith: What to replace all instances of the string with
'MatchCase [optional]: Match case or not [default: false]
'MessageBox [optional]: If set to true a msgbox appears when all instances
                        'are replaced with the number of replacements made [default: true]

'Example...
'String_Replace aolscreename$, "@aol.com", ""
'This will replace all @aol.com with "" or nothing
'So if aolscreename$ equaled "kwest one@aol.com, premier zero@aol.com"
'It would return kwest one, premier zero

Dim spot As Integer
Dim theleft As String
Dim theright As String
Dim danewstring As String
Dim times As Integer

If MatchCase = 0 Then MatchCase = False

If MatchCase = True And InStr(rString$, ReplaceWhat$) = 0 Then MsgBox "Search text not found.", 64, "Done...": String_Replace = "": Exit Function
If MatchCase = False And InStr(LCase$(rString$), LCase$(ReplaceWhat$)) = 0 Then MsgBox "Search text not found.", 64, "Done...": String_Replace = "": Exit Function

Do
    If MatchCase = False Then
        
        spot% = InStr(LCase$(rString$), LCase$(ReplaceWhat$))
    
    Else
        
        spot% = InStr(rString$, ReplaceWhat$)
    
    End If
    
    theleft$ = Left(rString$, spot% - 1)
    theright$ = Right(rString$, Len(rString$) - (Len(theleft$) + Len(ReplaceWhat$)))
    rString$ = theleft$ & ReplaceWith$ & theright$
    times% = times% + 1

Loop Until InStr(rString$, ReplaceWhat$) = 0

String_Replace = rString$

If MessageBox = False Then Exit Function
If MessageBox = True Then MsgBox "Search complete, " & times% & " replacements made.", vbInformation, "Search Complete..."

End Function
Public Function String_TrimNull(strng As String)

'Removes all null characters chr$(32) or space

'Arguments...
'Strng: String to trim null characters from

'Example...
'dim newstring as string
'newstring$ = String_TrimNull (mystring$)

String_TrimNull = String_Replace(strng$, " ", "", False, False)

End Function
Public Sub String_SplitToArray(Split_String As String, Delimiter As String, Arry() As String)

'Splits a string using given characters and adds them to a control

'Arguments...
'Split_String: String to split
'Delimiter: Chars to use to separate string
'Arry(): Array to add separated string to

'Example...

'Dim x as integer
'Dim strray(1 to 5)

'strray(1) = "a1+a2+a3"
'strray(2) = "b1+b2+b3"
'strray(3) = "c1+c2+c3"
'strray(4) = "d1+d2+d3"
'strray(5) = "e1+e2+e3"

'For X = 1 To 5

'       String_SplitToControl strray(x), "+", list1

'Next x

Dim Split_Spot As Integer
Dim Split_Start As Integer
Dim i As Integer

i% = 1
Split_Start% = 1

Do

    Split_Spot% = InStr(Split_Start%, Split_String$, Delimiter$) 'Find split characters
    
    If Split_Spot% = 0 Then Arry$(i) = Mid(Split_String$, Split_Start%, Len(Split_String$)): Exit Sub 'If split characters not found exit sub
    
    Arry$(i) = Mid(Split_String$, Split_Start%, Split_Spot% - Split_Start%)
    Split_Start% = Split_Spot% + Len(Delimiter$)
    
    i% = i% + 1
    
Loop

End Sub
Public Sub String_SplitToControl(Split_String As String, Delimiter As String, Ctl As Control)

'Splits a string using given characters and adds them to a control

'Arguments...
'Split_String: String to split
'Delimiter: Chars to use to separate string
'Ctl: Control to add separated string to

'Example...

'Dim x as integer
'Dim strray(1 to 5)

'strray(1) = "a1+a2+a3"
'strray(2) = "b1+b2+b3"
'strray(3) = "c1+c2+c3"
'strray(4) = "d1+d2+d3"
'strray(5) = "e1+e2+e3"

'For X = 1 To 5

'       String_SplitToControl strray(x), "+", list1

'Next x

Dim Split_Spot As Integer
Dim Split_Start As Integer

Split_Start% = 1

Do

    Split_Spot% = InStr(Split_Start%, Split_String$, Delimiter$) 'Find split characters
    
    If Split_Spot% = 0 Then Ctl.AddItem Mid(Split_String$, Split_Start%, Len(Split_String$)): Exit Sub 'If split characters not found exit sub
    
    Ctl.AddItem Mid(Split_String$, Split_Start%, Split_Spot% - Split_Start%)
    Split_Start% = Split_Spot% + Len(Delimiter$)
    
Loop

End Sub
Public Function String_TrimChar(strng As String, CharToTrim As String)

'Removes all chosen characters from a string

'Arguments...
'Strng: String to trim char
'CharToTrim: Character to trim from string

'Example...
'Dim Newstring as string
'NewString$ = String_TrimChar(mystring$, ".")

String_TrimChar = String_Replace(strng$, CharToTrim$, "", False, False)

End Function
Public Function Textbox_CharPosition(Textbx As TextBox)

'Returns character position on current line of a textbox
'Requires a timer to actually have some function

'Arguments...
'Textbx: Textbox to get the character position of cursor in

'Example...
'Put this in a timer
'lbl1.caption = TextBox_CharPosition(text1)

Dim Line As Integer
Dim i As Integer
Dim LineLen As Integer
Dim SelStrt As Integer

LineLen% = 0

Line% = TextBox_LinePosition(Textbx)

For i% = 1 To Line% - 1

    LineLen% = LineLen% + Len(String_GetLine(Textbx.Text, i%)) + 2

Next i%

Textbox_CharPosition = Textbx.SelStart - LineLen% + 1

End Function

Public Sub TextBox_Copy(Textbx As TextBox)

'Copies selected text to clipboard

'Arguments...
'Textbox to copy in

'Example...
'TextBox_Copy text1

Clipboard.SetText Textbx.SelText

End Sub

Public Sub TextBox_Cut(Textbx As TextBox)

'Copies selected text to clipboard
'Sets selected text to ""

'Arguments...
'Textbx: Textbox to cut in

'Example...
'TextBox_Cut text1

Clipboard.SetText Textbx.SelText
Textbx.SelText = ""

End Sub
Public Sub Textbox_OpenAsWebpage(Textbx As TextBox)

'Opens a textbox as html

TextBox_Save Textbx, "c:\preview.html"
Net_Webpage "c:\preview.html"

'Uncomment these to kill the temporary file
'Program_Pause 3
'Kill "c:\preview.html"

End Sub
Public Sub TextBox_Paste(Textbx As TextBox)

'Sets selected text in text to clipboards text

'Arguments...
'Textbox: Textbox to past in

'Example...
'TextBox_Paste text1

Textbx.SelText = Clipboard.GetText
Textbx.SetFocus

End Sub
Public Sub TextBox_SelectAll(Textbx As TextBox)

'Selects all contents of a textbox

'Arguments...
'Textbx: Textbox to selectall in

'Example...
'TextBox_SelectAll text1

Textbx.SelStart = 0
Textbx.SelLength = Len(Textbx)
Textbx.SetFocus

End Sub
Public Sub TextBox_Menu(Textbx As TextBox, frm As Form, mnu As Menu)

'If you right click a textbox the standard editing menu appears
'This allows you to replace it with a menu you have created

'Arguments...
'Textbx: Textbox to display menu in
'Frm: Form that has the textbox
'mnu: the menu you have created to display

'Example...

'In mousedown procedure of a textbox add this
'If button = 2 then
'Textbox_Menu text1, form1, mnufile

Textbx.enabled = False
Textbx.enabled = True
frm.PopupMenu mnu

End Sub
Public Sub Textbox_Undo(Textbx As TextBox)

'Undo function

'Arguments...
'Textbx: textbox to undo in

'Example...
'TextBox_Undo text1

Dim Undoit As Long

On Error Resume Next

Undoit& = SendMessage(Textbx.hwnd, EM_UNDO, 0&, 0&)

End Sub
Public Sub TextBox_Spell(Textbx As TextBox, tSpell As String, Speed As String)

'This will spell a string into a textbox with a defined speed

'Arguments:
'Textbx: Textbox to spell into
'tSpell: What to spell
'Speed: Speed described below

'Speeds:    [time in between in letter placed]
'       1[1.00 seconds]
'       2[0.90 seconds]
'       3[0.80 seconds]
'       4[0.70 seconds]
'       5[0.60 seconds]
'       6[0.50 seconds]
'       7[0.40 seconds]
'       8[0.30 seconds]
'       9[0.20 seconds]
'      10[0.10 seconds]

'Example: TextBox_Spell text1, mystory$, 8

Dim i As Integer
Dim speeda As Single
Dim spell As String

If Number_Valid(Speed$) = False Then Exit Sub
If Val(Speed$) < 1 Or Val(Speed$) > 10 Then MsgBox "Invalid Speed", 16, "Error..."

If Speed = 1 Then speeda! = 1 'define speeds
If Speed = 2 Then speeda! = 0.9
If Speed = 3 Then speeda! = 0.8
If Speed = 4 Then speeda! = 0.7
If Speed = 5 Then speeda! = 0.6
If Speed = 6 Then speeda! = 0.5
If Speed = 7 Then speeda! = 0.4
If Speed = 8 Then speeda! = 0.3
If Speed = 9 Then speeda! = 0.2
If Speed = 10 Then speeda! = 0.1

For i% = 1 To Len(tSpell$)

    spell$ = Mid(tSpell$, i%, 1) 'Set variable to letter
    Textbx.SelText = spell$ 'Add letter to textbox
    Program_Pause speeda! 'Timeout to  control thespeed

Next i%

End Sub
Public Sub TextBox_Find(Textbx As TextBox, FindWhat As String, Optional CaseSensitive As Boolean)

'Will highlight the first instance of a string in a textbox

'Arguments
'Textbx: Txtbox to find in
'FindWhat: What to find in the textbox
'CaseSensitive: Match case or not

'Example...
'TextBox_Find text1, "premier", true

Dim Length As Integer
Dim find As Integer
Dim rcount As Integer

On Error GoTo errorfix: 'If string not found goto label

Length% = Len(FindWhat$) 'Set variable to length of string to find
If CaseSensitive = True Then find% = InStr(Textbx.Text, FindWhat$) 'Find string [case sensitive]
If CaseSensitive = False Then find% = InStr(LCase$(Textbx.Text), LCase$(FindWhat$)) 'Find string
Textbx.SelStart = find% - 1 'Selstart to beginning of string to find
Textbx.SelLength = Length% 'Selength to length of string, find string is now selected
Textbx.SetFocus
Exit Sub

errorfix: MsgBox "Search text not found.", 64, "Done..."

End Sub
Public Sub TextBox_FindNext(Textbx As TextBox, Optional CaseSensitive As Boolean)

'This is a basic find next
'It will take what is highlighted in a textbox and find the next instance
'If no other instance is found it displays a messagebox

'Arguments...
'Textbx: Textbox to find next in
'CaseSensitive[default: false]: match case or not

'Example...
'Textbox_Findnext text1
 
Dim FoundEnd As Integer
Dim FoundString As String
Dim FoundLength As Integer
Dim Look As Integer

If Textbx.SelText = "" Then Exit Sub 'Exit sub if no text is selected
'If Textbx.SelText = "" Then MsgBox "No text selected", vbCritical, "Error..."

If CaseSensitive = True Then

        FoundEnd% = Textbx.SelStart + Textbx.SelLength 'Get place of last selected character
        FoundString$ = Textbx.SelText 'Set variable to selected text
        FoundLength% = Textbx.SelLength 'Set variable to selected text length
        Look% = InStr(FoundEnd%, Textbx, FoundString$) 'Find next instance
        
        If Look% = 0 Then MsgBox "The specified region has been searched", 16, "Finished...": Textbx.SetFocus: Exit Sub 'Check if done
        
        Textbx.SelStart = Look% - 1 'Set selstart to beginning of next instance
        Textbx.SelLength = FoundLength% 'Set sellength to length of string
        Textbx.SetFocus 'Set focus to finish
    
End If

If CaseSensitive = False Then

        FoundEnd% = Textbx.SelStart + Textbx.SelLength 'Get place of last selected character
        FoundString$ = Textbx.SelText 'Set variable to selected text
        FoundLength% = Textbx.SelLength 'Set variable to selected text length
        Look% = InStr(FoundEnd%, LCase$(Textbx), LCase$(FoundString$)) 'Find next instance
        
        If Look% = 0 Then MsgBox "The specified region has been searched", 16, "Finished...": Textbx.SetFocus: Exit Sub 'Check if done
        
        Textbx.SelStart = Look% - 1 'Set selstart to beginning of next instance
        Textbx.SelLength = FoundLength% 'Set sellength to length of string
        Textbx.SetFocus 'Setfocus to finish
    
End If

End Sub
Public Sub TextBox_Save(Textbx As TextBox, FullPath As String)

'Saves contents of a textbox

'Arguments...
'Textbx: textbox to save
'FullPath: Path to save file to

'Example...
'TextBox_Save text1, dir1.path & "\" & file1.file

Dim freenumber

File_CheckReadOnly FullPath$, 1 'If file is readonly fix that
If File_Validity(FullPath$, 2) = False Then Exit Sub 'Check for file existance

freenumber = FreeFile 'Set variable to free file

Open FullPath$ For Output As #freenumber

    Print #freenumber, Textbx.Text 'Print text to file

Close #freenumber

File_SetNormal FullPath$

End Sub
Public Sub TextBox_Replace(Textbx As TextBox, ReplaceWhat As String, ReplaceWith As String, Optional CaseSensitive As Boolean, Optional MessageBox As Boolean)

'Replaces all instances of a string in a textbox with another string

'Arguments...
'Textbx: Textbox to replace in
'ReplaceWhat: What to replace
'ReplaceWith: What to replace [ReplaceWhat] with
'CaseSensitive[default: false]: Match case or not
'Messagebox[default: true]: Display messagebox when complete giving number of replaced items

'Example...
'Textbox_ReplaceSelected(text1, "123", "abc")

'Replaces all instances of 123 with abc

Dim replacelength As Integer
Dim replacefind As Integer
Dim rcount As Integer

On Error GoTo errorfix: 'If string not found goto label

Do
    
    replacelength% = Len(ReplaceWhat$) 'Set variable to length of string to replace
    If CaseSensitive = True Then replacefind% = InStr(Textbx.Text, ReplaceWhat$) 'Find string to replace [case sensitive]
    If CaseSensitive = False Then replacefind% = InStr(LCase$(Textbx.Text), LCase$(ReplaceWhat$)) 'Find string to replace
    Textbx.SelStart = replacefind% - 1 'Selstart to beginning of string to replace
    Textbx.SelLength = replacelength% 'Selength to length of string, replace string is now selected
    Textbx.SelText = ReplaceWith$ 'Set selected string as string to replace with
        rcount% = rcount% + 1 'keep count of replacements

Loop Until InStr(Textbx.Text, ReplaceWhat$) = 0 'Loop until search string = 0

If MessageBox = True Then MsgBox "Search complete, " & rcount% & " replacements made.", vbInformation, "Search Complete..."

Exit Sub

errorfix: MsgBox "Search text not found.", 64, "Done..."

End Sub
Public Sub TextBox_ReplaceSelected(Textbx As TextBox, ReplaceWhat As String, ReplaceWith As String, Optional CaseSensitive As Boolean, Optional MessageBox As Boolean)

'Same as textbox_replace but this replaces only highlighted text

'Arguments...
'Textbx: Textbox to replace in
'ReplaceWhat: What to replace
'ReplaceWith: What to replace [ReplaceWhat] with
'CaseSensitive[default: false]: Match case or not
'Messagebox[default: true]: Display messagebox when complete giving number of replaced items

'Example...
'Textbox_ReplaceSelected(text1, "123", "abc")

'Replaces all instances in highlighted text of 123 with abc

Dim SeldText As String

'This is an optional line for error catching...uncomment to use
'if Textbx.SelText = "" then msgbox "No text selected" ,64,"Error..."

If CaseSensitive = "" Then CaseSensitive = False
If MessageBox = "" Then MessageBox = True

If String_Replace(Textbx.Text$, ReplaceWhat$, ReplaceWith$, CaseSensitive, MessageBox) = "" Then

    Exit Sub

Else

    SeldText$ = Textbx.SelText
    Textbx.SelText = String_Replace(SeldText$, ReplaceWhat$, ReplaceWith$, CaseSensitive, False)

End If

End Sub
Public Function TextBox_LinePosition(lpTextBox As TextBox)

'Returns the line of a textbox that the cursor is at
'Requires a timer

'Arguments...
'lpTextBox: The textbox you will be getting the line position from

'Example...
'info: Put this code in a timer

'label1.caption = TextBox_LinePosition(text1)

Dim i As Integer
Dim count As Integer
Dim Look As String

If lpTextBox.SelStart = 0 Then TextBox_LinePosition = 1: Exit Function
If lpTextBox.Text = "" Then Exit Function 'Exit if textbox is empty

count% = 1 'set count to one because it has to have atleast one line

For i% = 1 To lpTextBox.SelStart
    
    Look$ = Mid(lpTextBox.Text, i%, 1) 'Search each character
    If Look$ = Chr$(13) Then count% = count% + 1 'If character is return[chr$(13)] add 1 to line count

Next i%

TextBox_LinePosition = count%

End Function
Public Sub TextBox_Load(lTextBox As TextBox, fPath As String)

'Loads text into a text file

'Arguments...
'lTextBox: The textbox you wish to load the file into
'fPath: The full path of the file you are loading

'Example...
'info: Add a commondialog control to the form
'      This will load what ever file is chosen in the common dialog

'Commondialog1.showopen
'TextBox_Load text1, Commondialog1.filename

Dim DaText As String
Dim freenumber

If File_Validity(fPath$, 3) = False Then Exit Sub

freenumber = FreeFile 'Set variable to free file

Open fPath$ For Input As #freenumber 'Open file

    DaText$ = Input(LOF(freenumber), #freenumber) 'Set variable to each line of file

Close #freenumber

lTextBox.Text = DaText$

End Sub
Public Sub Window_Close(Winder As Long)

'Closes a given window

'Arguments...
'Winder: The window you want to close

'Example...
'info: This will close a message box

'Dim message as long
'message& = findwindow("#32770", vbnullstring)
'Window_Close message&

Call PostMessage(Winder&, WM_CLOSE, 0&, 0&)

End Sub
Public Sub Window_Hide(hwnd As Long)

'Hides a given window

'Arguments...
'Winder: The window you want to hide

'Example...
'info: This will hide notepad if it visible

'Dim pad as long
'pad& = findwindow("Notepad", vbnullstring)
'Window_Hide pad&

Call ShowWindow(hwnd&, SW_HIDE)

End Sub
Public Sub Window_Show(Winder As Long)
    
'Shows a given window

'Arguments...
'Winder: The window you want to show

'Example...
'info: This will show notepad if it has been hidden

'Dim pad as long
'pad& = findwindow("Notepad", vbnullstring)
'Window_Show pad&
    
Call ShowWindow(Winder&, SW_SHOW)

End Sub
Public Sub Window_SetText(Winder As Long, sString As String)

'Will set the text of an outside textbox to a given text

'Arguments...
'Winder: The window you wish to set the text to
'sString: The string you will set to the window

'Example...
'info: This sends "poop" to notepad

'Dim pad as long, edit as long
'pad& = FindWindow("Notepad", vbNullString)
'edit& = FindWindowEx(pad&, 0&, "Edit", vbNullString)
'Window_SetText edit&, "poop"

Dim DoIt As Long

On Error Resume Next

DoIt& = SendMessageByString(Winder&, WM_SETTEXT, 0, sString$)

End Sub
Public Sub Window_Enter(Winder As Long)

'Sends a key to a window

'Arguments...
'Winder: The handle of the window you want to send the key to

'Example...
'info: This just sends the enter kety to notepad, advancing to the next line

'dim pad as long
'pad& = findwindow("notepad", vbnullstring)
'Window_Enter pad&

'Call SendMessageByNumber(Winder, WM_CHAR, 13, 0&)

End Sub
Public Sub Window_Key(Winder As Long, TheKey As Integer)

'Sends a key to a window

'Arguments...
'Winder: The handle of the window you want to send the key to
'TheKey: The key you wish to send

'Example...
'info: This will advance to the next line in notepad

'dim pad as long
'pad& = findwindow("notepad", vbnullstring)
'Window_Key pad&, chr$(13)

'Call SendMessageByNum(Winder, WM_CHAR, TheKey%, 0&)

End Sub
Public Function Window_GetText(Winder As Long) As String

'Gets the text of a given window

'Arguments...
'Winder: The handle of the window you wish to get the text from

'Example...
'info: This will get the text from AOL

'Dim aol as long
'Dim aolcap as string
'aol& = findwindow("aol frame25", vbnullstring)
'aolcap& = Window_GetText(aol&)

Dim buf As String
Dim Length As Long

Length& = SendMessage(Winder&, WM_GETTEXTLENGTH, 0&, 0&)
buf$ = String(Length&, 0&)
Call SendMessageByString(Winder&, WM_GETTEXT, Length& + 1, buf$)

Window_GetText$ = buf$

End Function
