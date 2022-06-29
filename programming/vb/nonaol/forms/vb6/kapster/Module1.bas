Attribute VB_Name = "Mod1"
Type MyLong
    Value As Long
    End Type


Type MyIP
    a As Byte
    b As Byte
    c As Byte
    d As Byte
    End Type
Global MyName As String
Global TheRoom As String
Global TheVictem As String
Global TheVicName As String
Global TheVicNamea As Integer
Global TheTour As String
Global TheHook As String
Global TheCount As Integer
Global TheImCount As Integer
Global MyPassy As String
Global Idiot As String
Global Mimic As Boolean
Global MimicR As Boolean
Global TheVicMime As String
Global HookTalker As String
Global TheMagic As Boolean
Global TheMagic2 As Boolean
Global RoomPart As String
Global SearchN As Boolean
Global PlayerJoin As Boolean
Global MyNameArray(0 To 1000) As String
Global Seeka As Boolean
Global bombSTRING As String
'llllll Java Flood
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GettopWindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function setfocusapi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function SetWindowPos Lib "user32" (ByVal H%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Public Const base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
   Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
       '***Part of the bonus code********************************


   Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
       '*********************************************************
'Global Const MF_BITMAP = 4
Public Const MF_BITMAP = &H4

Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long



Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const KEY_SNAPSHOT = &H2C
Global chatsendbutton%
Global gesturebutton%
Global chattextbox%
Global user$



Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const GW_CHILD = 5
Public Const Gw_hwndFirst = 0
Public Const gw_hwndlast = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const MOUSE_MOVED = &H1

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)


Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type
Function StringToHex(thestring)
Dim TheHex, Final As String
If Len(thestring) <> 4 Then Exit Function
For i = 1 To Len(thestring)
TheHex = Hex(Asc(Mid(thestring, i, 1)))
If Len(TheHex) = 1 Then TheHex = "0" & TheHex
Final = Final & TheHex
Next i
StringToHex = Final
End Function

Function HexToString(TheHex)
Dim thestring, Final As String
Dim TheLast As Integer
If Len(TheHex) <> 8 Then Exit Function
TheLast = 1
For i = 1 To 4
thestring = Chr(CByte("&H" & Mid(TheHex, TheLast, 2)))
Final = Final & thestring
TheLast = TheLast + 2
Next i
HexToString = Final
End Function

Function IPToString(IP)
Dim TheSection, Final As String
IP = IP & "."
a = InStr(1, IP, ".")
For i = 1 To 4
TheSection = Mid(IP, 1, a - 1)
Final = Final & Chr(TheSection)
IP = Right(IP, Len(IP) - a)
a = InStr(1, IP, ".")
Next i
IPToString = Final
End Function

Function Login()
Login = Chr(0) & Chr(0) & Chr(4)
End Function

Function RealLen(TheNum)
Dim TheLen As String
p = Hex(TheNum)
Select Case Len(p)
Case 1
TheLen = Chr(0) & Chr(CByte("&H" & p))
Case 2
TheLen = Chr(0) & Chr(CByte("&H" & p))
Case 3
TheLen = Chr(CByte("&H" & Left(p, 1))) & Chr(CByte("&H" & Right(p, 2)))
Case 4
TheLen = Chr(CByte("&H" & Left(p, 2))) & Chr(CByte("&H" & Right(p, 2)))
End Select
RealLen = TheLen
End Function



'writtEn By: -I-MoUsE-I-
'mAn wAs this A sonnA BitCh
Function DEC(strin)
     If strin = "00" Then a = 0
     If strin = "01" Then a = 1
     If strin = "02" Then a = 2
     If strin = "03" Then a = 3
     If strin = "04" Then a = 4
     If strin = "05" Then a = 5
     If strin = "06" Then a = 6
     If strin = "07" Then a = 7
     If strin = "08" Then a = 8
     If strin = "09" Then a = 9
     If strin = "0A" Then a = 10
     If strin = "0B" Then a = 11
     If strin = "0C" Then a = 12
     If strin = "0D" Then a = 13
     If strin = "0E" Then a = 14
     If strin = "0F" Then a = 15
     If strin = "10" Then a = 16
     If strin = "11" Then a = 17
     If strin = "12" Then a = 18
     If strin = "13" Then a = 19
     If strin = "14" Then a = 20
     If strin = "15" Then a = 21
     If strin = "16" Then a = 22
     If strin = "17" Then a = 23
     If strin = "18" Then a = 24
     If strin = "19" Then a = 25
     If strin = "1A" Then a = 26
     If strin = "1B" Then a = 27
     If strin = "1C" Then a = 28
     If strin = "1D" Then a = 29
     If strin = "1E" Then a = 30
     If strin = "1F" Then a = 31
     If strin = "20" Then a = 32
     If strin = "21" Then a = 33
     If strin = "22" Then a = 34
     If strin = "23" Then a = 35
     If strin = "24" Then a = 36
     If strin = "25" Then a = 37
     If strin = "26" Then a = 38
     If strin = "27" Then a = 39
     If strin = "28" Then a = 40
     If strin = "29" Then a = 41
     If strin = "2A" Then a = 42
     If strin = "2B" Then a = 43
     If strin = "2C" Then a = 44
     If strin = "2D" Then a = 45
     If strin = "2E" Then a = 46
     If strin = "2F" Then a = 47
     If strin = "30" Then a = 48
     If strin = "31" Then a = 49
     If strin = "32" Then a = 50
     If strin = "33" Then a = 51
     If strin = "34" Then a = 52
     If strin = "35" Then a = 53
     If strin = "36" Then a = 54
     If strin = "37" Then a = 55
     If strin = "38" Then a = 56
     If strin = "39" Then a = 57
     If strin = "3A" Then a = 58
     If strin = "3B" Then a = 59
     If strin = "3C" Then a = 60
     If strin = "3D" Then a = 61
     If strin = "3E" Then a = 62
     If strin = "3F" Then a = 63
     If strin = "40" Then a = 64
     If strin = "41" Then a = 65
     If strin = "42" Then a = 66
     If strin = "43" Then a = 67
     If strin = "44" Then a = 68
     If strin = "45" Then a = 69
     If strin = "46" Then a = 70
     If strin = "47" Then a = 71
     If strin = "48" Then a = 72
     If strin = "49" Then a = 73
     If strin = "4A" Then a = 74
     If strin = "4B" Then a = 75
     If strin = "4C" Then a = 76
     If strin = "4D" Then a = 77
     If strin = "4E" Then a = 78
     If strin = "4F" Then a = 79
     If strin = "50" Then a = 80
     If strin = "51" Then a = 81
     If strin = "52" Then a = 82
     If strin = "53" Then a = 83
     If strin = "54" Then a = 84
     If strin = "55" Then a = 85
     If strin = "56" Then a = 86
     If strin = "57" Then a = 87
     If strin = "58" Then a = 88
     If strin = "59" Then a = 89
     If strin = "5A" Then a = 90
     If strin = "5B" Then a = 91
     If strin = "5C" Then a = 92
     If strin = "5D" Then a = 93
     If strin = "5E" Then a = 94
     If strin = "5F" Then a = 95
     If strin = "60" Then a = 96
     If strin = "61" Then a = 97
     If strin = "62" Then a = 98
     If strin = "63" Then a = 99
     If strin = "64" Then a = 100
     If strin = "65" Then a = 101
     If strin = "66" Then a = 102
     If strin = "67" Then a = 103
     If strin = "68" Then a = 104
     If strin = "69" Then a = 105
     If strin = "6A" Then a = 106
     If strin = "6B" Then a = 107
     If strin = "6C" Then a = 108
     If strin = "6D" Then a = 109
     If strin = "6E" Then a = 110
     If strin = "6F" Then a = 111
     If strin = "70" Then a = 112
     If strin = "71" Then a = 113
     If strin = "72" Then a = 114
     If strin = "73" Then a = 115
     If strin = "74" Then a = 116
     If strin = "75" Then a = 117
     If strin = "76" Then a = 118
     If strin = "77" Then a = 119
     If strin = "78" Then a = 120
     If strin = "79" Then a = 121
     If strin = "7A" Then a = 122
     If strin = "7B" Then a = 123
     If strin = "7C" Then a = 124
     If strin = "7D" Then a = 125
     If strin = "7E" Then a = 126
     If strin = "7F" Then a = 127
     If strin = "80" Then a = 128
     If strin = "81" Then a = 129
     If strin = "82" Then a = 130
     If strin = "83" Then a = 131
     If strin = "84" Then a = 132
     If strin = "85" Then a = 133
     If strin = "86" Then a = 134
     If strin = "87" Then a = 135
     If strin = "88" Then a = 136
     If strin = "89" Then a = 137
     If strin = "8A" Then a = 138
     If strin = "8B" Then a = 139
     If strin = "8C" Then a = 140
     If strin = "8D" Then a = 141
     If strin = "8E" Then a = 142
     If strin = "8F" Then a = 143
     If strin = "90" Then a = 144
     If strin = "91" Then a = 145
     If strin = "92" Then a = 146
     If strin = "93" Then a = 147
     If strin = "94" Then a = 148
     If strin = "95" Then a = 149
     If strin = "96" Then a = 150
     If strin = "97" Then a = 151
     If strin = "98" Then a = 152
     If strin = "99" Then a = 153
     If strin = "9A" Then a = 154
     If strin = "9B" Then a = 155
     If strin = "9C" Then a = 156
     If strin = "9D" Then a = 157
     If strin = "9E" Then a = 158
     If strin = "9F" Then a = 159
     If strin = "A0" Then a = 160
     If strin = "A1" Then a = 161
     If strin = "A2" Then a = 162
     If strin = "A3" Then a = 163
     If strin = "A4" Then a = 164
     If strin = "A5" Then a = 165
     If strin = "A6" Then a = 166
     If strin = "A7" Then a = 167
     If strin = "A8" Then a = 168
     If strin = "A9" Then a = 169
     If strin = "AA" Then a = 170
     If strin = "AB" Then a = 171
     If strin = "AC" Then a = 172
     If strin = "AD" Then a = 173
     If strin = "AE" Then a = 174
     If strin = "AF" Then a = 175
     If strin = "B0" Then a = 176
     If strin = "B1" Then a = 177
     If strin = "B2" Then a = 178
     If strin = "B3" Then a = 179
     If strin = "B4" Then a = 180
     If strin = "B5" Then a = 181
     If strin = "B6" Then a = 182
     If strin = "B7" Then a = 183
     If strin = "B8" Then a = 184
     If strin = "B9" Then a = 185
     If strin = "BA" Then a = 186
     If strin = "BB" Then a = 187
     If strin = "BC" Then a = 188
     If strin = "BD" Then a = 189
     If strin = "BE" Then a = 190
     If strin = "BF" Then a = 191
     If strin = "C0" Then a = 192
     If strin = "C1" Then a = 193
     If strin = "C2" Then a = 194
     If strin = "C3" Then a = 195
     If strin = "C4" Then a = 196
     If strin = "C5" Then a = 197
     If strin = "C6" Then a = 198
     If strin = "C7" Then a = 199
     If strin = "C8" Then a = 200
     If strin = "C9" Then a = 201
     If strin = "CA" Then a = 202
     If strin = "CB" Then a = 203
     If strin = "CC" Then a = 204
     If strin = "CD" Then a = 205
     If strin = "CE" Then a = 206
     If strin = "CF" Then a = 207
     If strin = "D0" Then a = 208
     If strin = "D1" Then a = 209
     If strin = "D2" Then a = 210
     If strin = "D3" Then a = 211
     If strin = "D4" Then a = 212
     If strin = "D5" Then a = 213
     If strin = "D6" Then a = 214
     If strin = "D7" Then a = 215
     If strin = "D8" Then a = 216
     If strin = "D9" Then a = 217
     If strin = "DA" Then a = 218
     If strin = "DB" Then a = 219
     If strin = "DC" Then a = 220
     If strin = "DD" Then a = 221
     If strin = "DE" Then a = 222
     If strin = "DF" Then a = 223
     If strin = "E0" Then a = 224
     If strin = "E1" Then a = 225
     If strin = "E2" Then a = 226
     If strin = "E3" Then a = 227
     If strin = "E4" Then a = 228
     If strin = "E5" Then a = 229
     If strin = "E6" Then a = 230
     If strin = "E7" Then a = 231
     If strin = "E8" Then a = 232
     If strin = "E9" Then a = 233
     If strin = "EA" Then a = 234
     If strin = "EB" Then a = 235
     If strin = "EC" Then a = 236
     If strin = "ED" Then a = 237
     If strin = "EE" Then a = 238
     If strin = "EF" Then a = 239
     If strin = "F0" Then a = 240
     If strin = "F1" Then a = 241
     If strin = "F2" Then a = 242
     If strin = "F3" Then a = 243
     If strin = "F4" Then a = 244
     If strin = "F5" Then a = 245
     If strin = "F6" Then a = 246
     If strin = "F7" Then a = 247
     If strin = "F8" Then a = 248
     If strin = "F9" Then a = 249
     If strin = "FA" Then a = 250
     If strin = "FB" Then a = 251
     If strin = "FC" Then a = 252
     If strin = "FD" Then a = 253
     If strin = "FE" Then a = 254
     If strin = "FF" Then a = 255
     DEC = a
End Function
Public Function DBL_Mod(ByVal N1 As Double, ByVal N2 As Double) As Double
    DBL_Mod = CDbl(N1 - (DBL_Divide(N1, N2)) * N2)
End Function

Public Function DBL_Divide(ByVal N1 As Double, ByVal N2 As Double) As Double
    DBL_Divide = Int(N1 / N2)
End Function

Public Function DEC_HEX(ByVal Number As Double) As String
    Dim i As Long, j As String, s As String
    Do
        j = Trim(CStr(DBL_Mod(Val(CStr(Number)), 16)))
        
        If j > 9 Then
            j = Chr((Val(j)) + 55)
        End If
        
        Number = DBL_Divide(Number, 16)
        s = Trim(j) & s
    Loop Until Number = 0
    
    DEC_HEX = CStr(s)
    
End Function

Function AsciiToHex(strin)
'this was written By: -I-MoUsE-I-!
    Dim NewSTrin As String
    
    Do Until strin = ""
        X = Hex(AscB(Left(strin, 1)))
        
        If Len(TrimSpaces(X)) = 2 Then
            NewSTrin = NewSTrin & X
        Else
            NewSTrin = NewSTrin & "0" & X
        End If
        
        strin = Right(strin, Len(strin) - 1)
    Loop
    
    AsciiToHex = NewSTrin
    
End Function

Function AsciiToHex2(strin As String)
'this was written By: -I-MoUsE-I-!
    Dim NewSTrin As String
    
    Do Until strin = ""
        X = Hex(AscB(Left(strin, 1)))
        
        If Len(TrimSpaces(X)) = 2 Then
            NewSTrin = NewSTrin & X & " "
        Else
            NewSTrin = NewSTrin & "0" & X & " "
        End If
        
        strin = Right(strin, Len(strin) - 1)
    Loop
    
    AsciiToHex2 = NewSTrin
    
End Function

Function Hex_Dec(Hex_val As String) As Variant


    Dim lstr
    Dim X, Y
    Dim ch_in
    Dim conv_temp As Variant
    lstr = Len(Hex_val)
    For X = 0 To lstr - 1
        Y = lstr - X
        ch_in = Mid$(Hex_val, Y, 1)
        If Asc(ch_in) >= 48 And Asc(ch_in) <= 57 Then
            ch_in = ch_in
        ElseIf Asc(ch_in) >= 65 And Asc(ch_in) <= 70 Then
            ch_in = Asc(ch_in) - 55
        ElseIf Asc(ch_in) >= 97 And Asc(ch_in) <= 102 Then
            ch_in = Asc(ch_in) - 87
        End If


        Hex_Dec = Hex_Dec + 16 ^ X * ch_in
    Next X


End Function

Function Hex_Dec2(Hex_val As String) As Variant


    Dim lstr
    Dim X, Y
    Dim ch_in
    Dim conv_temp As Variant
    lstr = Len(Hex_val)
    For X = 0 To lstr - 1
        Y = lstr - X
        ch_in = Mid$(Hex_val, Y, 1)
        If Asc(ch_in) >= 48 And Asc(ch_in) <= 57 Then
            ch_in = ch_in
        ElseIf Asc(ch_in) >= 65 And Asc(ch_in) <= 70 Then
            ch_in = Asc(ch_in) - 55
        ElseIf Asc(ch_in) >= 97 And Asc(ch_in) <= 102 Then
            ch_in = Asc(ch_in) - 87
        End If


        Hex_Dec2 = Hex_Dec2 + 16 ^ X * ch_in
        Hex_Dec2 = Hex_Dec2 & " "
    Next X


End Function

Function TrimSpaces(Text)


    If InStr(Text, " ") = 0 Then
        TrimSpaces = Text
        Exit Function
    End If

    For trimspace = 1 To Len(Text)
        thechar$ = Mid(Text, trimspace, 1)
        thechars$ = thechars$ & thechar$

        If thechar$ = " " Then
            thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
        End If
    Next trimspace

    TrimSpaces = thechars$
End Function

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function Findchildbytitle(parentw, childhand)
    firs% = GetWindow(parentw, 5)
    If UCase(GetCaption(firs%)) Like UCase(childhand) Then
        GoTo bone
    End If
    firs% = GetWindow(parentw, GW_CHILD)

    While firs%
        firss% = GetWindow(parentw, 5)
        If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then
            GoTo bone
        End If
        firs% = GetWindow(firs%, 2)
        If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then
            GoTo bone
        End If
        Wend
    Findchildbytitle = 0
bone:
    Room% = firs%
    Findchildbytitle = Room%
End Function

Function VPGetText(child)
'Get the text of a window
    gettrim = sendmessagebynum(child, 14, 0&, 0&)
    trimspace$ = Space$(gettrim)
    getstrin = SendMessageByString(child, 13, gettrim + 1, trimspace$)

    VPGetText = trimspace$
End Function

Sub StayOnTop(Frm As Form)
    On Error GoTo don
    Dim success%
    success% = SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
don:
End Sub

Sub Pause(interval)
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
   FileExist = (Dir(Fname) <> "")
End Function
Function VPWindow()
    VP% = FindWindow("VPFrame", vbNullString)
    VPWindow = VP%
End Function
Function base64_encode(DecryptedText) As String
Dim c1, c2, c3 As Integer
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim retry As String
   For n = 1 To Len(DecryptedText) Step 3
      c1 = Asc(Mid$(DecryptedText, n, 1))
      c2 = Asc(Mid$(DecryptedText, n + 1, 1) + Chr$(0))
      c3 = Asc(Mid$(DecryptedText, n + 2, 1) + Chr$(0))
      w1 = Int(c1 / 4)
      w2 = (c1 And 3) * 16 + Int(c2 / 16)
      If Len(DecryptedText) >= n + 1 Then w3 = (c2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
      If Len(DecryptedText) >= n + 2 Then w4 = c3 And 63 Else w4 = -1
      retry = retry + mimeencode(w1) + mimeencode(w2) + mimeencode(w3) + mimeencode(w4)
   Next
   base64_encode = retry
End Function

Function base64_decode(a) As String
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim retry As String

   For n = 1 To Len(a) Step 4
      w1 = mimedecode(Mid$(a, n, 1))
      w2 = mimedecode(Mid$(a, n + 1, 1))
      w3 = mimedecode(Mid$(a, n + 2, 1))
      w4 = mimedecode(Mid$(a, n + 3, 1))
      If w2 >= 0 Then retry = retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
      If w3 >= 0 Then retry = retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
      If w4 >= 0 Then retry = retry + Chr$(((w3 * 64 + w4) And 255))
   Next
   base64_decode = retry
End Function

Private Function mimeencode(w As Integer) As String
   If w >= 0 Then mimeencode = Mid$(base64, w + 1, 1) Else mimeencode = ""
End Function

Private Function mimedecode(a As String) As Integer
   If Len(a) = 0 Then mimedecode = -1: Exit Function
   mimedecode = InStr(base64, a) - 1
End Function

Public Sub PlaySound(strFileName As String)
    sndPlaySound strFileName, 1
End Sub

Function FileExista(Fname As String) As Boolean
    On Local Error Resume Next
   FileExista = (Dir(Fname) <> "")
End Function

Function Wave_Lenght(Dateiname)
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("stop sound", RS, 128, cb)
    i = mciSendString("close sound", RS, 128, cb)
    i = mciSendString("open waveaudio!" & Dateiname & " Alias sound", RS, 128, cb)
    i = mciSendString("status sound length", RS, 128, cb)
    Wave_Lenght = RS
    i = mciSendString("stop sound", RS, 128, cb)
   
    i = mciSendString("close sound", RS, 128, cb)
End Function

Public Sub PlayMouseSound(MouseSoundPath As String)
    Dim i As Long, RS As String, cb As Long
    RS = Space$(128)
    i = mciSendString("open waveaudio!" & MouseSoundPath & " Alias sound", RS, 128, cb)
    i = mciSendString("play sound", RS, 128, cb)
End Sub
Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function
Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function



Function VPGetUser()
'Get the user name of the person using VP
hwndz% = FindWindow(vbNullString, "My Identity")
If hwndz% = 0 Then
If GetCaption(VPWindow) = vbNullString Then Exit Function
AppActivate "vplaces"
SendKeys "%AE", True
hwndz% = FindWindow(vbNullString, "My Identity")
End If
id1% = Findchildbytitle(hwndz%, "Basic Info")
firs% = GetWindow(id1%, GW_CHILD)
VPGetUser = VPGetText(firs%)
hwndz2% = Findchildbytitle(hwndz%, "Cancel")
VPbutton (hwndz2%)
VPbutton (hwndz2%)
End Function

Public Sub VPbutton(but%)
'Click on the button
clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub OpenURL(lol)
ShellExecute hwnd, "open", lol, vbNullString, vbNullString, SW_SHOWMAXIMIZED
End Sub

Function VPbackwards(strin)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
VPbackwards = newsent$

End Function

Function IPToStrin(Value As Double) As String
    Dim l As MyLong
    Dim i As MyIP
    l.Value = DoubleToLong(Value)
    LSet i = l
    IPToStrin = i.a & "." & i.b & "." & i.c & "." & i.d
End Function


Function DoubleToLong(Value As Double) As Long


    If Value <= 2147483647 Then


        DoubleToLong = Value
        Else


            DoubleToLong = -(4294967296# - Value)
            End If
End Function


