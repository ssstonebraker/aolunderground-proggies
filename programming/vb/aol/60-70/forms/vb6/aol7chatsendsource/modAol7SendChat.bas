Attribute VB_Name = "modAol7SendChat"
'Send Text To Chat In Aol 7.0 Example By Source
'Example Coded : October 22nd 2001 - 10:54 pm eastern time
'Contact Me On Aim: ciasource



'SendMessage API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'SendMessageLong API (clicking icon/button)
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'SendMessageByString (inserting text into chat textbox) via API
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'Find basic window (aol fram25)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'finds sub-windows (mdiclient, aolchild etc...)
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Const WM_SETTEXT = &HC       'sets text into chat textbox
Public Const WM_LBUTTONDOWN = &H201 'presses certain icon/button
Public Const WM_KEYUP = &H101       'key up (release) icon/button
Public Const VK_SPACE = &H20        'vk space (release)icon/button
Public Const WM_CHAR = &H102
Public Const ENTER_KEY = 13

Public Function SendChat(txtChat As String)
'usage:
'SendChat ("source's aol 7.0 SendChat example")
'or
'SendChat (""+text1.text+"")
'or
'SendChat (""+label1.caption+"")


'dims all variables as long for use in code
Dim RICHCNTL As Long, AOLChild As Long, MDIClient As Long
Dim AOLFrame As Long, i As Long, AOLIcon As Long
'defines AOLFrame, uses Public Declare FindWindow (AOL Frame25)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'defines MDIClient, uses Public Declare FindWindowEx
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'defines AolChild, uses Public Declare FindWindowEx, search for MDIClient
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
'defines RICHCNTL which is the Aol Chat TextBox
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)


'uses the Public Declare SendMessageByString and Const WM_SETTEXT
'and txtChat$ as the variable string, this inputs the text to textbox
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, txtChat$)
'defines the aol icon which will send the inserted text to chat.
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
'for i statement, finds the correct icon
For i& = 1& To 5&
    'defines again
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
'Pause (0.2)

'uses the Public Declare SendMessageLong to click AOLIcon&, with the
'const WM_LBUTTONDOWN (presses icon/button)
'Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
'uses the Public Declare SendMessageLong to click AOLIcon&, with the
'const WM_KEYUP and VK_SPACE, this (releases icon/button)
'Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

'update 10/24/01
'after realizing that click the Send icon wasnt full proof
'ie: if you minimized aol it would send right or were talking to ppl
'it wouldnt send write, so i used Public Const WM_CHAR and ENTER_KEY
'this instead of using api to click an icon, it sends it via keys
Call SendMessageLong(RICHCNTL&, WM_CHAR, ENTER_KEY, 0&)
End Function
