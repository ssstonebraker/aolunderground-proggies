Attribute VB_Name = "modAol7Keyword"
'Aol 7.0 Keyword Example By Source
'Visits a website via toolbar icon, keyword window textbox, then
'keyword window icon. All this is fully commented
'released : October 23rd, 2001 - 12:01 am. eastern time
'contact on aim: ciasource

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

Public Function Keyword(WebUrl As String)
Dim i As Long           'for i statement
Dim AOLIcon As Long     'dims aolicon
Dim AOLToolbar2 As Long 'dims second toolbar
Dim AOLToolbar As Long  'dims first toolbar
Dim AOLFrame As Long    'dims aol frame (aol frame 25)
Dim MDIClient As Long   'dims mdiclient (layer)
Dim AOLChild As Long    'dims aol child (keyword window)
Dim AOLChild2 As Long   'dims aol child 2 (altered)
Dim AOLIcon2 As Long    'dims aolicon2 (second icon, "go")
Dim AOLEdit As Long     'dims aoledit , which is aol's textbox's

'defines aolframe, using public declare FindWindow
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'defines aoltoolbar, uses FindWindowEx, and AOLFrame&
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
'defines aoltoolbar2, uses FineWindowEx, and first toolbar coding
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
'defines aolicon, uses FindWindowEx, and the second aol toolbar
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 26& 'for i statemenet
    'defines same aol icon again
    AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i& 'moves on to locate correct one
'uses public declare SendMessageLong, locates AOLICON&, then uses
'const WM_LBUTTONDOWN, to press the fixed button/icon
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
'uses public declare SendMessageLong, locates AOLICON&, then uses
'const WM_KEYUP and VK_SPACE, (which release the icon)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

'pause allows window to load (can be increased or decreased depending
'how computer speed and internet connection, .4/.5 is average
Pause 0.6
'the follow is for inserting the url into the keyword textbox
'defines MDIClient, uses FindWindowEx, and the AolFrame
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'defines AolChild with the use of FindWindowEx and MDIClient coding
'along with that it makes sure title is "Keyword" rather then the
'usual vbNullString
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Keyword")
'defines aoledit(textbox), uses FindWindowEx, AolChild
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
'From Public Declare Function SendMessageByString, it inserts text
'with use of AolEdit coding and const WM_SETTEXT and WebURL$ as the
'variable string
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, WebUrl$)

'this clicks the go icon..right have text is entered (sometimes
'you might want a tiny pause in here

'defines aolchild2, uses FindWindowEx, and mdiclient coding, and
'the title "Keyword"
AOLChild2& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Keyword")
'defines AolIcon2, which is located in AOLCHILD2, in the Keyword window
AOLIcon2& = FindWindowEx(AOLChild2&, 0&, "_AOL_Icon", vbNullString)
'SendMessageLong(for buttons/icons) from Public Declare then the
'const WM_LBUTTONDOWN
Call SendMessageLong(AOLIcon2&, WM_LBUTTONDOWN, 0&, 0&)
'SendMessageLong(for buttons/icons) from Public Declare then  the
'consts WM_KeyUp and VK_Space which release the pressed icon
Call SendMessageLong(AOLIcon2&, WM_KEYUP, VK_SPACE, 0&)

End Function
Public Function Pause(time As Long)
'pause for a certain amount of time
'Call pause(1)
Dim Current As Long 'dims current as long
Current = Timer     'sets current as timer
Do Until Timer - Current >= time 'greater then or equal to(time as long)
DoEvents            'doevents
Loop                'loopfor certain time
End Function
