Attribute VB_Name = "ModAol7SendIM"
'Aol Version 7.0 Send Instant Message Example By Source
'Published: Saturday, October 20th, 2001 - 11:37:05 eastern time
'This is fully commented and aimed to help you better understand
'Public Declare Functions and Public Const, something that isnt usually understood
'just copied.
'website that might be up but might not : www.dazedworld.com
'contact aim: ciasource

'FindWindow
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'FindWindowEx
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'SendMessageLong(usually to click something)
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'SendMessageByString(usally to send text to something)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Public Const WM_LBUTTONDOWN = &H201 ' Const to push button/icon down
Public Const WM_LBUTTONUP = &H202   ' Const to release button/icon
Public Const WM_KEYUP = &H101       ' Const to key up (another why of lbuttonup)
Public Const VK_SPACE = &H20        ' Const space key (used to release icon)
Public Const WM_SETTEXT = &HC       ' Const to SetText/Send Text to a Textbox
Public Function SendInstantMessage(txtsn As String, txtmsg As String)
'txtsn as string, txtmsg as string lets you use it in the form coding

Dim i As Long 'When theres icons underneath each other they use for i
Dim x As Long 'Same as above but i split up variables for different processes
Dim AOLIcon As Long 'AolIcon sets the Icon on the Toolbar
Dim AolToolBar2 As Long 'Top Tool Bar
Dim AolToolBar As Long  'Second Tool Bar
Dim AOLFrame As Long    'AolFrame (aol frame25) enables location of everything else
Dim AOLEdit As Long     'Aol Edit
Dim AOLChild As Long    'Aol's Child
Dim MDIClient As Long   'Aol's MDI Client layer
Dim AolIcon2 As Long    'AolIcon2 sets the Send Button Icon
Dim RICHCNTL As Long    'This is RichTextBox, for the message/body of coding

'defines the aol frame, uses Public Declare Function FindWindow
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'defines AolToolBar, uses Public Declare Function FindWindowEx,(sub-layers)
AolToolBar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
'defines AolToolBar2, uses Public Declare Function FindWindowEx,(sub-layers)
AolToolBar2& = FindWindowEx(AolToolBar&, 0&, "_AOL_Toolbar", vbNullString)
'defines AolIcon, uses FindWindowEx, _aol_icon, and then the for to statements
AOLIcon& = FindWindowEx(AolToolBar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4& 'goes on the tool bar over a certain numbers of icons
    'defines AolIcon inside the for i statement
    AOLIcon& = FindWindowEx(AolToolBar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i& ' Next Icon...then moves to right one
'SendMessageLong, to the AolIcon, Uses Const WM_LBUTTONDOWN
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
'SendMessageLong, to the AolIcon, Uses Const WM_KEYUP, VK_SPACE
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
'uses the Public Function to pause for one/fourth(1/4) of a second
'depending on your web connection and how fast your computer is you might
'have to increase that number to allow time for the IM window to load. If
'you have a fast computer you can lower it and coding will move along faster
'but there will be an increase of the risk that it could skip over and not
'fill in the text.
Pause 0.4
'Sets MDICLient (Public Declare Function FindWindowEx)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'Sets AolChild (Public Declare Function FindWindowEx)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
'Sets AolEdit (Public Declare Function FindWindowEx)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
'SendMessageByString(text) for AolEdit, using Const WM_SETTEXT, then txtsn$ as the string
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, txtsn$)
'RICHCNTL is defined using FindWindowEx
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
'The message/body of an im isn't just a textbox its a RichTextBox (allows
'the use of colors..underline etc..
'SendMessageByString, using the Const WM_SETTEXT, and txtmsg$ as string
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, txtmsg$)

'AolIcon2 is defined using Public Declare Function FindWindowEx
AolIcon2& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
'the x is like the i but different for the seperate coding

For x& = 1& To 9&
    'Uses AolIcon2&=, Finds for AolChild and the icon itself
    AolIcon2& = FindWindowEx(AOLChild&, AolIcon2&, "_AOL_Icon", vbNullString)
'next x& , this goes to the correct icon the the aol frame.
Next x&
'SendMessageLong on AolIcon2, using Const WM_LBUTTONDOWN
Call SendMessageLong(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
'SendMessageLong on AolIcon2, using Const WM_KEYUP and VK_SPACE
Call SendMessageLong(AolIcon2&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function Pause(time As Long)
'pause for a certain amount of time
'Call pause(1)
'or
'Pause 0.5
Dim Current As Long 'Dims current as long
Current = Timer 'sets current as timer
Do Until Timer - Current >= time 'do until timer - current is greater than/equal
DoEvents 'DoEvents Statement
Loop 'Loop for the time then continue with coding.
End Function
