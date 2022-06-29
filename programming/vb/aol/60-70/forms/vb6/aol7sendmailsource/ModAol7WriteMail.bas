Attribute VB_Name = "ModAol7WriteMail"
'SendMessage
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Find basic window (aolframe25)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Find more advanced window/layering (toolbar)
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'sendmessagelong usually for clicking icons/buttons
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'sendmessagebystring usually for sending text.
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const WM_LBUTTONDOWN = &H201 'pressing button down
Public Const WM_KEYUP = &H101       'key up
Public Const VK_SPACE = &H20        'space bar
Public Const WM_SETTEXT = &HC       'set/send text


Public Function WriteMail(txtTo As String, txtSub As String, txtMsg As String)
Dim i As Long '}
Dim j As Long '}  i, j, x are all used for the For i/j/x statements
Dim x As Long '}
Dim AOLIcon As Long    'icon on toolbar
Dim AOLIcon2 As Long   'icon in write mail (send)
Dim AOLIcon3 As Long   'icon on comfirmation modal
Dim AOLToolbar2 As Long 'top toolbar
Dim AOLToolbar As Long  'bottom toolbar
Dim AOLFrame As Long   'general aolframe (aolframe25)
Dim AOLEdit As Long    'aoledit is textbox related
Dim AOLEdit2 As Long   'aoledit2 is textbox related
Dim AOLChild As Long   'aol child
Dim MDIClient As Long  'mdiclient layer
Dim RICHCNTL As Long   'rich text control box
Dim AOLModal As Long   'aol pop-up comfirmation modal.
'
'CLICKS ICON ON TOOLBAR
'defines the aol frame using Public Declare  FindWindow
AOLFrame& = FindWindow("AOL Frame25", vbNullString)

'defines AOLToolbar using Public Declare FindWindowEx
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)

'defines AOLToolbar2 using Public Declare FindWindowEx
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)

'defines AolIcon using Public Declare FindWindowEx, in toolbar2
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)

'for i statement (related to locating icons)
For i& = 1& To 2&
    'same as before but inside the For I Statement.
    AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)

'next i (locates and finds the icon)
Next i&

'uses Public Declare SendMessageLong To click icons/buttons, then Const WM_LBUTTONDOWN
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)

'uses Public Declare SendMessageLong To click icons/buttons, then the Const WM_KEYUP & VK_SPACE
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
'
'FILLS IN TEXT FOR THE "MAIL TO:" SECTION
'Gives a pause of half a second to let the window load...this can be changed depending
'on how good/fast your computer is.
Pause 0.5

'defines MDIClient using Public Declare FindWindowEx, and uses AOLframe
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)

'defines AOLChild using Public Declare FindWindowEx, and MDIClient, and looks for caption "Write Mail"
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")

'defines AOLEdit using Public Declare FindWindowEx, and any Caption
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)

'SendMessageByString in Public Declars, uses Const WM_SETTEXT
'also uses txtto$ as the string which is implanted
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, txtTo$)
'
'FILLS IN TEXT FOR THE "SUBJECT:" SECTION
'for j statment (same as For I)
For j& = 1& To 3&

'sets AolEdit2 using FindWindowEx, AolEdit2 itself and AolChild
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)

'next j, find the right spot and fill in text.
Next j&
'Public Declare SendMessageByString, with aoledit2 and const WM_SETTEXT, and txtsub$ as the string.
Call SendMessageByString(AOLEdit2&, WM_SETTEXT, 0&, txtSub$)
'
'FILLS IN THE BODY/MESSAGE OF THE EMAIL

'defines RICHCNTL(textbox), using FindWindowEx from Pub Decl.
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)

'SendMessageByString from Public Declars, and WM_SETTEXT for const, txtmsg$ as string
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, txtMsg$)
'
Pause 0.2 'pause 2/10ths a second before hitting send
AOLIcon2& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)

'for x statement (same as for i)
For x& = 1& To 17&
    'defines aolicon2
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next x&

'SendMessageLong using AolIcon2, and const WM_LBUTTONDOWN
Call SendMessageLong(AOLIcon2&, WM_LBUTTONDOWN, 0&, 0&)

'SendMessageLong using aol icon2, and const (WM_KEYUP & VK_SPACE)
Call SendMessageLong(AOLIcon2&, WM_KEYUP, VK_SPACE, 0&)

Pause 0.6 ' pause a 6th of a second while mail sends and when the
'comfirmation modal comes up hopefully this will click it in time.

'defines AOLModal using findwindow
AOLModal& = FindWindow("_AOL_Modal", vbNullString)

'3rd icon is the one inside the modal which will be clicked
AOLIcon3& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)

'Public Declare SendMessageLong and const's WM_LBUTTONDOWN, WM_KEYUP, and VK_Space
Call SendMessageLong(AOLIcon3&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon3&, WM_KEYUP, VK_SPACE, 0&)
Call SendMessageLong(AOLIcon3&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon3&, WM_KEYUP, VK_SPACE, 0&)
'was having problems clicking last icon, so i doubled coding for
'pressing down and up..works fine now
End Function
Public Function Pause(time As Long)
'pause for a certain amount of time
'Call pause(1)
'Pause 0.4
Dim Current As Long
Current = Timer
Do Until Timer - Current >= time
DoEvents
Loop
End Function
