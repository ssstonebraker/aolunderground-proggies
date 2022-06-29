Attribute VB_Name = "y0shi"
'Name:           Yoshi.bas
'Creator:        YoSHi
'E-Mail At:      Yoshiii@aol.com (Only mail me for help, or to report bugs)
'Status:         Not Close to completion
'Used for:       Use for participation in YoSHi Productions
'Also Used for:  People who are lazy and need explanations.
'Credits:        YoSHi, VïØ, Volter, Wgf
'Extra Code:     This section is the code which can't be put in a function
'                because it requires special parts which not all programs
'                have and in result, causes an error. Here are those codes:

'                1)Printing data from a RichTextBox - I know this code
'                bothers everyone, so here is an example:

'                Using the SelPrint Property of a RichTextBox, we send
'                the data to the Printer. The hDC property is just a way
'                of telling the data to go to the printers handle.

'                RichTextBox1.SelPrint (Printer.hDC)

'API Tutorial:   FindWindow(ParentClass, ParentWindowCaption)
'                FindWindowEx(ParentWindow, SameTimeOfClass, WindowClass, WindowCaption)
                 
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wflags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Public Const WM_CLOSE = &H10
Public Const WM_SETTEXT = &HC
Public Const WM_CHAR = &H102
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5

Public Const TAB_KEY = &H9
Public Const SPACE_BAR = &H20
Public Const ENTER_KEY = 13

Function StringWriteToFile(LocationOnSystem As String, TextToInput As String)
'This sub creates a file on the system in the location stored under
'LoactionOnSystem, and fills it with the text stored under 'TextToInput'
'Example:

'Call StringWriteToFile("C:\TestingFile.txt", "Testing 123")

'or you could input data from a textbox. Example:

'Call StringWriteToFile("C:\TestingFile.txt", Text1.Text)

'Code:

'Step 1: Open the location for output, and name it #1, as the first.

'Note: If you just want to add to that file on the system, instead of
'totally deleting it and inputing the data stored under TextToInput,
'Then make these changes:

'Instead Of: Open LocationOnSystem For Output As #1
'Make    It: Open LocationOnSystem For Append As #1

Open LocationOnSystem$ For Output As #1

'Write the text stored in the TextToInput variable in the opened location.
Write #1, TextToInput$

'Close the file since we do not need it anymore.
Close #1
End Function

Function ProgramOpenInBrowser(URL As String)
'This sub opens the target URL in your default browser. The URL is
'saved in the URL variable. Example:

'Call ProgramOpenInBrowser("http://microsoft.com")

'Code:

'Explorer runs HTML or webpages automatically by default. By using
'Shell, we can tell Explorer to start the webpage, or by default,
'open it in a browser.
Shell ("Start " & URL$)
End Function

Function ProgramRun(TargetProgram As String)
'This is a very simple code that runs the program. Example:

'Call ProgramRun("C:\Command.com")

'Code:

'Tells program to run the program stored under the variable TargetProgram.
Shell TargetProgram$
End Function

Function StringRemoveSpaces(TheString As String)
'This sub removes all the spaces from the desired string. Example:

'Text1.Text = StringRemoveSpaces(Text1.Text)

'This will make the text in Text1.Text equal Text1.Text after all
'spaces have been removed.

'Code:

'If the program finds no spaces in TheString (returns of a value of 0),
'then make the program keep TheString the same as it was before since
'there are no spaces to remove.
If InStr(TheString$, " ") = 0 Then
    StringRemoveSpaces = TheString$
    'Go to point 'SkipRest' so the program doesnt go through the
    'following code.
    GoTo SkipRest
End If

'Now that we know there are spaces, make the function do the
'following steps from the beggining to the end of 'TheString'.
'The Len function gets the length of the strength from the
'beggining to the end.
For StringRemoveSpaces = 1 To Len(TheString$)
    'Make the SearchString variable equal the point in 'TheString'
    'From which RemoveSpaces starts, and going up 1 more point.
    'This will eventually lead to searching of the whole string for
    'spaces. Notice the code we wrote before that mentions 1 to the
    'end of the string. It was all in 'StringRemoveSpaces', which is what
    'we are starting from and going up by 1.
    SearchString$ = Mid(TheString$, StringRemoveSpaces, 1)
    'Store the SearchString variable in 'BlankSearch', but dont remove
    'the previous data in BlankSearch because that may contain spaces
    'to be removed.
    BlankSearch$ = BlankSearch$ & SearchString$
    'If the SearchString variable from before was just a blank space
    'then...
    If SearchString$ = " " Then
        'Then search 'BlankSearch' from the beggining to the end,
        'minusing 1 each time so we are sure that the space is
        'removed.
        BlankSearch$ = Mid(BlankSearch$, 1, Len(BlankSearch$) - 1)
    End If
'Begin the next series in StringRemoveSpaces until, as we wrote on the first
'line (For StringRemoveSpaces = 1 To Len(TheString)), we reach the end of the
'string (TheString).
Next StringRemoveSpaces

'Once the entire loop is complete, make the StringRemoveSpaces function equal
'our finished product without any spaces, 'BlankSearch'.
StringRemoveSpaces = BlankSearch$

'The skipping point for where the program would go if there were no
'spaces.
SkipRest:
End Function

Function StringKillListDuplicates(TheListBox As ListBox)
'This is a sub that removes all duplicate items in a listbox. Example:

'Call StringKillListDuplicates(List1)

'This will call the RemoveListDuplicates function for TheListBox which
'is in this case, List1.

'Code:

'Since 0 is counted as the first item in a listbox, we have to minus
'1 so that we are not including an extra item in the list.
For ListItems = 0 To TheListBox.ListCount - 1
    'The part of the next code that says TheListBox.List brings up
    'the name of any part of a listbox. The part of the listbox is
    'currently named 'ListItems', since we named it that in the first
    'line of code. Store all this data in the CurrentItem variable.
    CurrentItem = TheListBox.List(ListItems)
    'Search the list again to search for a duplicate of the
    'CurrentItem variable.
    For TheDuplicate = 0 To TheListBox.ListCount - 1
        'Name each item in the rest of the listbox 'NewerItem'.
        NewerItem = TheListBox.List(TheDuplicate)
        'If 'TheDuplicate' equals 'ListItems', then we know
        'that it is not a dupe, just the same item again.
        'We dont do anything in response because there is no
        'need for duplicate removal. We add the GoTo
        'NotADuplicate so the program skips the next step.
        If TheDuplicate = ListItems Then GoTo NotADuplicate
        'If 'NewerItem' equals 'CurrentItem', then it is ovious that there is
        'a duplicate.
        'To handle with the duplicate, we tell the listbox to remove the
        'duplicate which is stored under 'TheDuplicate'. We remove it with
        'the RemoveItem property of a listbox. We also use the RemoveSpaces
        'function and the LCase statement because they make sure the spaces
        'are removed and that it doesn't matter what case the text is in.
        If StringRemoveSpaces(LCase(NewerItem)) = StringRemoveSpaces(LCase(CurrentItem)) Then
        TheListBox.RemoveItem (TheDuplicate)
        End If

'The skipping point for the program if there is no duplicate.
NotADuplicate:
    
    'Repeat this entire statement, but for the next item in the list. We keep
    'repeating the statement until we search the entire listbox.
    Next TheDuplicate

'Now repeat the whole function again, but for the next item in the list.
'The program will now repeat the whole thing, checking the list from begining
'to end until we have scanned the entire listbox once for the amount of items
'in the listbox
Next ListItems
End Function

Sub StringLoadFontsToComboBox(TheComboBox As ComboBox)
'This sub adds all registered fonts on the computer to a combobox. Example:

'StringLoadFontsToComboBox Combo1

'This loads all the fonts to TheComboBox, or in this case, Combo1.
'NOTE: You may want to set the 'Sorted' property of the combbox to
'True. This will make it so all the fonts are arranged alphabetically,
'instead of put in by the number it associates with.

'Code:

'We start by clearing the combobox so that there isnt anything in the box
'that isn't a font
TheComboBox.Clear

'Next we put the amount of fonts on the computer into the ComputerFonts
'variable. We use the - 1 because 0 is counted as the first font, and we
'dont want to add an extra space to the combobox.
For ComputerFonts = 0 To Screen.FontCount - 1
    'We add the fonts to TheComboBox by using the AddItem property. The rest
    'of the code on this line means that we are taking the ComputerFonts
    'variable, which is the number of fonts on this computer, and adding
    'the computer font that matches the number of the current value of
    'ComputerFonts.
    TheComboBox.AddItem (Screen.Fonts(ComputerFonts))
'We add this line to tell the program to loop again. It will keep looping
'until the ComputerFonts variable has reached the amount of fonts on the
'computer, or in other words, when all the fonts are added to the combobox.
Next ComputerFonts

'This is just an extra line so that when the process is complete, the title
'of the the combobox is not blank. Instead, it will say "Choose A Font"
TheComboBox.Text = "Choose A Font"
End Sub

Function ZoneLobbyChatSend(MessageToSend As String)
'This is a pretty cool sub that finds the Internet Gaming Zone chat lobby and
'sends a message to it. Example:

'Call ZoneLobbyChatSend("TestingSendChat")

'This sends the message, "TestingSendChat", to the Internet Gaming Zone
'lobby.

'Code:

'NOTE: When you look at these codes, and you see where I wrote
'"ZoneLobbyClientWnd", "Edit", and "ZoneLobbyWnd", don't think I made them
'up. There are no functions or subs titled those names. What I did to
'retrieve those names was use an API Spy which you can get almost anywhere
'over the web. I used the API Apy and retrieved the Window class for that
'certain part and it gave me with the names "Edit" and "ZoneLobbyClientWnd"
'so that I could easily put it in my program and start programming with it.

'First we have to dim our three values to store the handles of the different
'parts of the Zone chat window. We must declare this values as Long, because,
'as you will see, the variables will have to store numbers in them, so its
'safest to use the Long type because it supports the longest range of numbers.
Dim ZoneChatWindow As Long, ZoneChatBox As Long, ZoneTextBox As Long

'Here we use the FindWindow API, which is found in the declarations section,
'to find the Chat room window. The FindWindow API is used in this manner:

'FindWindow(The window class, the name of the class).

'The only time we would put in 'the name of the class', would be if there
'were multiple windows with the same window class. In this case, the window
'class is 'ZoneLobbyClientWnd', which is so unique that we dont need the
'name.
ZoneChatWindow& = FindWindow("ZoneLobbyClientWnd", vbNullString)
'Here, using the FindWindowEx API (The same as the FindWindow API but this
'API finds the parts that are in the window found with the FindWindow API),
'we find the handle of the "ZoneLobbyWnd" window, one step in from
'ZoneChatWindow, which is the parent handle. A parent is basically the main
'window which has different parts, and a child is one of the parts in a
'parent. FindWindowEx retrieves children, FindWindow retrieves parents.
ZoneChatBox& = FindWindowEx(ZoneChatWindow&, 0&, "ZoneLobbyWnd", vbNullString)
'Now we take our third step in, to the "Edit" window which is the class name
'of the textbox in which Zone users type. We store that last handle value
'in the ZoneTextBox variable.
ZoneTextBox& = FindWindowEx(ZoneChatBox&, 0&, "Edit", vbNullString)

'Now we get into the code section that actually sends the data. In this first line,
'we use the SendMessageByString API to send the text to the zone textbox.
'The handle for the zonetextbox is stored in the variable 'ZoneTextBox', so
'we have the handle and we leave the rest up to the API. The
'SendMessageByString API works like this:

'SendMessageByString(The handle by the place, the method you want to use to
'input the data, the handle for any numerical data or information you want
'to send, the message to send)

'So the handle of the place is the textbox handle stored in 'ZoneTextBox',
'we want to use the WM_SETTEXT method which places text in the handle,
'we dont want to send numerical data, and MessageToSend is part of the
'function, which is the message to send to the chatroom.
Call SendMessageByString(ZoneTextBox, WM_SETTEXT, 0&, MessageToSend)
'SendMessageLong is the same thing as SendMessageByString but it sends
'"actions" to the handle. Over here, we simply send the WM_CHAR and
'ENTER_KEY actions which send the ascii character(WM_CHAR) for the enter
'key, which is 13 (ENTER_KEY), and that will send the text to the chatroom.
Call SendMessageLong(ZoneTextBox, WM_CHAR, ENTER_KEY, 0&)
End Function

Function ZoneGameChatSend(MessageToSend As String)
'This is a pretty cool sub that finds the current game you are in
'(not current lobby), and sends a message to it. Example:

'Call ZoneGameChatSend("TestingSendChat")

'This sends the message, "TestingSendChat", to the Internet Gaming Zone
'game room that you are currentl in.

'Code:

'NOTE: When you look at these codes, and you see where I wrote
'"ZoneLaunchPad" and "SuperEdit", don't think I made them up. There are no
'functions or subs titled those names. What I did to retrieve those names
'was use an API Spy which you can get almost anywhere over the web. I used
'the API Apy and retrieved the Window class for that certain part and it
'gave me with the names "SuperEdit" and "ZoneLaunchPad" so that I could
'easily put it in my program and start programming with it.

'First we have to dim our two variables to store the handles of the
'different parts of the Zone game room window. We must declare these
'values as Long, because, as you will see, the variables will have to store
'numbers in them, so its safest to use the Long type because it supports
'the longest range of numbers.
Dim ZoneRoomWindow As Long, ZoneChatTextBox As Long

'Here we use the FindWindow API, which is found in the declarations section,
'to find the Chat room window. The FindWindow API is used in this manner:

'FindWindow(The window class, the name of the class).

'The only time we would put in 'the name of the class', would be if there
'were multiple windows with the same window class. In this case, the window
'class is 'ZoneLaunchPad', which is so unique that we dont need the name.
ZoneRoomWindow& = FindWindow("ZoneLaunchPad", vbNullString)
'Here, using the FindWindowEx API (The same as the FindWindow API but this
'API finds the parts that are in the window found with the FindWindow API),
'we find the handle of the "SuperEdit" window, because that is the textbox
'that users send chat too. A parent is basically the main window which has
'different parts, and a child is one of the parts in a parent. FindWindowEx
'retrieves children, FindWindow retrieves parents. In this case, our parent
'would be ZoneRoomWindow, and the child would be ZoneChatTextBox.
ZoneChatTextBox& = FindWindowEx(ZoneRoomWindow&, 0&, "SuperEdit", vbNullString)

'Now we get into the code section that actually sends the data. In this first
'line, we use the SendMessageByString API to send the text to the zone textbox.
'The handle for the zonetextbox is stored in the variable 'ZoneTextBox', so we
'have the handle and we leave the rest up to the API. The SendMessageByString
'API works like this:

'SendMessageByString(The handle by the place, the method you want to use to
'input the data, the handle for any numerical data or information you want
'to send, the message to send).

'So the handle of the place is the textbox handle stored in 'ZoneChatTextBox',
'we want to use the WM_SETTEXT method which places text in the handle, we dont
'want to send numerical data, and MessageToSend is part of the function, which
'is the message to send to the chatroom.
Call SendMessageByString(ZoneChatTextBox, WM_SETTEXT, 0&, MessageToSend)
'SendMessageLong is the same thing as SendMessageByString but it sends "actions"
'to the handle. Over here, we simply send the WM_CHAR and ENTER_KEY actions
'which send the ascii character(WM_CHAR) for the enter key, which is the
'13 (ENTER_KEY), and that will send the text to the chatroom.
Call SendMessageLong(ZoneChatTextBox, WM_CHAR, ENTER_KEY, 0&)
End Function

Function APIButtonClick(ButtonHandle As Long)
'This is a helpful sub that uses the SendMessage API to
'click any button. Example:

'Call APIButtonClick(ButtonHandle&)

'You would first have to get the buttons handle and replace
'the 'ButtonHandle' variable with the buttons handle.

'Code:

'Use the SendMessage API to declare the LeftButtonDown
'and LeftButtonUp constants on the ButtonHandle.
Call PostMessage(ButtonHandle&, WM_LBUTTONDOWN, 0&, 0&)
'We use this to simulate a "real" click. Its not necessary
'but makes it really show the click.
Call ProgramPause(0.1)
Call PostMessage(ButtonHandle&, WM_LBUTTONUP, 0&, 0&)
End Function

Function APICloseWindow(WindowHandle As Long)
'This simple sub will close any window using the PostMessage API and the
'WM_Close constants. Example:

'Call APICloseWindow(WindowHandle&)

'First you must replace 'WindowHandle' with the actual handle of the window.

'Code:

'Using the PostMessage API, apply the WM_Close constant on the WindowHandle.
Call PostMessage(WindowHandle&, WM_CLOSE, 0&, 0&)
End Function

Function ProgramPause(Seconds As Long)
'This is a very simple function that pauses the program for a requested
'amount of time. Example:

'Call ProgramPause(1)

'This will pause the program for 1 second.

'Code:

'Using the Sleep API, we simply tell the program to sleep for the duration
'that the programmer sets, which is stored in the Seconds
'variable.
'NOTE: We declared the Seconds variable as long because it
'is safest. It is safest because the type Long can support the longest
'range of numbers.
'We multiply the Seconds variable by 1000 so that the programmer doesnt
'have to count in milliseconds. Multiplying by 1000 will make it seconds
'instead of milliseconds.
Call Sleep(Seconds& * 1000)
End Function

Function AIMSendInstantMessage(ScreenName As String, TheMessage As String)
'This is a cool sub that sends an AIM Instant Message to the user
'specified ScreenName. It sends the user specified message stored under
'TheMessage. Example:

'Call AIMSendInstantMessage("yoshiii", "haha, your coding ownz me")
 
'Hehe. You can also send data from two textboxes...

'Call AIMSendInstantMessage(Text1.Text, Text2.Text)

'And now onto the huge code:

'This is the lengthy but important declaration section. We declare all
'variables as long because the variables will be storing handles which
'are the longer number code names for each window. We use the Long type
'because it can store the longest range of numbers.
Dim IMWindow As Long, Frame As Long, SnBox As Long
Dim ImWindowTwo As Long, MessageBox As Long
Dim ImWindowThree As Long, SendButton As Long

'First we open a new, clean Instant Message with my AIM_OpenInstantMessage
'function.
Call AIMOpenInstantMessage

'NOTE: Both the FindWindow and FindWindowEx Api are reviewed in the
'Zone_ChatSend functions. I won't explain them here.

'Begin a loop to find the Screename box...
Do: DoEvents
    'First we store the handle of the whole InstantMessage in the ImWindow
    'variable.
    IMWindow& = FindWindow("AIM_IMessage", vbNullString)
    'Now we store the handle of the frame that contains the SN textbox in the
    'Frame variable.
    Frame& = FindWindowEx(IMWindow&, 0&, "_Oscar_PersistantCombo", vbNullString)
    'Finally we store the handle of the SN textbox in the TextBox variable.
    SnBox& = FindWindowEx(Frame&, 0&, "Edit", vbNullString)
Loop Until IMWindow& And Frame& And SnBox& <> 0
'Use my APISendMessage function to send data in the ScreenName variable
'to the SnBox.
Call APISendMessage(SnBox&, ScreenName$)

'We use the my ProgramPause function to slow down the program. We do this
'because first of all we know that API is the fastest possible way of doing
'operations. In fact, it is so fast, that the program can't keep up with it.
'So to "cut the program some slack", we use the ProgramPause function to
'slow down the API and let the program catch up.
Call ProgramPause(0.001)

'Now we begin the search for the message box. We use the FindWindow API to
'target the whole IM again, which is why I named the variable ImWindowTwo,
'which is really not necessary.
ImWindowTwo& = FindWindow("AIM_IMessage", vbNullString)
'Here we use the FindWindowEx API to target the message box in the AIM
'InstantMessage window. This is where it gets tricky...
MessageBox& = FindWindowEx(ImWindowTwo&, 0&, "WndAte32Class", vbNullString)
'I don't know if AIM meant it this way or if it was an accident, but there
'are multiple windows named WndAte32Class and all with the same caption.
'On the next line I used the GetWindow API to retrieve the second
'WndAte32Class, which is the one that stores the messagebox. I found this
'out by using Spy ++. I'm not going to start talking about that, but it
'comes with Visual C++. It allows you to see every window handle and
'caption, and all the windows that they hold. This is how I knew the
'second WndAte32Class was the right one.
MessageBox& = GetWindow(MessageBox&, 2)

'Here we use my APISendMessage function to send the message stored in
'the TheMessage variable, to the message boxes' window handle stored under
'the MessageBox variable.
Call APISendMessage(MessageBox&, TheMessage$)

'Finally we use API on our last part - The Send IM Button. Here we use the
'FindWindow API to find the entire AIM instant message AGAIN, which is why
'I named it ImWindowThree.
ImWindowThree& = FindWindow("AIM_IMessage", vbNullString)
'Now we use the FindWindowEx API to retrieve the handle for the button and
'store it in the SendButton variable.
SendButton& = FindWindowEx(ImWindowThree&, 0&, "_Oscar_IconBtn", vbNullString)

'Now we use my my API_ButtonClick function to click the Send IM button.
Call APIButtonClick(SendButton&)
'Here we use my ProgramPause function to slow the program down for 1 second
'so that the window doesnt get closed before the IM is sent.
Call ProgramPause(0.001)
'This is optional, and you can delete it if you dont want this (but it
'is reccomended which you will see why later on). This line will use
'my API_CloseWindow function to close the entire Instant Message, stored
'under the ImWindowThree variable
Call APICloseWindow(ImWindowThree&)
End Function

Function AIMOpenInstantMessage()
'This sub simply opens a new, clean InstantMessage in AIM. Example:

'Call AimOpenInstantMessage

'This will open a new, clean IM in AIM.

'Code:

'Here we declare all our variables that will be storing the handles of
'the AIM windows. We declare them as Long because long is the safest
'because it supports the longest range of numbers.
Dim AimWindow As Long, ButtonTab As Long, Button As Long

'Here, we store the handle for the main AIM window under the
'AimWindow variable. We use the FindWindow API (explained in the
'ZoneChatSend functions), to find the window with a handle name of
'_Oscar_BuddyListWin, which is the handle name of the main AIM window.
AimWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
'Now we store the handle for the frame that holds the Send IM button in
'the ButtonTab variable. We use the FindWindowEx API (explained in the
'ZoneChatSend functions) to find a handle with the name _Oscar_TabGroup
'in the AimWindow.
ButtonTab& = FindWindowEx(AimWindow&, 0&, "_Oscar_TabGroup", vbNullString)
'Finally we store the handle for the Send IM button in the Button variable,
'using the FindWindowEx API. Using this API, we look for the handle named
'_Oscar_IconBtn, which is the name for the button in the button frame.
Button& = FindWindowEx(ButtonTab&, 0&, "_Oscar_IconBtn", vbNullString)

'Using my API_ButtonClick function, we click the button using its
'handle to identify it.
Call APIButtonClick(Button&)
End Function

Function APISendMessage(WindowHandle As Long, Message As String)
'This is a quick sub that sends a message to any window. Example:

'Call APISendMessage(WindowHandle&, Message)

'Before using, you must first replace WindowHandle with the actual
'handle, and Message with the message you wish to send.

'Code:

'First, using the Len statement, check to see if the length of the string
'is greater than 0. If it is 0, then there is no use in sending becuase
'there is nothing to send. If it is not 0, then that means that there is
'something to send so...
If Len(Message) > 0 Then
    'Using the SendMessageByString API, we sent the message stored under
    'the Message variable toe the window handle stored under the
    'WindowHandle variable. We use the WM_SETTEXT constant to place
    'the message in the window.
    Call SendMessageByString(WindowHandle&, WM_SETTEXT, 0&, Message$)
'End the sub.
End If
End Function

Function APIPushEnter(WindowHandle As Long)
'This is a quick sub that pushes the enter key in any window. Example:

'Call APIPushEnter(WindowHandle&)

'You must first replace WindowHandle with the handle of the window you
'are dealing with.

'Code:

'Using the SendMessageLong API, we imply the ascii code (WM_CHAR) of the
'enter key (ENTER_KEY), on the window. This results the same way as the
'user pushing the Enter key.
Call SendMessageLong(WindowHandle&, WM_CHAR, ENTER_KEY, 0&)
End Function

Function APIPushSpaceBar(WindowHandle As Long)
'This quick subs pushes the space bar in any window. Example:

'Call APIPushSpaceBar(WindowHandle&)

'You must first replace the WindowHandle variable with the actual
'handle of the window you are dealing with.

'Code:

'Using the SendMessageLong API, we apply the ascii key (WM_CHAR) of
'the space bar key (SPACE_BAR) on the window handle. This results
'the same way as the user pushing the space bar in that particular
'window.
Call SendMessageLong(WindowHandle&, WM_CHAR, SPACE_BAR, 0&)
End Function

Function AOLGoKeyword(Keyword As String)
'This is a cool function to go to a keyword in AOL. Example

'Call AOLGoKeyword("mtv")

'This will go to the mtv keyword on AOL.

'Code:

'This is the declaring section. We declare our variables as Long because
'they will be storing the numerical handles of our windows. We choose the
'Long type because it is safest because it supports the longest range
'of numbers.
Dim AOL As Long, Bar As Long, InBar As Long
Dim Combo As Long, KeyWordBox As Long

'NOTE: The FindWindow and FindWindowEx API will not be explained here.
'They are explained in the Zone chat send functions.

'First we get the handle of the entire AOL window and store it under
'the AOL variable.
AOL& = FindWindow("AOL Frame25", vbNullString)
'Next we go one step in from the entire AOL window to the toolbar,
'and we will store this handle in the Bar variable.
Bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
'Now we go one step into the bar, to the panel that holds the
'keyword textbox. We store this panels' handle in the InBar variable.
InBar& = FindWindowEx(Bar&, 0&, "_AOL_Toolbar", vbNullString)
'Now we go one step in from the panel to the Combobox that supports
'the dropdown ability of the Keyword textbox. We store this handle
'under the Combo variable.
Combo& = FindWindowEx(InBar&, 0&, "_AOL_Combobox", vbNullString)
'Finally we reach keyword textbox itself, and we store its handle
'under the KeyWordBox variable.
KeyWordBox& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)

'Now we use my APISendMessage function to send the user defined
'keyword, stored under the KeyWord variable, to the textbox.
Call APISendMessage(KeyWordBox&, Keyword$)
'Now this is optional, but it makes the program look professional.
'This is my function to add a space to the textbox. In general,
'the textbox is held as all selected/highlited when the mouse is
'not in it and the cursor is not in it. This sends the space bar,
'so all the text in the box will be deleted, so the user doesn't
'see how the program actually deposits the text in the textbox.
Call APIPushSpaceBar(KeyWordBox&)
'Finally, the code that starts it all, the enter key. Using my
'APIPushEnter function, we use the enter key to submit the keyword,
'and open it in AOL.
Call APIPushEnter(KeyWordBox&)
End Function

Function APIPushTabKey(WindowHandle As Long)
'This quick subs pushes the tab key in any window. Example:

'Call APIPushTabKey(WindowHandle&)

'You must first replace the WindowHandle variable with the actual
'handle of the window you are dealing with.

'Code:

'Using the SendMessageLong API, we apply the ascii key (WM_CHAR) of
'the tab key (TAB_KEY) on the window handle. This results
'the same way as the user pushing the tab key in that particular window.
Call SendMessageLong(WindowHandle&, WM_CHAR, TAB_KEY, 0&)
End Function

Function AOLChangePassword(OldPassword As String, NewPassword As String)
'This cool sub changes your AOL password. It is slightly buggy,
'so please e-mail me if you find any bugs. Example:

'Call AOLChangePassword("chinatown", "lilypad")

'This would work if your old password was "chinatown". It would change
'your password to "lilypad".

'NOTE: AOL requires a password length of a minimum of 6.

'Code:

'Here is the declaration section. We declare them all as long because
'Long supports the longest range of numbers, which makes it the safest
'for storing window handles because window handles are numbers.
Dim Form As Long, Button As Long
Dim FormTwo As Long, OldPass As Long, NewPass As Long
Dim NewPassTwo As Long, Finish As Long
Dim Box As Long, ButtonOK As Long

'NOTE: The FindWindow and FindWindowEx API are explained in the
'Zone send chat functions.

'Here we use my AOLGoKeyword function to go to keyword "Password".
Call AOLGoKeyword("Password")
'Let AOL catch up to the program.

'Start a loop
Do: DoEvents
'Here we begin the first section of API. Using the FindWindow and
'FindWindowEx API, we find the change password form, and store
'its handle in the 'Form' variable.
    Form& = FindWindow("_AOL_Modal", vbNullString)
'Now we go one step in from the form, to the Change Password button.
    Button& = FindWindowEx(Form&, 0&, "_AOL_Icon", vbNullString)
'We start a loop here. This is more accurate than pausing. It is
'a sure way to make sure that the windows have appeared on the
'screen. We want to make sure they are not 0 because if they were
'zero, then that means the program can't find the window.
Loop Until Form& And Button& <> 0

'We click the Change Password button.
Call APIButtonClick(Button&)
'We pause for 1 second so AOL can catch up with the program.

'Start a loop for which you will see why soon...
Do: DoEvents
'Here we find the password selection form, and store it under
'the FormTwo variable.
    FormTwo& = FindWindow("_AOL_Modal", "Change Your Password")
'Now this was a very complex part and I thank Volter for helping
'me out. On the form, there are three textboxes with the same
'handle of "_AOL_Edit", and they all have the same parent. This
'is where the 0& part in most of my FindWindowEx API codes comes
'in. We use the name of the last textbox, so that the API knows not
'to find the same textbox over and over again. We store the textbox
'for the old password in the OldPass variable.
    OldPass& = FindWindowEx(FormTwo&, 0&, "_AOL_Edit", vbNullString)
'Now we tell the API to find the other "_AOL_Edit" that is not the
'OldPass variable.
    NewPass& = FindWindowEx(FormTwo&, OldPass&, "_AOL_Edit", vbNullString)
'Now we tell the API to find the last "_AOL_Edit" that is neither
'the OldPass variable or the NewPass variable. We store this lastbox
'under the NewPassTwo variable.
    NewPassTwo& = FindWindowEx(FormTwo&, NewPass&, "_AOL_Edit", vbNullString)
'Now, instead of pauses, the program will loop until the window handles
'for each of the variables is not 0. This will prevent the need for
'pauses and mess-ups. In other words, this is a sure way to get the
'job done.
Loop Until OldPass& And NewPass& And NewPassTwo& <> 0

'We send the old password (OldPassword$) to the old password box.
Call APISendMessage(OldPass&, OldPassword$)
'We send the new password (NewPassword$) to the new password box.
Call APISendMessage(NewPass&, NewPassword$)
'We send the new password again because AOL wants the users to
'confirm their passwords. We send the same new password (NewPassword$)
'to the last box.
Call APISendMessage(NewPassTwo&, NewPassword$)

'Use the FindWindowEx API to store the Final Save New Password button
'in the Finish variable.
Finish& = FindWindowEx(FormTwo&, 0&, "_AOL_Icon", vbNullString)
'Click the Save New Password button using my API_ButtonClick function.
Call APIButtonClick(Finish&)

'Now that we changed the password, we get rid of the message box that
'comes up to alert you. You can take this out if you want, but I think
'it is better.

'Start another loop for the same reason I explained previously..
Do: DoEvents
'Find the handle of the message box and store it under the Box variable
'using the FindWindow API.
    Box& = FindWindow("#32770", "America Online")
'Go one step into the message box to the OK button, and store that handle
'under the ButtonOK variable.
    ButtonOK& = FindWindowEx(Box&, 0&, "Button", vbNullString)
'We loop so that we are sure we find the windows before proceeding. I
'didn't explain this before, but if the variables equal 0, then it is
'clear that the program didn't find the window.
Loop Until Box& And ButtonOK& <> 0

'Click the button once to enable the button because it is not set by
'default due to the nature of Windows. A message boxes button is not
'set by default when it is not the top window. This is why we use
'API - It is the fastest and most accurate way to do something without
'error.
Call APIButtonClick(ButtonOK&)
'Now that the button is enabled, we click it once more, and now the
'message box is gone and we have successfully changed your password.
Call APIButtonClick(ButtonOK&)
End Function

Function AOLSendInstantMessage(Person As String, Message As String)
'This is a long but fairly simple code to send an AOLInstantMessage to
'someone. Example:

'Call AOLSendInstantMessage("yoshiii", "this function is awesome")

'That will send an IM to me saying "this function is awesome". You
'could also send Info from text boxes like this...

'Call AOLSendInstantMessage(Text1.Text, Text2.Text)

'This will send an IM to the person whose name is written in Text1,
'with a message of the text written in Text2.

'Code:

'We declare our variables as long since they will be holding handle
'value.
Dim AOLWindow As Long, MDI As Long, IM As Long, IMTwo As Long
Dim Box As Long, SendButton As Long, MessageBox As Long
Dim MessageBoxButton As Long

'We use the AOLGoKeyword function because it will load the keyword
'that automatically opens an IM.
Call AOLGoKeyword("Aol://9293:" & Person$)

'Begin the do loop for which you will see why after reading ahead...
Do: DoEvents
'Find AOL's handle and store it under the AOLWindow variable.
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
'Find the MCIClient(The gray space in the background) and store its
'handle under the MDI variable.
    MDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
'Find the handle of the IM and put it in the IM variable.
    IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
'Find the handle of the message box and put it in the Box variable.
    Box& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
'This part gets tricky. There are multiply _AOL_Icons, so skip
'them 1 by 1 until we find the one we need.
    SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
'Here is the one we need. It is the 9th one, which is the Send
'IM button.
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
'Keep doing this process over and over until all the windows are found.
Loop Until AOLWindow& And MDI& And IM& And Box& And SendButton& <> 0

'Send the user defined Message$ to the message box (Box).
Call APISendMessage(Box&, Message$)
'Click the send button.
Call APIButtonClick(SendButton&)
'Push the space bar on the send button.
Call APIPushSpaceBar(SendButton&)
'Click it again (I just wanted to make sure it gets clicked)
Call APIButtonClick(SendButton&)
'I want to make sure it gets the space (to activate it).
Call APIPushSpaceBar(SendButton&)

'Begin another loop...
Do: DoEvents
    'Make the MessageBox variable equal the handle of any
    'mesage boxes that may currently be open. The AOLMessageBoxFind
    'is a function I wrote to automatically retrieve that handle
    MessageBox& = AOLMessageBoxFind
    'Get the IM windows handle and store it in the IMTwo variable.
    IMTwo& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
'Keep repeating this loop until there a message box is found (that means
'we know there is an error), or if the IM dissapears (which means that
'the IM was sent successfully).
Loop Until MessageBox& <> 0 Or IMTwo& = 0

'Now, if a message box was found (meaning there was an error), then...
If MessageBox& <> 0& Then
    '...Click the spacebar on the button in the message box. We can
    'find the button on the message box easily with my AOLMessageBoxButton
    'function. We click the space bar to enable it.
    Call APIPushSpaceBar(AOLMessageBoxButton)
    'Click the OK button to close the messagebox.
    Call APIButtonClick(AOLMessageBoxButton)
    'And since there was an error with the IM, that means that the IM
    'window hasn't dissapeared. We use my APICloseWindow function to
    'close the window.
    Call APICloseWindow(IMTwo&)
End If
End Function
Function AOLMessageBoxFind()
'This is just like a mini-sub which I will be using for my other
'functions. It finds the any message boxes that AOL brings up. Example:

'If AOLMessageBoxFind = 0 then
'MsgBox "There are no message boxes."
'End If

'This will display a message box if AOL has no message boxes
'currently open.

'Code:

'Simply declare the MessageBox variable as long because we are going
'to store a handle in it.
Dim MessageBox As Long
'Store the handle number of the message box in the MessageBox variable.
MessageBox& = FindWindow("#32770", "America Online")
'Make the function equal the handle value of the message box so we can
'use the function name as displayed in the example code.
AOLMessageBoxFind = MessageBox&
End Function

Function AOLMessageBoxButton()
'This is basically the same as the AOLMessageBoxFind function but it
'goes one step further. This gets the handle for the "OK" button in an
'AOL message box. Here's an example you could use with this function and
'the AOLMessageBoxFind function and the APIButtonClick function:

'If AOLMessageBoxFind <> 0 Then
'Call APIButtonClick(AOLMessageBoxButton)
'MsgBox "There are no AOL message boxes currently open."
'End If

'This example will click the OK button in the message box and display
'the message box saying "There are no AOL message boxes currently open"
'if there is an AOL message box currently open.

'Code:

'Declare both variables as long since they will be storing handles.
Dim MessageBox As Long, TheButton As Long
'Find the message box and store its handle in the MessageBox variable.
MessageBox& = FindWindow("#32770", "America Online")
'Go one step further and find the Button in the message box and store
'its handle value in the TheButton variable.
TheButton& = FindWindowEx(MessageBox&, 0&, "Button", vbNullString)
'Make the function equal the handle of the button, so that you may use
'it as shown in the sample code above.
AOLMessageBoxButton = TheButton&
End Function

Function AIMSendChat(Message As String)
'This is an easy sub that sends a message to the AIM chatroom.
'Example:

'AIMSendChat("Where is yoshi that cool kid?")

'Heh, this just sends a message to the chatroom saying "Where is
'yoshi that cool kid?"

'Code:

'Declare all variables as long because they will be holding handles.
Dim ChatRoom As Long, Holder As Long, MessageBox As Long
Dim SendButton As Long

'Find the AIM chatroom and store its handle under the ChatRoom variable.
ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)
'Go one step in to the chatwindow to the group which has the textbox.
'Store its handle in the Holder variable.
Holder& = FindWindowEx(ChatRoom&, 0&, "WndAte32Class", vbNullString)
'Since there are two "WndAte32Class"'s, we have to go to the next one
'by skipping the handle of the old one so the program knows to find the
'next one.
Holder& = FindWindowEx(ChatRoom&, Holder&, "WndAte32Class", vbNullString)
'Finally, the textbox were se send messages from. Store its handle
'under the MessageBox variable.
MessageBox& = FindWindowEx(Holder&, 0&, "Ate32Class", vbNullString)

'Use my APISendMessage function to send the message stored under Message$
'to the message box.
Call APISendMessage(MessageBox&, Message$)

'This is not necessary, but I do it to be safe, so that we are sure
'the program has sent the message before proceeding.
Call ProgramPause(0.1)

'Since there are many buttons on the chatroom form with the same handle,
'we find the first one and store its handle under the SendButton variable.
SendButton& = FindWindowEx(ChatRoom&, 0&, "_Oscar_IconBtn", vbNullString)
'Go to the next button...
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_Oscar_IconBtn", vbNullString)
'..Keep going...
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_Oscar_IconBtn", vbNullString)
'Keep moving to the send button...
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_Oscar_IconBtn", vbNullString)
'Here is the send button, store its handle in the SendButton variable.
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_Oscar_IconBtn", vbNullString)
'Click the Send button, which will lead to the sending of our message
'to the chatroom.
Call APIButtonClick(SendButton&)
End Function

Function AIMOwnChatRoom()
'This is a rather silly sub which I made for fun. It fills the entire
'chatroom with the 'at' symbol (@). Example:

'AIMOwnChatRoom

'LoL, that was very complex. This will eat the chatroom.

'Code:

'Using my AIMSendChat function, send a chat of 473 or something like that
'@ symbols. If you want, you can change all those @'s to W's, which will
'take up even more space since W is the longest character.
Call AIMSendChat("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
End Function

Function StringSpiral(TheString As String)
'This sub takes a string and displays message boxes of it after spiraling
'completely. Example:

'Call StringSpiral("Cow")

'This will display 4 message boxes. They will say "Cow", "owC", "wCo", "Cow".
'See the pattern? It will rotate it.

'NOTE: This sub was not meant for message boxes. It was meant for what you
'chose. A good idea would be to replace the MsgBox functions with an
'AOL Sendchat code or an AIM  Sendchat code.

'Code:

'Declare the variables. We will need strings and longs because we need to
'store characters and integers.
Dim TextLength, Product As String, Count As Long
Dim FirstPart As String, TheRest As String

'Make the TextLength variable equal the length of the TheString variable.
TextLength = Len(TheString$)
'Make the Product variable equal the TheString variable, which is the string
'to be spiraled.
'NOTE: The & " " part of this line of code is not necessary, I just add it
'because it is assumed that when spiraling, you do not want the first word
'to touch the end word and make a long jumble after a rotationg. You can kill
'the & " " if you want them to collide.
Product$ = TheString$ & " "
'Send a message box for the first time, which will start off the spiral.
MsgBox Product$

'Fill the Count variable from 1 (the start of the string) to the length of
'the string (the end of the string).
For Count& = 1 To TextLength
    'Get the first letter in the string and put it in the FirstPart variable.
    FirstPart$ = Mid(Product$, 1, 1)
    'Get the rest of the string, and put it in the TheRest variable.
    TheRest$ = Mid(Product$, 2, TextLength)
    'Make the product variable equal the data in the TheRest variable,
    'which is the second letter and on, and then the FirstPart variable,
    'which was the first letter. It is now at the end of the string to
    'make a spiraling effect.
    Product$ = TheRest$ + FirstPart$
    'Message box the data in the Product variable, to see the spiral.
    MsgBox Product$
'Do this all again, until the whole string has been spiraled from begginging
'and back to begginging again.
Next Count&
End Function

Function AOLFindChatRoom() As Long
'This sub will tell you if there is a chatroom open. Example:
'In the command button click you may want to do something like this:

'Private Sub Command1_Click()
'Dim TheRoom as long
'TheRoom& = AOLFindChatRoom&
'If TheRoom& = 0 then msgbox "No room open!"
'If TheRoom& <> 0 then msgbox "Room found!"
'End Sub

'This will display the No room open! message if there is no chatroom open
'and the Room found! message if a room is found.

'Code:

'Declare variables as long so they can store handles.
Dim AOL As Long, MDI As Long, Room As Long, TheBox As Long, TheCombo As Long
Dim RoomList As Long, ButtonCheck As Long, LabelCheck As Long

'Find the AOL window, store its handle in the AOL variable.
AOL& = FindWindow("AOL Frame25", vbNullString)
'Find the MDI Client and store its handle in the MDI variable.
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
'Go one step further to an AOL Child. We are not sure if it is the chatroom
'because there are many AOL Childs. This will be corrected soon...
Room& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
'Go a step into the Room variable to the big chatwindow.
TheBox& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
'Find the listbox in the room...
RoomList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
'Find the ComboBox..
TheCombo& = FindWindowEx(Room&, 0&, "_AOL_Combobox", vbNullString)

'Look to make sure there is a button, in this case, the button is
'the Forecolor button. We don't really care about what buttons we are
'looking for yet. We just want to make sure, if you havent noticed,
'that we find a window that has all this stuff in it. Thats what makes the
'chatroom unique, it has many parts to it.
ButtonCheck& = FindWindowEx(Room&, 0&, "_AOL_Icon", vbNullString)
'Find the label and store its value in the LabelCheck variable...
LabelCheck& = FindWindowEx(Room&, 0&, "_AOL_Static", vbNullString)
'These 4 lines say that if all of these upper variables dont equal zero,
'which means they were found which means that this is a room, then make
'this function equal the room handle and exit the functino because we
'are done now.

If TheBox& <> 0& And RoomList& <> 0& And ButtonCheck& <> 0& And LabelCheck& <> 0& And TheCombo& <> 0& Then
    AOLFindChatRoom& = Room&
    Exit Function
Else

'On the other hand, if its not found, then start this looping process...
    Do: DoEvents
         
'Im not going to say all these individually, but find all those
'parts of the room again...
         Room& = FindWindowEx(MDI&, Room&, "AOL Child", vbNullString)
         TheBox& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
         RoomList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
         ButtonCheck& = FindWindowEx(Room&, 0&, "_AOL_Icon", vbNullString)
         LabelCheck& = FindWindowEx(Room&, 0&, "_AOL_Static", vbNullString)
         TheCombo& = FindWindowEx(Room&, 0&, "_AOL_Combobox", vbNullString)
'And if now it finds them all, then make the function equal the room handle
'and exit the function.
        If TheBox& And RoomList& And ButtonCheck& And LabelCheck& And TheCombo& <> 0& Then
            AOLFindChatRoom& = Room&
            Exit Function
        End If

'Now keep looping until we can't find the room anymore. If we can't find the
'room, then end this loop. We use this code because in the code above,
'if the room is ever found, the program will send to it. If its not found,
'that means Room& = 0& so we dont need to continue...
    Loop Until Room& = 0&
'End the loop...
End If
'Make the function equal the value of the room. If the room equals 0, then
'the AOLFindChatRoom will also equal 0. If the room is not 0, that means it
'is found, so put the value of the room in the AOLFindChatRoom function.
AOLFindChatRoom = Room&
End Function

Function AOLChatSend(Message As String)
'This is a cool sendchat code! It replaces the text you were writing
'before the program sendchatted! Example:

'Call AOLChatSend("hah")

'This will send the message "hah" to the AOL chatroom.

'Code:

'Declare variables as long to store handles. We also need strings to
'store characters.
Dim ChatRoom As Long, textbox As Long, Storage As String, SendButton As Long

'If there is not chatroom open, then exit this function since there is no
'need to keep going.

If AOLFindChatRoom& = 0 Then Exit Function

'Just use my AOLFindChatRoom variable to get the room handle and put it in
'the ChatRoom variable.
ChatRoom& = AOLFindChatRoom&

'Find the text box...
textbox& = FindWindowEx(ChatRoom&, 0&, "RICHCNTL", vbNullString)
'...Text box found. Store its value udner the TextBox variable.
textbox& = FindWindowEx(ChatRoom&, textbox&, "RICHCNTL", vbNullString)

'Find the send button...
SendButton& = FindWindowEx(ChatRoom&, 0&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_AOL_Icon", vbNullString)
'... Send button found. Store it in the SendButton variable
SendButton& = FindWindowEx(ChatRoom&, SendButton&, "_AOL_Icon", vbNullString)

'This is cool.. call my APIGetText function on the textbox and store its
'data in the Storage variable for which you will see why later...
Storage$ = APIGetText(textbox&)
'Make the box equal nothing.
Call APISetText(textbox&, "")
'Send the message to the textbox..
Call APISendMessage(textbox&, Message$)

'Call do events to eliminate freezing.
Do: DoEvents
'Push enter in the box...
Call APIButtonClick(SendButton&)
'Keeping push enter until the value has gone through.
Loop Until APIGetText(textbox&) <> Message$

'Call DoEvents to kill freezing...
Do: DoEvents
'THis is where Storage$ comes in, it will paste the data in the Storage
'variable into the textbox so that if you were typing in the chatroom and
'then it sends chat, its irritating because your text mixes with the
'the data the program wanted to send. Now, after it was paste, it will bring
'that text back into the box so you can continue talking...
Call APISetText(textbox&, Storage$)
'Loop until the message has gone through.
Loop Until APIGetText(textbox&) = Storage$
End Function

Function APIGetText(WindowHandle As Long)
'This sub gets the text from any handle. Example:

'Call APIGetText(BlaBla&)

'This will get the text from the handle stored in BlaBla.

'Code:

'Declare our variables, we need one string and one long.
Dim Filter As String, Text As Long
    
'First, get the length of the string in the handle so we know how much
'to copy paste, and store it in the Text variable.
Text& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    
'We fill the Filter variable with the amount of spaces and text necesary
'in the Text variable.
Filter$ = Space$(Text&)
    
'Now we can't use my APISendMessage function here because there are some
'alterations to make, and besides, this gets text, not places it. We Send
'the message to the WindowHandle, using the GETTEXT parameter. We tell it
'to Add 1 to the Text variable so the length of the string is defined and we
'dont miss one. This is because even strings start at 0, and we have to add
'one since the GETTEXTLENGTH paramter starts at 1.
Call SendMessageByString(WindowHandle&, WM_GETTEXT, Text& + 1, Filter$)

'Overall, make the function equal the Filter variable, which is our result.
APIGetText = Filter$
End Function

Function APISetText(WindowHandle As Long, thetext As String)
'This sub, instead of the APISendMessage function, will force the box
'to equal this text, instead of adding it. Example:

'Call APISetText(WindowHandle, "HeH")

'This will, no matter how much text is in the WindowHandle, it will remove
'that text and then display "HeH".

'Code:

'Make the box equal nothing.
Call SendMessageByString(WindowHandle&, WM_SETTEXT, 0, "")
'Make the box equal the requested text.
Call SendMessageByString(WindowHandle&, WM_SETTEXT, 0, thetext$)
End Function

Function CompPushStartMenuButton()
'This sub clicks the start button for the Windows 98 OS. It may work for 95,
'but I haven't tried it out on 95 yet. Example:

'CompPushStartMenuButton

'Ok, that was complicated. This will push the StartMenu button.

'Code:

'Declare as long to store handles...
Dim Holder As Long, Button As Long

'Find the button holder and put it in the Holder variable...
Holder& = FindWindow("Shell_TrayWnd", vbNullString)
'...Get the button and store it in the Button& variable...
Button& = FindWindowEx(Holder&, 0&, "Button", vbNullString)

'...Click the button.
Call APIButtonClick(Button&)
End Function

Function CompHideStartMenuButton()
'This evil sub hides the start menu button =P. Example:

'CompHideStartMenuButton

'That was difficult.. that hides the start menu button.

'Code:

'Declare as long to store handles...
Dim Holder As Long, Button As Long
'...Get the button holder and store it in the Holder variable...
Holder& = FindWindow("Shell_TrayWnd", vbNullString)
'...Get the button...
Button& = FindWindowEx(Holder&, 0&, "Button", vbNullString)
'Using the ShowWindow API, call the SW_HIDE constant on the Button, which
'will hide the startmenu button.
Call ShowWindow(Button&, SW_HIDE)
End Function

Function CompShowStartMenuButton()
'This safe sub shows the start menu button =P. Example:

'CompShowStartMenuButton

'That was difficult.. that shows the start menu button.

'Code:

'Declare as long to store handles...
Dim Holder As Long, Button As Long
'...Get the button holder and store it in the Holder variable...
Holder& = FindWindow("Shell_TrayWnd", vbNullString)
'...Get the button...
Button& = FindWindowEx(Holder&, 0&, "Button", vbNullString)
'Using the ShowWindow API, call the SW_SHOW constant on the Button, which
'will show the startmenu button.
Call ShowWindow(Button&, SW_SHOW)
End Function

Function CompHideStartMenu()
'This neat sub hides the startmenu. Example:

'CompHideStartMenu

'This hides the startmenu.

'Code:

'Declare variable as long to store handles...
Dim StartMenu As Long
'...Get the startmenu and store it in the StartMenu variable...
StartMenu& = FindWindow("Shell_TrayWnd", vbNullString)

'...Using the ShowWindow API, call the SW_HIDE constant to hide the StartMenu.
Call ShowWindow(StartMenu&, SW_HIDE)
End Function

Function CompShowStartMenu()
'This sub shows the startmenu. Example:

'CompShowStartMenu

'This shows the startmenu.

'Code:

'Declare variable as long to store handles...
Dim StartMenu As Long
'...Get the startmenu and store it in the StartMenu variable...
StartMenu& = FindWindow("Shell_TrayWnd", vbNullString)

'...Using the ShowWindow API, call the SW_SHOW constant to show the StartMenu.
Call ShowWindow(StartMenu&, SW_SHOW)
End Function

Function CompChangeTime(TimeOfDay As String, AMorPM As String)
'This very short sub changes the time of day on the computer. Example:

'Call CompChangeTime("5:00", "PM")

'This will change the time to 5:00 PM.

'Code:

'Just use the Time constant which is the visual basic source to retrieve the
'time. In this case, we are making the time equal the Time of the day and the
'AM or PM.
Time = TimeOfDay$ & AMorPM$
End Function

Function CompChangeDate(Month As Long, TheDate As Long, Year As Long)
'This quick sub changes the date on the computer. Example:

'Call CompChangeDate("4", "18", "1985")

'That changes the date on the computer to my birthday, heh.

'Code:

'Using the Date constant which is vb's way to get the date, we change the date
'by using the */*/**** format.
Date = Month& & "/" & TheDate& & "/" & Year&
End Function

Function APIShowWindow(handle As Long)
'This sub shows any window. Example:

'Call APIShowWindow(TheWindow&)

'You must first replace TheWindow with the window handle.

'Code:

'Use the ShowWindow API to apply ith SW_SHOW Constant on the window handle,
'therefore showing it.
Call ShowWindow(handle&, SW_SHOW)
End Function

Function APIHideWindow(handle As Long)
'This sub hides any window. Example:

'Call APIHideWindow(TheWindow&)

'You must first replace TheWindow with the window handle.

'Code:

'Use the ShowWindow API to apply ith SW_HIDE Constant on the window handle,
'therefore hidin it.
Call ShowWindow(handle&, SW_HIDE)
End Function

Function AOLHideToolbar()
'This very irritating sub hides the toolbar and leaves AOL looking very
'terrible. Example:

'AOLHideToolbar

'That hides the toolbar.

'Code:

'Declare as long to hold handles.
Dim AOL As Long, Bar As Long

'Find AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'...Find the toolbar...
Bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)

'Use my APIHideWindow function to hide the bar.
Call APIHideWindow(Bar&)
End Function

Function AOLShowToolbar()
'This is the reversal to my evil sub, the AOLHideToolbar. This brings
'the AOL Toolbar back. Example:

'AOLShowToolbar

'That shows the toolbar.

'Code:

'Declare as long to hold handles.
Dim AOL As Long, Bar As Long

'Find AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'...Find the toolbar...
Bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)

'Use my APIHideWindow function to show the bar.
Call APIShowWindow(Bar&)
End Function

Function APIGetWindowCaption(handle As Long)
'This API function retrieves the caption of any window. Example:

'Call APIGetWindowCaption(Handle&)

'You must first replace Handle& with the actual handle of the window
'you are dealing with.

'Code:

'Declare TheCaptionLength as long to store the length of the caption, and
'declare TheCaption as string to store the actual caption.
Dim TheCaptionLength As Long, thecaption As String

'Make the length equal the length of the text of the window, using the
'GetWindowTextLength API.
TheCaptionLength& = GetWindowTextLength(handle&)
'Make TheCaption equal a string with the length of the caption of the
'window.
thecaption$ = String(TheCaptionLength&, 0&)
'Using the GetWindowText API, get the caption of the window and store it in
'the TheCaption variable. We have to use the TheCaptionLength variable here
'so that the GetWindowText API
Call GetWindowText(handle&, thecaption$, TheCaptionLength& + 1)
'Make this function equal the caption of the window.
APIGetWindowCaption = thecaption$
End Function

Function AOLGetUploadStatus()
'This function retrieves the percent at which the current upload is at.
'Example:

'Label1.Caption = AOLGetUploadStatus

'This puts the upload status as the caption of Label1. Make sure Label1 is on
'the form when you use it.

'Code:

'Declare as long to store handles, and as string to hold characters/integers.
Dim Modal As Long, StatusBar As Long, Percent As String

'Begint he loop...
Do: DoEvents
'Find the AOLModal and put its handle in the Modal variable.
Modal& = FindWindow("_AOL_Modal", vbNullString)
'Find the status bar and put its handle in the StatusBar variable.
StatusBar& = FindWindowEx(Modal&, 0&, "_AOL_Gauge", vbNullString)
'Keep looping until these are found. Since I put DoEvents at the time, this
'won't lock up the program while its finding.
Loop Until Modal& And StatusBar& <> 0

'Using my APIGetWindowCaption function, get the caption of the AOL upload
'form, and move into it 17 characters, and read from there to 4 more after it.
'This will store the exact percent in the Percent variable.
Percent$ = Mid(APIGetWindowCaption(Modal&), 17, 4)
'Make this function equal the percent at which the upload is at.
AOLGetUploadStatus = Percent$
End Function

Function StringRemoveCharacter(TheString As String, Character As String)
'This sub removes all characters of your choice from the desired string.
'Example:

'Text1.Text = StringRemoveCharacter(Text1.Text, "w")

'This will make the text in Text1.Text equal Text1.Text after all
'the "w"'s have been removed.

'Code:

'If the program finds no characters of that type in TheString (returns a
'value of 0), then make the program keep TheString the same as it was before
'since there are no characters of that type to remove.
If InStr(TheString, Character$) = 0 Then
    StringRemoveCharacter = TheString$
    'Go to point 'SkipRest' so the program doesnt go through the
    'following code.
    GoTo SkipRest
End If

'Now that we know there is that character, make the function do the
'following steps from the beggining to the end of 'TheString'.
'The Len function gets the length of the strength from the
'beggining to the end.
For StringRemoveCharacter = 1 To Len(TheString)
    'Make the SearchString variable equal the point in 'TheString'
    'From which RemoveSpaces starts, and going up 1 more point.
    'This will eventually lead to searching of the whole string for
    'spaces. Notice the code we wrote before that mentions 1 to the
    'end of the string. It was all in 'StringRemoveCharacter', which is what
    'we are starting from and going up by 1.
    SearchString$ = Mid(TheString$, StringRemoveCharacter, 1)
    'Store the SearchString variable in 'BlankSearch', but dont remove
    'the previous data in BlankSearch because that may contain that character
    'to be removed.
    BlankSearch$ = BlankSearch$ & SearchString$
    'If the SearchString variable from before was just that character
    'then...
    If SearchString$ = Character$ Then
        'Then search 'BlankSearch' from the beggining to the end,
        'minusing 1 each time so we are sure that the character is
        'removed.
        BlankSearch$ = Mid(BlankSearch$, 1, Len(BlankSearch$) - 1)
    End If
'Begin the next series in StringRemoveCharacter until, as we wrote on the first
'line (For StringRemoveCharacter = 1 To Len(TheString)), we reach the end of the
'string (TheString).
Next StringRemoveCharacter

'Once the entire loop is complete, make the RemoveSpaces function equal
'our finished product without any spaces, 'BlankSearch'.
StringRemoveCharacter = BlankSearch$

'The skipping point for where the program would go if there were no
'spaces.
SkipRest:
End Function

Sub ProgramStayOnTop(TheForm As Form)
'This sub makes any form the topmost window on the computer while its running.
'For example, you would put this in the Form_Load proc:

'Call ProgramStayOnTop(Me)

'This will make the current form stay ontop all other windows.

'Code:

'Using the SetWindowPos API, we call the HWND_TOPMOST constant on the form.
Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Function ProgramFormMove(TheForm As Form)
'This sub will alow the user to drag the form with the mouse. For example,
'put this code in the Form_MouseDown Proc:

'Call ProgramFormMove(Me)

'This will allow the user to drag the form.

'Code:

'Call the ReleaseCapture API. This API is kinda complex, but what it does is
'restore the normal mouse input to the computer, kind of simulating a mouse
'proclomation. Don't pay too much attention to this.
Call ReleaseCapture
'Send the WM_SYSCOMMAND const to call the WM_MOVE constant on the form, which
'will move the form.
Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Function

Function AOLUpChat()
'This makes AOL "useable" during an upload. Example:

'AOLUpchat

'Uhh, ok. This makes AOL enabled during an upload.

'Code:

'Declare variables...
Dim AOL As Long, UpLoad As Long, Gauge As Long
Dim Upchat As Long

'Get AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'..Get the UpLoad window...
UpLoad& = FindWindowEx(AOL&, 0&, "_AOL_Modal", vbNullString)
'..Get the UpLoad progress bar...
Gauge& = FindWindowEx(UpLoad&, 0&, "_AOL_Gauge", vbNullString)

'If the progress bar is found, then make the UpChat variable equal the UpLoad
'variable.
If Gauge& <> 0& Then Upchat& = UpLoad&

'Enable AOL and...
Call EnableWindow(AOL&, 1)
'...Disable the UpLoad window with the EnableWindow API.
Call EnableWindow(Upchat&, 0)
End Function

Function AOLUnUpChat()
'This is the opposite of the AOLUpchat function. Example:

'AOLUnUpchat

'Uhh, ok. This makes AOL disabled during an upload.

'Code:

'Declare variables...
Dim AOL As Long, UpLoad As Long, Gauge As Long
Dim Upchat As Long

'Get AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'..Get the UpLoad window...
UpLoad& = FindWindowEx(AOL&, 0&, "_AOL_Modal", vbNullString)
'..Get the UpLoad progress bar...
Gauge& = FindWindowEx(UpLoad&, 0&, "_AOL_Gauge", vbNullString)

'If the progress bar is found, then make the UpChat variable equal the UpLoad
'variable.
If Gauge& <> 0& Then Upchat& = UpLoad&

'Disable AOL and...
Call EnableWindow(AOL&, 0)
'...Enable the UpLoad window with the EnableWindow API.
Call EnableWindow(Upchat&, 1)
End Function

Function StringEnCrypt(TextToEnCrypt As String)
'This function encrypts a string. Example:

'Text1.Text = StringEncrypt("Yoshi is cool")

'This will make the text in Text1 equal encrypted version of "Yoshi is cool".

'Code:

'Declare variables...
Dim Length As Long, Character As String, Holder As String

'Make the variable Length hold all integers from 1 to the end of
'the string which is stored in TextToEnCrypt$.
For Length& = 1 To Len(TextToEnCrypt$)
    'This is a new and rather complicated event. First using the Mid
    'function, we go into the TextToEnCrypt string according to which
    'number the Length& loop is at. Lets say this is the first loop..
    'meaning that Length& would equal 1. This means we go into the string
    '1 space, and from there on, move 1 more space. This means we encompass
    'one character. Now since we have this 1 character, get the ascii of it
    'using the Asc function. Then overall, since we have the ascii, subtract
    '1 from it so the ascii is changed, meaning the character is different.
    'That means we just successfully encrypted our first character. Store it
    'under the Character string.
    'NOTE: Number 1 can be changed, but 1 is used in most programs and it
    'is more "default". For serious encrypting, you could do something big
    'like 27 and the characters will be completely different. By subtracting
    'one, you would go down one. Example: If you encrypted Y, the outcome
    'would be X, because the ascii was changed down by 1.
    Character$ = Asc(Mid(TextToEnCrypt$, Length&, 1)) - 1
    'Make the Holder string equal what it was before plus the character form
    'of ascii, meaning the new character.
    Holder$ = Holder$ & Chr(Character$)
'Loop Lengt& again, and keep looping until the entire string has undergone
'this same process.
Next Length&

'Make this function equal the encrypted string so that this function can
'be used as described the in the example above.
StringEnCrypt = Holder$
End Function

Function StringDeCrypt(TextToDeCrypt As String)
'This function is the opposite from the encrypter, it decrypts
'the previously encrypted messages. Example:

'Text1.Text = StringEncrypt("Xnrghhrbnnk")

'This will make the text in Text1 equal the decrypted version of
'"Xnrghhrbnnk", which is "Yoshi is cool".

'Code:

'Declare variables...
Dim Length As Long, Character As String, Holder As String

'Make the variable Length hold all integers from 1 to the end of
'the string which is stored in TextToDeCrypt$.
For Length& = 1 To Len(TextToDeCrypt$)
    'This is a new and rather complicated event. First using the Mid
    'function, we go into the TextToDeCrypt string according to which
    'number the Length& loop is at. Lets say this is the first loop..
    'meaning that Length& would equal 1. This means we go into the string
    '1 space, and from there on, move 1 more space. This means we encompass
    'one character. Now since we have this 1 character, get the ascii of it
    'using the Asc function. Then overall, since we have the ascii, add
    '1 to it so the ascii is changed, meaning the character is different.
    'That means we just successfully decrypted our first character. Store it
    'under the Character string.
    'NOTE: When decrypting, it will only succesfully decrypt the encrypted
    'string if the number of add/subtraction is the same.
    Character$ = Asc(Mid(TextToDeCrypt$, Length&, 1)) + 1
    'Make the Holder string equal what it was before plus the character form
    'of ascii, meaning the new character.
    Holder$ = Holder$ & Chr(Character$)
'Loop Lengt& again, and keep looping until the entire string has undergone
'this same process.
Next Length&

'Make this function equal the encrypted string so that this function can
'be used as described the in the example above.
StringDeCrypt = Holder$
End Function

Function AOLMailBoxToList(TheListBox As ListBox)
'This long sub takes all your mail and puts it in a listbox. Example:

'Call AOLMailBoxToList(List1)

'This will put all your mails into List1, as long as the mailbox is open.

'Declare variables...
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim location As Long, StringText As String, Count As Long
Dim TheMail As Long, TextLength As Long, MailList As Long
    
'Using my AOLFindMailBox function, determine if their is a mailbox open
'and then get its handle.
MailBox& = AOLFindMailBox&

'If the handle is 0, then quit this function since there is no use proceeding.
If MailBox& = 0& Then Exit Function

'Find the tab control in the mailbox...
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
'Find the tab page in that control...
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
'Find the mail list itself in the tab page.
MailList& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)

'Using the LB_GETCOUNT constant, get the number of items in the mailbox.
Count& = SendMessage(MailList&, LB_GETCOUNT, 0&, 0&)

'If there are 0 mails in the mailbox, then quit, since there is no use
'proceeding.
If Count& = 0& Then Exit Function

'For TheMail variable, make it go from 0 to the amount of mails minus 1, since
'0 is counted as an item.
For TheMail& = 0 To Count& - 1
        
        'Kill locking up...
        DoEvents
        'Get the text length of the current mail...
        TextLength& = SendMessage(MailList&, LB_GETTEXTLEN, TheMail&, 0&)
        'Get the number and convert it into a numerical string...
        StringText$ = String(TextLength& + 1, 0)
        
        'Get the text of the mail and put it in StringText$.
        Call SendMessageByString(MailList&, LB_GETTEXT, TheMail&, StringText$)
        
        'Find the point where Chr(9) is, the space that comes between
        'the screename and the subject...
        location& = InStr(StringText$, Chr(9))
        'Skip from that space and get the rest, since we don't want the SN.
        location& = InStr(location& + 1, StringText$, Chr(9))
        
        'Make the StringText variable equal itself without the Chr(9)..
        StringText$ = Right(StringText$, Len(StringText$) - location&)
        'Add the currernt item to the listbox...
        TheListBox.AddItem "[" & TheMail& + 1 & "]   " & StringText$

'Loop the process.
Next TheMail&
End Function

Function AOLFindMailBox() As Long
'This function finds the AOL Mailbox and stores its handle in
'the function itself. Example:

'MyHandle& = AOLFindMailBox&

'This will get the AOLMailBox' handle and put it in MyHandle.

'Decalre variables.
Dim AOL As Long, MDI As Long, OneChild As Long
Dim TabControl As Long, TabPage As Long

'Find AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'Find the MDI Client...
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
'Find any child, no matter what it is...
OneChild& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
'Find the Tab Control in that child...
TabControl& = FindWindowEx(OneChild&, 0&, "_AOL_TabControl", vbNullString)
'Find the Tab Page in the Tab Control..
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)

'If the TabControl and TabPage were found, meaning they don't equal 0, then
'make this function equal that one child, which is the mailbox.
If TabControl& And TabPage& <> 0& Then
    AOLFindMailBox& = OneChild&
    'Then exit the function...
    Exit Function
'If the TabControl and TabPage were'nt found, then...
Else
        'Begin a loop...
        Do
            'Find another child...
            OneChild& = FindWindowEx(MDI&, OneChild&, "AOL Child", vbNullString)
            'Find the tab control in it...
            TabControl& = FindWindowEx(OneChild&, 0&, "_AOL_TabControl", vbNullString)
            'Find the tab page in the tab control...
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            'Now if it equals 0, then loop this process again, till it finds
            'the right child. If it doesn't equal 0, then make this function
            'equal the mailbox' handle and exit the function.
            If TabControl& <> 0& And TabPage& <> 0& Then
                AOLFindMailBox& = OneChild&
                Exit Function
            End If
        'Do this loop again just in case it didint find the mailbox that time
        'either.
        Loop Until OneChild& = 0&
'End that loop...
End If
'Make this function equal 0 if the mailbox was not found at all.
AOLFindMailBox& = 0&
End Function

Function AOLOpenNewMail()
'Opens your AOL mailbox. Example:

'AOLOpenNewMail

'Ok. That opens your AOL mailbox.

'Declare variables...
Dim AOL As Long, Toolbar As Long
Dim ToolBarTwo As Long, ReadMail As Long

'Find AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'Find the toolbar...
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
'Go one step into the toolbar...
ToolBarTwo& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
'Find the ReadMail icon...
ReadMail& = FindWindowEx(ToolBarTwo&, 0&, "_AOL_Icon", vbNullString)

'Click the ReadMail icon.
Call APIButtonClick(ReadMail&)
End Function

Function AOLOpenComposeMail()
'Opens your a clean AOL write a mail form. Example:

'AOLOpenComposeMail

'Ok. That opens an AOL write mail form.

'Declare variables...
Dim AOL As Long, Toolbar As Long
Dim ToolBarTwo As Long, WriteMail As Long

'Find AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'Find the toolbar...
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
'Go one step into the toolbar...
ToolBarTwo& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
'Find the ReadMail icon...
WriteMail& = FindWindowEx(ToolBarTwo&, 0&, "_AOL_Icon", vbNullString)
'Find the WriteMail icon...
WriteMail& = FindWindowEx(ToolBarTwo&, WriteMail&, "_AOL_Icon", vbNullString)

'Click the WriteMail icon.
Call APIButtonClick(WriteMail&)
End Function

Function AOLFindWriteMail() As Long
'This function finds the AOL WriteMail box and stores its handle in
'the function itself. Example:

'MyHandle& = AOLFindWriteMail&

'This will get the AOLWriteMailBox' handle and put it in MyHandle.

'Decalre variables.
Dim AOL As Long, MDI As Long, OneChild As Long
Dim TheLabel As Long

'Find AOL...
AOL& = FindWindow("AOL Frame25", vbNullString)
'Find the MDIClient...
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
'Find any child window in AOL...
OneChild& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
'The Send To: label in the write mail form.
TheLabel& = FindWindowEx(OneChild&, 0&, "_AOL_Static", vbNullString)

'If it finds the mail, then make this function equal the write mail windows
'handle.
If TheLabel& <> 0 Then
    AOLFindWriteMail& = OneChild&
    'Then exit the function.
    Exit Function
Else
    'Start the loop to find the write mail window...
    Do
        'Find any window...
        OneChild& = FindWindowEx(MDI&, OneChild&, "AOL Child", vbNullString)
        'Find the Send To: window...
        TheLabel& = FindWindowEx(OneChild&, 0&, "_AOL_Static", "Send To:")
        'If it finds the label, then leave.
        If TheLabel& <> 0& Then
            AOLFindWriteMail& = OneChild&
            Exit Function
        End If
    'Keep looping till it finds the window...
    Loop Until OneChild& = 0&
'End this loop.
End If
'Make this function equal 0 since it didint find the window.
AOLFindWriteMail& = 0&
End Function

Function AOLWriteMail(ThePerson As String, Subject As String, Message As String)
'This sends a mail in AOL. Example:

'Call AOLWriteMail("insanelaff", "stop laughing", "hahaha")

'This will send an e-mail to InsaneLaff with a subject of stop laughing and the
'message inside will say hahaha.

'Declare variables...
Dim SendTo As Long, TheSubject As Long, TheMessage As Long
Dim SendButton As Long


If AOLFindWriteMail = 0& Then Exit Function


Do: DoEvents
'Find the reciever screename textbox...
SendTo& = FindWindowEx(AOLFindWriteMail&, 0&, "_AOL_Edit", vbNullString)

'Find the subject box...
TheSubject& = FindWindowEx(AOLFindWriteMail&, SendTo&, "_AOL_Edit", vbNullString)
'Find the subject box...
TheSubject& = FindWindowEx(AOLFindWriteMail&, TheSubject&, "_AOL_Edit", vbNullString)

'Find the message box...
TheMessage& = FindWindowEx(AOLFindWriteMail&, 0&, "RICHCNTL", vbNullString)

'Find the sendbutton...
SendButton& = FindWindowEx(AOLFindWriteMail&, 0&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)
'Found the send button.
SendButton& = FindWindowEx(AOLFindWriteMail&, SendButton&, "_AOL_Icon", vbNullString)

Loop Until SendTo& And TheSubject& And TheMessage& And SendButton& <> 0&

'Catch up...
Call ProgramPause(1)
'Send the person's name to the send to box...
Call APISendMessage(SendTo&, ThePerson$)
'Send the subject to the subject box...
Call APISendMessage(TheSubject&, Subject$)
'Send the message to the message box...
Call APISendMessage(TheMessage&, Message$)
'Pause the program just to let AOL catch up...
Call ProgramPause(0.1)
'Click the send button.
Call APIButtonClick(SendButton&)
End Function

Function AOLLastChatLine()
'This function gets the last line of the chat. Example:

'Text1.Text = AOLLastChatLine

'This will make Text1 equal the last line that was said in the
'AOL chatroom.

'Decalre Variables...
Dim ChatBox As Long, ChatText As String, NextLine As String
Dim Finalize As String, TextLength As String

'Find the ChatRoom Text Box....
ChatBox& = FindWindowEx(AOLFindChatRoom&, 0&, "RICHCNTL", vbNullString)
'Get its text...
ChatText$ = APIGetText(ChatBox&)
'Go in from the end and find a space...
NextLine$ = InStrRev(ChatText$, Chr(13))
'Make this variable equal the length of all the text in the chatroom..
TextLength$ = Len(ChatText$)
'Find the position of the chr(13) by minusing the reversed position
'from the length of the whole text...
NextLine$ = Val(TextLength$) - Val(NextLine$)
'Go in from that position and put it in Finalize$...
Finalize$ = Right(ChatText$, NextLine$)
'Make this function equal the last chat line.
AOLLastChatLine = Finalize$
End Function

Function AOLWaitForMail()
'This function waits for all the mail in the mailbox to load. Example:

'AOLWaitForMail

'This will wait for all the mail to load before proceeding.

'Declare variables...
Dim AOL As Long, MDIClient As Long, MailBox As Long, TabControl As Long
Dim TabPage As Long, AOTree As Long, OneCounter As Long, TwoCounter As Long

'Open the mailbox..
AOLOpenNewMail

'Start a loop...
Do: DoEvents
    'Find the TabControl...
    TabControl& = FindWindowEx(AOLFindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    'Find the Tab page...
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    'Find the mail list...
    AOTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
'Loop until all parts were found...
Loop Until AOTree& <> 0&
'Begin a loop...
Call ProgramPause(2)
Do
    'Get the mail...
    OneCounter& = SendMessageByNum(hWndMailLB, LB_GETCOUNT, 0&, 0&)
    'Catch up...
    Call ProgramPause(0.0001)
    'Get the mail amount...
    TwoCounter& = SendMessageByNum(hWndMailLB, LB_GETCOUNT, 0&, 0&)
'Keep looping until all mail is loaded, meaning that the variables will
'be equal.
Loop Until OneCounter& = TwoCounter&
End Function

Function AOLLastChatLineSN()
'This function gets the SN from the last chat line. Example:

'Text1.Text = AOLLastChatLineSN

'This will make the text in Text1 equal the screename from the last chat line.

'Declare variables...
Dim NextLine As String, TextLength As String

'Find the colon in the last chat line...
NextLine$ = InStr(AOLLastChatLine, ":")
'Get the text from the last chat line from the beggining to where the colon is...
TextLength$ = Left(AOLLastChatLine, NextLine$)
'Make this function equal that text from beggining to the colon without the colon.
AOLLastChatLineSN = StringRemoveCharacter(TextLength$, ":")
End Function

Function AOLLastChatLineText()
'This function gets the Text from the last chat line. Example:

'Text1.Text = AOLLastChatLineText

'This will make the text in text1 equal the message from the last
'chat line.

'Declare variables...
Dim ChatBox As Long, ChatText As String, NextLine As String
Dim Finalize As String, TextLength As String, Looker As String
Dim Checker As String, TextTwo As String

'Find the chatbox...
ChatBox& = FindWindowEx(AOLFindChatRoom&, 0&, "RICHCNTL", vbNullString)
'Get the text...
ChatText$ = APIGetText(ChatBox&)
'Get the newest line in the chatroom...
NextLine$ = InStrRev(ChatText$, Chr(13))
'Get its length...
TextLength$ = Len(ChatText$)
'I forgot... I wrote this code in a hurry...
NextLine$ = Val(TextLength$) - Val(NextLine$)
'Umm dont pay attention I know your not reading this anyway...
Finalize$ = Right(ChatText$, NextLine$)
'Make TextTwo equal the last chat line...
TextTwo$ = Finalize$

'Go in and find the colon...
NextLine$ = InStr(Finalize$, ":")
'Get the text before the colon...
TextLength$ = Left(Finalize$, NextLine$)
'Remove the colon...
TextThree$ = StringRemoveCharacter(TextLength$, ":")

'Get the length of the last chat line...
Looker$ = Len(Finalize$)
'Get the next line...
NextLine$ = Val(Looker$) - Val(NextLine$)
'Do something I;m too lazy to understand my own code...
ChatText$ = Right(Finalize$, NextLine$)
'Remove the Chr(9), the big gap between the SN and the Message in the chatroom..
Checker$ = StringRemoveCharacter(ChatText, Chr(9))

'Something...
Finalize$ = Right(Checker$, Len(Checker$) - 1)
'Make this function equal that something.
AOLLastChatLineText = Finalize$
End Function

Function StringReplaceCharacter(TheString As String, LengthOfChar As Long, ToFind As String, Replacement As String)
'This function replaces a character in a string with another character.
'Example:

'Text1.Text = StringReplaceCharacter("DoGz OwN", "O", "W")

'This will make Text1.Text equal DoGz WwN

'Declare variables...
Dim Switch As Long, Storage As String, Holder As String

'Clear out the previously done record...
Holder$ = ""

'If it can't find it...
If InStr(TheString$, ToFind$) = 0 Then
    'Don't make any changes...
    StringReplaceCharacter = TheString$
    'Quit this...
    Exit Function
'Get outta hea...
End If

'For Switch make it equal the length of the string...
For Switch& = 1 To Len(TheString$)
    'In our storage, go into the string and go to the next character...
    Storage$ = Mid(TheString$, Switch&, 1)
    'Add that character...
    Holder$ = Holder$ & Storage$
    'If the storage equals the character we want to remove...
    If Storage$ = ToFind$ Then
        'Go into the holder and replace it ...
        Holder$ = Mid(Holder$, LengthOfChar&, Len(Holder$) - LengthOfChar&) & Replacement$
    End If
'Keep repeating till completely removed...
Next Switch&

'Make the function equal the entire thing complete..
StringReplaceCharacter = Holder$
End Function
