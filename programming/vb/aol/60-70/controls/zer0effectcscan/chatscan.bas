Attribute VB_Name = "ChatScan"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindParent& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function FindChild& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1&, ByVal hWnd2&, ByVal lpsz1$, ByVal lpsz2$)
Public Declare Function SendIt& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Public Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Const GetTxt = 13
Public Const GetTxtLen = 14
'
'well have fun and e-mail me with comments
'e-mail: stupid_limits@bored.com
'im: stupid limits
'
'code written by martyr
'with ideas from
'previous versions of Syn
'and PAA's 2nd chat scan.
'made specifically for vb6



Public Function Line_Sn(Line As String) As String
    
'**New to Chat Scan**
    
    'this saves time (trust me)
    Line_Sn$ = Left(Line$, InStr(Line$, ":") - 1)
    'this gets all the text to the left of the colon
    Line_Sn$ = Replace(Line_Sn$, Chr(13), "")
    'this line takes out any enter's that may be at the end of the text
End Function


Public Function Line_Msg(Line As String) As String
    
'**New to Chat Scan**
    
    'this saves time too
    Line_Msg$ = Mid(Line$, InStr(Line$, ":") + 3)
    'this gets all the text to the left, plus 3, of the colon
    Line_Msg$ = Replace(Line_Msg$, Chr(13), "")
    'this takes out any added enter's that may be at the end
End Function


Public Function ChatBox5() As Long
'**TAKEN FROM ORIGINAL CHAT SCAN**


'this is used to find the first RICHCNTL in the aol chat

    Dim AOL As Long
    Dim AoMDI As Long
    Dim AoChild As Long
    Dim AoList As Long
    Dim AoTxt1 As Long
    Dim AoTxt2 As Long
    
    AOL& = FindParent&("aol frame25", vbNullString)
    'check if aol is open
        If AOL& = 0& Then ChatBox5 = 0&: Exit Function
        'if aol isn't open then get outta here
        AoMDI& = FindChild&(AOL&, 0&, "mdiclient", vbNullString)
        'find aol's mdiclient
            AoChild& = FindChild&(AoMDI&, 0&, "aol child", vbNullString)
            'find the topmost child
                AoList& = FindChild&(AoChild&, 0&, "_aol_listbox", vbNullString)
                'find a listbox
                AoTxt1& = FindChild&(AoChild&, 0&, "richcntl", vbNullString)
                'find a textbox
                AoTxt2& = FindChild&(AoChild&, AoTxt1&, "richcntl", vbNullString)
                'find another text box
                    If AoList& <> 0& And AoTxt1& <> 0& And AoTxt2& <> 0& Then ChatBox5& = AoTxt1&: Exit Function
                    'if everything is found then set ChatBox to the value of the first textbox on the aol child
                        While AoChild& <> 0&
                        'if the child couldn't be found this will loop through all the children looking for the same stuff as before
                            DoEvents
                            AoChild& = FindChild&(AoMDI&, 0&, "aol child", vbNullString)
                                AoList& = FindChild&(AoChild&, 0&, "_aol_listbox", vbNullString)
                                AoTxt1& = FindChild&(AoChild&, 0&, "richcntl", vbNullString)
                                AoTxt2& = FindChild&(AoChild&, AoTxt1&, "richcntl", vbNullString)
                                    If AoList& <> 0& And AoTxt1& <> 0& And AoTxt2& <> 0& Then ChatBox5& = AoTxt1&: Exit Function
                                    'if we finally find what we're looking for then this will do what the if statement does above
                        Wend
            ChatBox5& = 0&
            'just in case the chat isn't open this will take care of any mishaps
            
End Function
Public Function ShorterText() As String
'i don't need to get the last 10 lines kuz my scan
'is a bit faster then PAA's but 5 is not enough
'so we're gonna play with the last 7 lines
'(mainly kuz it's my fav number)

'**New to Chat Scan**

'**TAKEN FROM PAA's SCAN**
    Dim Count As Integer
    Dim Txt As String
    Dim TxtBrk As Long
    Dim TxtLines As String
    
    Txt = GetText(ChatBox)
    If InStr(Txt, "Link -1") = 0 Then GoTo skip:
Txt = ReplaceString(Txt, Mid(Txt, InStr(Txt, "Link -1"), Len(Txt)), "")
skip:
    'set Txt to the string in the chat room
        For Count% = 0 To 6
        'loop through the text 7 times getting a new last line each time
            If Txt = "" Then Exit For
            'if there is no text then exit the for statement
            TxtBrk = InStrRev(Txt, Chr$(13))
            'look for the last enter
            TxtLines = Mid(Txt, TxtBrk, Len(Txt)) & TxtLines
            'get the last line of text
            Txt = Mid(Txt, 1, Int(TxtBrk - 1))
            'set Txt to equal all the text except the last line you just took out
        Next Count%
        'continue through the for statement til the number 6 is reached
    ShorterText = TxtLines
    'set ShorterText to the string of TxtLines
    
End Function

Public Function Line_Count(Txt As String) As Integer

'taken from Pat's site
'**New to Chat Scan**
    
    Dim Kount As Integer
    Dim Text As String
    
    Text = Txt
    'set Text to the string of Txt
    Kount% = 1
    'set Kount to the value of 1
    While InStr(Text, Chr$(13)) <> 0&
    'loop this while there is an enter in the string
        Kount = Kount + 1
        'each time kount will go up 1
        Text = Mid$(Text, InStr(Text, Chr$(13)) + 1)
        'the text will be set to a new string so we don't just loop through the same old thing
    Wend
        Line_Count% = Kount%
        'all the enters have been accounted for so we need to set the function to Kount for a number
End Function

Public Function Line_Text(Txt As String, Line As Integer) As String
'FROM THE ORIGINAL CHAT SCAN
    Dim findchar As Integer
    Dim TheChar As String
    Dim TheChars As String
    Dim TempNum As Integer
    Dim TheText As String
    Dim TextLength As String
    Dim TheCharsLength As Integer
    Dim Text As String

    Text = Txt
    'set Text to equal Txt
    TextLength = Len(Text$)
    'get the text length
    For findchar% = 1 To TextLength
        'go through the characters 1 by 1
        TheChar$ = Mid(Text$, findchar%, 1)
        'get the current character
        TheChars$ = TheChars$ & TheChar$
        'have TheChars equal themselves and the new character
            If TheChar$ = Chr(13) Then
            'if we find an enter character then...
                TempNum% = TempNum% + 1
                'keep track of the number of lines we're going through
                TheCharsLength% = Len(TheChars)
                'get the current length of the characters we have
                TheText$ = Mid(TheChars$, 1, TheCharsLength% - 1)
                'set TheText to equal our new text so far
                If Line% = TempNum% Then GoTo skipit
                'if the line we wanted to stop on is found then go to SkipIt
                TheChars = ""
                'set TheChars value back to nothing
            End If
    Next findchar%
        Line_Text$ = TheChars$
        'this will simply get the last line, if the number of lines is too high
    Exit Function

skipit:
    TheText$ = Replace(TheText$, Chr(13), "")
    'take out any unwanted enters
    Line_Text$ = TheText$
    'set the funtion TheText variable
    
End Function
