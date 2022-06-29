Attribute VB_Name = "GetChat"
Option Explicit

'Version 1.0
'For Aol 5.0
'by K¡m0
'the maker of Listboxes exposed

'-Intro to why I did this
'''''''''''''''''''''''''''''''''''''''''''''
'   This is Free to whome ever wants it I   '
'need not take any credit for this bas file '
'                                           '
'   Dos said that he would rather code      '
'everything himself.  Well now you have the '
'same chance to do so.  I hope that you will'
'learn from this, and not just use it. I    '
'have tryed to document the uncommon parts. '
'basicly the non subclassing parts.         '
'                                           '
'   I have not searched knk's website for   '
'Last line of chat code I am only hopping   '
'that if there is one my way of coding is   '
'faster.                                    '
'''''''''''''''''''''''''''''''''''''''''''''

'-Greets
'''''''''''''''''''''''''''''''''''''''''''''
'I must send out some greets to             '
'Pat or JK's API Spy It spead up the process'
'                                           '
'Just Remember there is no SPOON!!!!        '
'''''''''''''''''''''''''''''''''''''''''''''

'-Discription
'''''''''''''''''''''''''''''''''''''''''''''
'Find_ChatRoom: Does what is4 says          '
'               (will return the handle).   '
'-Code: Msgbox Find_ChatRoom                '
'-------------------------------------------'
'SendChat: This will send the text you      '
'          spesify to the chat room.        '
'-Code: Call SendChat("Your Text")          '
'-Code: call sendChat(Text1)                '
'-------------------------------------------'
'Text_ChatTitle: Gets Chat room title       '
'                                           '
'-Code: Text1 = Text_ChatTitle              '
'-------------------------------------------'
'Text_GetChat:  This will get everything    '
'             that is said in the chat room '
'-Code: Text1 = Text_GetChat                '
'-------------------------------------------'
'Text_LastLine:  This will get last line    '
'                of chat                    '
'-Code: Text1 = Text_LastLine(Text_GetChat) '
'-------------------------------------------'
'Text_LastSN:  This will get SN for the last'
'              person to say somthing in the'
'              chat                         '
'-Code: Text1 = Text_LastSN(Text_GetChat)   '
'-------------------------------------------'


Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_CHILD = 5

Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const WM_SETTEXT = &HC
Const WM_CHAR = &H102

Public Function Find_ChatRoom() As Long
Dim lngAOLframe As Long, lngMDIclient As Long, lngAOLchild As Long

lngAOLframe = FindWindow("aol frame25", vbNullString)
lngMDIclient = FindWindowEx(lngAOLframe, 0&, "mdiclient", vbNullString)
lngAOLchild = FindWindowEx(lngMDIclient, 0&, "aol child", vbNullString)

Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long

FindOtherWin = GetWindow(lngAOLchild, GW_HWNDFIRST)

Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid2 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_combobox", vbNullString)
       Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid6 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_image", vbNullString)
       Winkid9 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       If (Winkid1 <> 0) And _
          (Winkid2 <> 0) And _
          (Winkid3 <> 0) And _
          (Winkid4 <> 0) And _
          (Winkid5 <> 0) And _
          (Winkid6 <> 0) And _
          (Winkid7 <> 0) And _
          (Winkid8 <> 0) And _
          (Winkid9 <> 0) Then
              Find_ChatRoom = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
Find_ChatRoom = 0
End Function

Public Function Text_ChatTitle() As String
    Dim lngTitle  As Long
    lngTitle = Find_ChatRoom
    Dim TheText As String, TL As Long
    TL = SendMessageLong(lngTitle, WM_GETTEXTLENGTH, 0&, 0&)
    TheText = String(TL + 1, " ")
    Call SendMessageByString(lngTitle, WM_GETTEXT, TL + 1, TheText)
    Text_ChatTitle = Left(TheText, TL)
End Function

Public Function Text_GetChat() As String
    Dim aolchild As Long, richcntl As Long
    
    aolchild = Find_ChatRoom
    richcntl = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
    
    Dim TheText As String, TL As Long
    
    TL = SendMessageLong(richcntl, WM_GETTEXTLENGTH, 0&, 0&)
    TheText = String(TL + 1, " ")
    
    Call SendMessageByString(richcntl, WM_GETTEXT, TL + 1, TheText)
    
    Text_GetChat = Left(TheText, TL)
End Function

Public Function Text_LastLine(strTexT As String) As String
    'setup half A**ED error handleing
    On Error GoTo FunctionError:
    'setup the Counter Variable
    Dim intCount As Integer
    'setup a temp. place to hold the single letter/number
    Dim strString As String
    
    'this is a for next loop but insted of setpping
    '0 to lenght of text it will go in reverse
    'because we told it to go negitive 1
    For intCount = Len(strTexT) To 0 Step -1
        'strString is holding out single leter/number
        strString = Mid$(strTexT, intCount, 1)

        'We want to see if that letter/number is equal to chr number 9
        If strString = Chr(9) Then
            'if that is true then get right of chr number 9
            'and that will be the last line of chat
            Text_LastLine = Right(strTexT, Len(strTexT) - intCount)
            Exit Function
        End If
    Next intCount


FunctionError:
    Text_LastLine = "''NO TEXT or Room no Found''"
End Function

Public Function Text_LastSN(strTexT As String) As String
    'setup half A**ED error handleing
    On Error GoTo FunctionError:
    'setup the Counter Variable
    Dim intCount As Integer
    'setup a temp. place to hold the single letter/number
    Dim strString As String
    'setup a temp place to hold the lastline of
    'chat and the screen name
    Dim strbuffer As String
    
    'this is a for next loop but insted of setpping
    '0 to lenght of text it will go in reverse
    'because we told it to go negitive 1
    For intCount = Len(strTexT) To 0 Step -1
        'strString is holding out single leter/number
        strString = Mid$(strTexT, intCount, 1)

        'We want to see if that letter/number is equal to chr number 9
        If strString = Chr(13) Then
            'if that is true then get right of chr number 9
            'and that will be the last line of chat
            strbuffer = Right(strTexT, Len(strTexT) - intCount)
            'this seperates the SN from the lastline of chat
            Text_LastSN = Left(strbuffer, InStr(strbuffer, ":") - 1)
            Exit Function
        End If
    Next intCount


FunctionError:
    Text_LastSN = "''NO TEXT or Room not found''"
End Function

Public Sub SendChat(strTexT As String)
    
    Dim aolchild As Long
    Dim richcntl As Long
    
    aolchild = Find_ChatRoom
    richcntl = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
    richcntl = FindWindowEx(aolchild, richcntl, "richcntl", vbNullString)
    
    Call SendMessageByString(richcntl, WM_SETTEXT, 0&, strTexT)
    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
End Sub
