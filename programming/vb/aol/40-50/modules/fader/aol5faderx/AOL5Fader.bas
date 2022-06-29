Attribute VB_Name = "AOL5Fader"
' AOL5Fader 1.0 bas by PAT or JK
' website: www.patorjk.com
' made for: the new aol chat rooms

' This bas file allows you to send faded text
' to the new aol chat rooms. It also lets you
' pick any color to be the two colors to fade.

' The way it works is that instead of fading
' each character, it fades the text in groups of
' characters.

' Example on how to use:
' Call SendFadedText(Text1, RGB(255,0,0), RGB(0,0,255))

' API window finding code was generated with the 4.0 version of my api spy.

' special thanks to helix for helping out with
' the idea

Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const VK_SPACE = &H20

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
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Sub SendFadedText(text As String, Color1 As Long, Color2 As Long, Optional errorcheck As Boolean = False)
    ' This sub lets you send up to 49 faded characters to a chat room
    Dim i As Integer, LaColor As Long, NumAt As Integer, HTMLString As String
    Dim NewString As String, FadeNum As Integer, TheStep As Integer
    If Len(text) > 49 Then
        MsgBox "Error: Too much text.", 16, "Error"
        Exit Sub
    End If
    If Len(text) > 40 Then
        FadeNum = 7
    Else
        FadeNum = 6
    End If
    TheStep = FadeNum
    If Len(text) < 6 Then
        HTMLString = "<font color=#" & GetRed(Color1) & GetGreen(Color1) & GetBlue(Color1) & ">"
        NewString = HTMLString & Mid(text, 1)
    Else
        For i = 1 To Len(text) Step TheStep
            NumAt = NumAt + 1
            LaColor = GetFadedColor(Color1, Color2, NumAt, FadeNum)
            HTMLString = "<font color=#" & GetRed(LaColor) & GetGreen(LaColor) & GetBlue(LaColor) & ">"
            NewString = NewString & HTMLString & Mid(text, i, TheStep)
        Next
    End If
    Call chatsend(NewString, errorcheck)
End Sub

Private Function GetFadedColor(c1 As Long, c2 As Long, FN As Integer, FS As Integer) As Long
    Dim i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, cx1!, cx2!, cx3!
    
    ' get the red, green, and blue values out of the different
    ' colors
    red1% = (c1 And 255)
    green1% = (c1 \ 256 And 255)
    blue1% = (c1 \ 65536 And 255)
    red2% = (c2 And 255)
    green2% = (c2 \ 256 And 255)
    blue2% = (c2 \ 65536 And 255)
    
    ' get the step of the color changing
    pat1 = (red2% - red1%) / FS
    pat2 = (green2% - green1%) / FS
    pat3 = (blue2% - blue1%) / FS

    ' set the cx variables at the starting colors
    cx1 = red1%
    cx2 = green1%
    cx3 = blue1%

    ' loop till you reach the faze you are at in the fading
    For i% = 1 To FN
        cx1 = cx1 + pat1
        cx2 = cx2 + pat2
        cx3 = cx3 + pat3
    Next
    GetFadedColor = RGB(cx1, cx2, cx3)
End Function

Private Function GetRed(c1 As Long) As String
    ' returns the amount of red in the color (in html form)
    GetRed = Cov(Hex(c1 And 255))
End Function

Private Function GetGreen(c1 As Long) As String
    ' returns the amount of green in the color (in html form)
    GetGreen = Cov(Hex(c1 \ 256 And 255))
End Function

Private Function GetBlue(c1 As Long) As String
    ' returns the amount of blue in the color (in html form)
    GetBlue = Cov(Hex(c1 \ 65536 And 255))
End Function

Private Function Cov(txt As String) As String
    ' adds a "0" to the front of the number if it's lacking two digits
    Do While Len(txt) < 2
        txt = "0" & txt
    Loop
    Cov = txt
End Function

Public Function findaol5chat() As Long
    ' This function finds the aol5 chat room window
    Dim aolframe&
    Dim mdiclient&
    Dim aolchild&
    aolframe& = FindWindow("aol frame25", vbNullString)
    mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
    aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
    Dim Winkid1&, Winkid2&, Winkid3&, Winkid4&, Winkid5&, Winkid6&, Winkid7&, Winkid8&, Winkid9&, FindOtherWin&
    FindOtherWin& = GetWindow(aolchild&, GW_HWNDFIRST)
    Do While FindOtherWin& <> 0
           DoEvents
           Winkid1& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
           Winkid2& = FindWindowEx(FindOtherWin&, 0&, "richcntl", vbNullString)
           Winkid3& = FindWindowEx(FindOtherWin&, 0&, "_aol_combobox", vbNullString)
           Winkid4& = FindWindowEx(FindOtherWin&, 0&, "_aol_icon", vbNullString)
           Winkid5& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
           Winkid6& = FindWindowEx(FindOtherWin&, 0&, "richcntl", vbNullString)
           Winkid7& = FindWindowEx(FindOtherWin&, 0&, "_aol_icon", vbNullString)
           Winkid8& = FindWindowEx(FindOtherWin&, 0&, "_aol_image", vbNullString)
           Winkid9& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
           If (Winkid1& <> 0) And (Winkid2& <> 0) And (Winkid3& <> 0) And (Winkid4& <> 0) And (Winkid5& <> 0) And (Winkid6& <> 0) And (Winkid7& <> 0) And (Winkid8& <> 0) And (Winkid9& <> 0) Then
                  findaol5chat = FindOtherWin&
                  Exit Function
           End If
           FindOtherWin& = GetWindow(FindOtherWin&, GW_HWNDNEXT)
    Loop
    findaol5chat = 0
End Function

Public Sub chatsend(txt As String, Optional errorcheck As Boolean = False)
    ' This is a pretty cool chatsending sub, it does a lot that others
    ' don't. Like letting you catch errors and waiting till the text is
    ' gone before sending more.
    Dim aolframe&, mdiclient&, aolchild&, richcntl&, aolicon&
    Dim aolmsgbox&, button&, TheText$, TL As Long, errormsg As Boolean
    aolframe& = FindWindow("aol frame25", vbNullString)
    mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
    ' find the chat room
    aolchild& = findaol5chat
    richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
    richcntl& = FindWindowEx(aolchild&, richcntl&, "richcntl", vbNullString)
    Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, txt)
    
    If richcntl& = 0 Then
       MsgBox "Error: Cannot find window.", 16, "Error"
       Exit Sub
    End If
    
    aolchild& = findaol5chat
    aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
    Call SendMessageLong(aolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(aolicon&, WM_LBUTTONUP, 0&, 0&)

    ' loop until the text is no longer in the sending box
    Do
        TL = SendMessageLong(richcntl&, WM_GETTEXTLENGTH, 0&, 0&)
        TheText$ = String(TL + 1, " ")
        Call SendMessageByString(richcntl&, WM_GETTEXT, TL + 1, TheText$)
        TheText$ = Left(TheText$, TL)
    Loop Until TheText$ = ""
    DoEvents
    
    If errorcheck = True Then
        ' check for an aol error message box
        Timeout 0.5
        aolmsgbox& = FindWindow("#32770", vbNullString)
        button& = FindWindowEx(aolmsgbox&, 0&, "button", vbNullString)
        If button& <> 0 Then
            Do
                aolmsgbox& = FindWindow("#32770", vbNullString)
                button& = FindWindowEx(aolmsgbox&, 0&, "button", vbNullString)
                Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
                DoEvents
            Loop Until button& <> 0
            errormsg = True
        End If
    End If
End Sub

Public Sub Timeout(duration As Double)
    Dim hCurrent As Long
    hInterval = hInterval * 1000
    hCurrent = GetTickCount
    Do While GetTickCount - hCurrent < Val(hInterval)
    DoEvents
    Loop
End Sub
