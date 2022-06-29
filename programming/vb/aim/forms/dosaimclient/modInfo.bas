Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public Const EM_GETLINECOUNT = &HBA

Public Const GWL_STYLE As Long = -16&
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWTEXT = 8

Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003

Public Const TVIF_STATE As Long = &H8

'treeview styles
Public Const TVS_HASLINES As Long = 2
Public Const TVS_CHECKBOXES = &H100
Public Const TVS_FULLROWSELECT As Long = &H1000

'treeview style item states
Public Const TVIS_CUT        As Long = &H4
Public Const TVIS_BOLD       As Long = &H10
Public Const TVIS_CHECK      As Long = &H3000
Public Const TVIS_CHECKED    As Long = &H2000
Public Const TVIS_UNCHECKED  As Long = &H1000

Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9

Public Type TVITEM
   mask As Long
   hItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256

Sub FormatBold(mdiFrm As Form, Optional vntForce As Variant)
    If IsMissing(vntForce) Then
        vntForce = False
    End If
    
   With frmSetInfo.InfoEditor
        If (IsNull(.SelBold) = True) Or (.SelBold = False) Or (vntForce = True) Then
            .SelBold = True
          
        ElseIf .SelBold = True Then
            .SelBold = False
            
        End If
    End With
End Sub

Sub FormatItalic(mdiFrm As Form, Optional vntForce As Variant)
    If IsMissing(vntForce) Then
        vntForce = False
    End If
    With frmSetInfo.InfoEditor
        If (IsNull(.SelItalic) = True) Or (.SelItalic = False) Or (vntForce = True) Then
            .SelItalic = True
           
        ElseIf .SelItalic = True Then
           
            .SelItalic = False
           
        End If
    End With
End Sub

Sub FormatUnderline(mdiFrm As Form, Optional vntForce As Variant)
    If IsMissing(vntForce) Then
        vntForce = False
    End If
   With frmSetInfo.InfoEditor
        If (IsNull(.SelUnderline) = True) Or _
            (.SelUnderline = False) Or (vntForce = True) Then
                       
            .SelUnderline = True
           
        ElseIf .SelUnderline = True Then
            
            .SelUnderline = False
            
        End If
    End With
End Sub

Sub FormatAlign(mdiFrm As Form, intIndex As Integer)
    With frmSetInfo
        Select Case intIndex
            Case 0
                frmSetInfo.InfoEditor.SelAlignment = rtfLeft
            Case 1
                 frmSetInfo.InfoEditor.SelAlignment = rtfCenter
            Case 2
                 frmSetInfo.InfoEditor.SelAlignment = rtfRight
        End Select
    End With
End Sub

Function RichToHTML(strRTF As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String
'The Conversion from RTB to Html

        Dim strHTML As String
        Dim l As Long
        Dim lTmp As Long
        Dim lRTFLen As Long
        Dim lBOS As Long 'beginning of section
        Dim lEOS As Long 'end of section
        Dim strTmp As String
        Dim strTmp2 As String
        Dim strEOS 'string To be added to End of section
        Const gHellFrozenOver = False 'always false
        Dim gSkip As Boolean 'skip To Next word/command
        Dim strCodes As String 'codes For ascii To HTML char conversion
        strCodes = "  {00}© {a9}´ {b4}« {ab}» {bb}¡ {a1}¿{bf}À{c0}à{e0}Á{c1}"
        strCodes = strCodes & "á{e1}Â {c2}â {e2}Ã{c3}ã{e3}Ä {c4}ä {e4}Å {c5}å {e5}Æ {c6}"
        strCodes = strCodes & "æ {e6}Ç{c7}ç{e7}Ð{d0}ð{f0}È{c8}è{e8}É{c9}é{e9}Ê {ca}"
        strCodes = strCodes & "ê {ea}Ë {cb}ë {eb}Ì{cc}ì{ec}Í{cd}í{ed}Î {ce}î {ee}Ï {cf}"
        strCodes = strCodes & "ï {ef}Ñ{d1}ñ{f1}Ò{d2}ò{f2}Ó{d3}ó{f3}Ô {d4}ô {f4}Õ{d5}"
        strCodes = strCodes & "õ{f5}Ö {d6}ö {f6}Ø{d8}ø{f8}Ù{d9}ù{f9}Ú{da}ú{fa}Û {db}"
        strCodes = strCodes & "û {fb}Ü {dc}ü {fc}Ý{dd}ý{fd}ÿ {ff}Þ {de}þ {fe}ß {df}§ {a7}"
        strCodes = strCodes & "¶ {b6}µ {b5}¦{a6}±{b1}·{b7}¨{a8}¸ {b8}ª {aa}º {ba}¬{ac}"
        strCodes = strCodes & "­{ad}¯ {af}°{b0}¹ {b9}² {b2}³ {b3}¼{bc}½{bd}¾{be}× {d7}"
        strCodes = strCodes & "÷{f7}¢ {a2}£ {a3}¤{a4}¥{a5}"
        strHTML = ""
        lRTFLen = Len(strRTF)
        'seek first line with text on it
        lBOS = InStr(strRTF, vbCrLf & "\deflang")
        If lBOS = 0 Then GoTo finally Else lBOS = lBOS + 2
        lEOS = InStr(lBOS, strRTF, vbCrLf & "\par")
        If lEOS = 0 Then GoTo finally


        While Not gHellFrozenOver
            strTmp = Mid(strRTF, lBOS, lEOS - lBOS)
            l = lBOS


            While l <= lEOS
                strTmp = Mid(strRTF, l, 1)


                Select Case strTmp
                    Case "{"
                    l = l + 1
                    Case "}"
                    strHTML = strHTML & strEOS
                    l = l + 1
                    Case "\" 'special code
                    l = l + 1
                    strTmp = Mid(strRTF, l, 1)


                    Select Case strTmp
                        Case "b"


                        If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                            strHTML = strHTML & "<B>"
                            strEOS = "</B>" & strEOS
                            If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                        ElseIf (Mid(strRTF, l, 7) = "bullet ") Then
                            strHTML = strHTML & "•" 'bullet
                            l = l + 6
                        Else
                            gSkip = True
                        End If
                        Case "e"


                        If (Mid(strRTF, l, 7) = "emdash ") Then
                            strHTML = strHTML & "—"
                            l = l + 6
                        Else
                            gSkip = True
                        End If
                        Case "i"


                        If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                            strHTML = strHTML & "<I>"
                            strEOS = "</I>" & strEOS
                            If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                        Else
                            gSkip = True
                        End If
                        Case "l"


                        If (Mid(strRTF, l, 10) = "ldblquote ") Then
                            strHTML = strHTML & "“"
                            l = l + 9
                        ElseIf (Mid(strRTF, l, 7) = "lquote ") Then
                            strHTML = strHTML & "‘"
                            l = l + 6
                        Else
                            gSkip = True
                        End If
                        Case "p"


                        If ((Mid(strRTF, l, 6) = "plain\") Or (Mid(strRTF, l, 6) = "plain ")) Then
                            strHTML = strHTML & strEOS
                            strEOS = ""
                            If Mid(strRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5 'catch Next \ but skip a space
                        Else
                            gSkip = True
                        End If
                        Case "r"


                        If (Mid(strRTF, l, 7) = "rquote ") Then
                            strHTML = strHTML & "’"
                            l = l + 6
                        ElseIf (Mid(strRTF, l, 10) = "rdblquote ") Then
                            strHTML = strHTML & "”"
                            l = l + 9
                        Else
                            gSkip = True
                        End If
                        Case "t"


                        If (Mid(strRTF, l, 4) = "tab ") Then
                            strHTML = strHTML & Chr$(9) 'tab
                            l = l + 3
                        Else
                            gSkip = True
                        End If
                        Case "'"
                        strTmp2 = "{" & Mid(strRTF, l + 1, 2) & "}"
                        lTmp = InStr(strCodes, strTmp2)


                        If lTmp = 0 Then
                            strHTML = strHTML & Chr("&H" & Mid(strTmp2, 2, 2))
                        Else
                            strHTML = strHTML & Trim(Mid(strCodes, lTmp - 8, 8))
                        End If
                        l = l + 2
                        Case "~"
                        strHTML = strHTML & " "
                        Case "{", "}", "\"
                        strHTML = strHTML & strTmp
                        Case vbLf, vbCr, vbCrLf 'always use vbCrLf
                        strHTML = strHTML & vbCrLf
                        Case Else
                        gSkip = True
                    End Select


                If gSkip = True Then
                    'skip everything up until the next space
                    '     or "\"


                    While ((Mid(strRTF, l, 1) <> " ") And (Mid(strRTF, l, 1) <> "\"))
                        l = l + 1
                    Wend
                    gSkip = False
                    If (Mid(strRTF, l, 1) = "\") Then l = l - 1
                End If
                l = l + 1
                Case vbLf, vbCr, vbCrLf
                l = l + 1
                Case Else
                strHTML = strHTML & strTmp
                l = l + 1
            End Select
    Wend
    lBOS = lEOS + 2
    lEOS = InStr(lEOS + 1, strRTF, vbCrLf & "\par")
    If lEOS = 0 Then GoTo finally
    strHTML = strHTML & "<br>"
Wend
finally:
RichToHTML = strHTML


End Function

Public Function GetUrlSource(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long

    'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    'if we have the handle, then start reading the web page
    If hInternet Then
        'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        'if there's more data then keep reading it into the buffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    'close the URL
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData
End Function


Public Sub MakeBold()
Dim TVI As TVITEM
Dim r As Long
Dim hitemTV As Long
Dim hwndTV As Long
Dim tmpNode As Node

   hwndTV = frmBuddyList.tvwBuddies.hwnd
   hitemTV = SendMessageLong(hwndTV, TVM_GETNEXTITEM, TVGN_CARET, 0&)
  'if a valid handle get and set the  'item's state attributes
   If hitemTV > 0 Then
      With TVI
         .hItem = hitemTV
         .mask = TVIF_STATE
         .stateMask = TVIS_BOLD
         r = SendMessageAny(hwndTV, TVM_GETITEM, 0&, TVI)
         'flip the bold mask state
         Select Case .state And TVIS_BOLD
           Case TVIS_BOLD
             .state = 0
           Case Else
             .state = TVIS_BOLD
         End Select
      End With
      r = SendMessageAny(hwndTV, TVM_SETITEM, 0&, TVI)
   End If
   Set frmBuddyList.tvwBuddies.SelectedItem = tmpNode
End Sub


Public Function FixUserInfo(MessCode As String)

MessCode = Replace(MessCode, "<HTML>", "")
MessCode = Replace(MessCode, "</HTML>", "")
MessCode = Replace(MessCode, "<HEAD>", "")
MessCode = Replace(MessCode, "</HEAD>", "")
MessCode = Replace(MessCode, "<TITLE>", "")
MessCode = Replace(MessCode, "</TITLE>", "")
MessCode = Replace(MessCode, "User Information for ", "")
MessCode = Replace(MessCode, "<BODY BGCOLOR=#CCCCCC>", "")
MessCode = Replace(MessCode, "<IMG SRC=" & Chr(34) & "admin_icon.gif" & Chr(34) & ">", "")
MessCode = Replace(MessCode, "<IMG SRC=" & Chr(34) & "dt_icon.gif" & Chr(34) & ">", "")
MessCode = Replace(MessCode, "<IMG SRC=" & Chr(34) & "free_icon.gif" & Chr(34) & ">", "")
MessCode = Replace(MessCode, "<IMG SRC=" & Chr(34) & "aol_icon.gif" & Chr(34) & ">", "")
MessCode = Replace(MessCode, "<SUP>", "")
MessCode = Replace(MessCode, "</SUP>", "")
MessCode = Replace(MessCode, "<br>", "")
MessCode = Replace(MessCode, "<H1>", "")
MessCode = Replace(MessCode, "<H2>", "")
MessCode = Replace(MessCode, "<H3>", "")
MessCode = Replace(MessCode, "<PRE>", "")
MessCode = Replace(MessCode, "</PRE>", "")
MessCode = Replace(MessCode, "<PRE=", "")
MessCode = Replace(MessCode, "</A>", "")
MessCode = Replace(MessCode, "Username :", "")
MessCode = Replace(MessCode, "<B>", "")
MessCode = Replace(MessCode, "</B>", "")
MessCode = Replace(MessCode, "      ", "")
MessCode = Replace(MessCode, m_strFormattedSN$, "")

FixUserInfo = MessCode
End Function

Public Function FixUserInfoText(MessInfo As String)

Dim dabob As String
dabob = "<br><I>Legend:</I><br><br> : Normal AIM User<br> : AOL User <br> : Trial AIM User <br> : Administrator"

MessInfo = Replace(MessInfo, "<HTML>", "")
MessInfo = Replace(MessInfo, "</HTML>", "")
MessInfo = Replace(MessInfo, "<HEAD>", "")
MessInfo = Replace(MessInfo, "</HEAD>", "")
MessInfo = Replace(MessInfo, "<TITLE>", "")
MessInfo = Replace(MessInfo, "</TITLE>", "")
MessInfo = Replace(MessInfo, "User Information for ", "")
MessInfo = Replace(MessInfo, "<BODY BGCOLOR=#CCCCCC>", "")
MessInfo = Replace(MessInfo, "<IMG SRC=" & Chr(34) & "admin_icon.gif" & Chr(34) & ">", "")
MessInfo = Replace(MessInfo, "<IMG SRC=" & Chr(34) & "dt_icon.gif" & Chr(34) & ">", "")
MessInfo = Replace(MessInfo, "<IMG SRC=" & Chr(34) & "free_icon.gif" & Chr(34) & ">", "")
MessInfo = Replace(MessInfo, "<IMG SRC=" & Chr(34) & "aol_icon.gif" & Chr(34) & ">", "")
MessInfo = Replace(MessInfo, "<SUP>", "")
MessInfo = Replace(MessInfo, "</SUP>", "")
MessInfo = Replace(MessInfo, "<H1>", "")
MessInfo = Replace(MessInfo, "<H2>", "")
MessInfo = Replace(MessInfo, "<H3>", "")
MessInfo = Replace(MessInfo, "<hr>", "")
MessInfo = Replace(MessInfo, "<PRE>", "")
MessInfo = Replace(MessInfo, "</PRE>", "")
MessInfo = Replace(MessInfo, "<PRE=", "")
'MessInfo = Replace(MessInfo, "</A>", "")
MessInfo = Replace(MessInfo, dabob, "")
MessInfo = Replace(MessInfo, "<B></B>", "")
'MessInfo = Replace(MessInfo, ":-)", "<img src=" & Chr(34) & "http://www.notthesame.net/projectaim/smile.gif" & Chr(34) & ">")
'MessInfo = Replace(MessInfo, ":-/", "<img src=" & Chr(34) & "http://www.notthesame.net/projectaim/undecided.gif" & Chr(34) & ">")
'MessInfo = Replace(MessInfo, ":-\", "<img src=" & Chr(34) & "http://www.notthesame.net/projectaim/undecided.gif" & Chr(34) & ">")
'MessInfo = Replace(MessInfo, ":-(", "<img src=" & Chr(34) & "http://www.notthesame.net/projectaim/frown.gif" & Chr(34) & ">")



FixUserInfoText = MessInfo

End Function

Public Function GetOnTime(TimeOn As String)
Dim RawTime As String

RawTime = Replace(TimeOn, "Mon ", "")
RawTime = Replace(RawTime, "Tue ", "")
RawTime = Replace(RawTime, "Wed ", "")
RawTime = Replace(RawTime, "Thu ", "")
RawTime = Replace(RawTime, "Fri ", "")
RawTime = Replace(RawTime, "Sat ", "")
RawTime = Replace(RawTime, "Sun ", "")
RawTime = Replace(RawTime, "Jan ", "1/")
RawTime = Replace(RawTime, "Feb ", "2/")
RawTime = Replace(RawTime, "Mar ", "3/")
RawTime = Replace(RawTime, "Apr ", "4/")
RawTime = Replace(RawTime, "May ", "5/")
RawTime = Replace(RawTime, "Jun ", "6/")
RawTime = Replace(RawTime, "Jul ", "7/")
RawTime = Replace(RawTime, "Aug ", "8/")
RawTime = Replace(RawTime, "Sep ", "9/")
RawTime = Replace(RawTime, "Oct ", "10/")
RawTime = Replace(RawTime, "Nov ", "11/")
RawTime = Replace(RawTime, "Dec ", "12/")
RawTime = Replace(RawTime, Year(Date), "")
GetOnTime = RawTime
End Function

Public Function RemoveDate(daString As String)
Dim RawTime As String

RawTime = Replace(daString, "Sun ", "")
RawTime = Replace(RawTime, "1/", "")
RawTime = Replace(RawTime, "2/", "")
RawTime = Replace(RawTime, "3/", "")
RawTime = Replace(RawTime, "4/", "")
RawTime = Replace(RawTime, "5/", "")
RawTime = Replace(RawTime, "6/", "")
RawTime = Replace(RawTime, "7/", "")
RawTime = Replace(RawTime, "8/", "")
RawTime = Replace(RawTime, "9/", "")
RawTime = Replace(RawTime, "10/", "")
RawTime = Replace(RawTime, "11/", "")
RawTime = Replace(RawTime, "12/", "")
RemoveDate = RawTime
End Function

Public Function Dates()
Dim i As Long
For i = 1 To 31
Dates = i
Next
End Function
