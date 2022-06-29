Attribute VB_Name = "modAIM"
Option Explicit
Global lastgroup

Global ServerSentCount

'###########################################################################################
'# Visual Basic 6 AOL Instant Messenger Example                                            #
'#                                                                                         #
'#   This example is in no way intended to be considered a releasable client. Instead it   #
'#   it has been built to serve as an example only. Not all of the protocol is handled.    #
'#   Several features such as an away option as well as others have not been implimented.  #
'#                                                                                         #
'#   The protocol used to write this project was released publicly with their Tik client.  #
'#   Although AOL later removed this client, it can still be found at                      #
'#   http://irc.themes.org.                                                                #
'#                                                                                         #
'#   Thanks goes out to Pre (pre@dosfx.com) for his influence on this project. You can     #
'#   his website at www.dosfx.com/~pre.                                                    #
'#                                                                                         #
'#   All questions and comments are welcome.                                               #
'#                                                                                         #
'#   Author:  Chad J. Cox (aka dos)                                                        #
'#   Email:   chad@dosfx.com                                                               #
'#   WebSite: www.dosfx.com                                                                #
'#   AIM:     dosfx or dosrox                                                              #
'###########################################################################################
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1

'the following constants are a prebuilt rtf header. i added the fonts and colors to be
'used in this example. i chose this method to make updating the rich edit controls fast
'and easy
Public Const RTF_HEADER As String = "{\rtf1\ansi\deff0\deftab720"
Public Const RTF_FONT_TABLE As String = "{\fonttbl{\f0\fswiss Arial;}}"
Public Const RTF_COLOR_TABLE As String = "{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red255\green0\blue0;\red0\green130\blue0;\red0\green0\blue130;}"
Public Const RTF_START_TEXT As String = "\viewkind4\uc1\pard\cf1\lang1033"
Public Const RTF_START As String = RTF_HEADER & vbCrLf & RTF_FONT_TABLE & vbCrLf & RTF_COLOR_TABLE & vbCrLf & RTF_START_TEXT

Public m_lngLocalSeq As Long      'local sequence
Public m_lngServerSeq As Long     'incoming sequence (not really important)
Public m_strScreenName As String  'lower case screen name with no spaces
Public m_strPassword As String    'account password
Public m_strFormattedSN As String 'screen name formatted by server for display
Public m_strMode As String        'permit/deny mode
Public m_strPDList As String      'permit/deny list

Public strSoundSignOn As String   'buddy in sound
Public strSoundSignOff As String  'buddy out sound
Public strSoundFirstIM As String  'first im sound
Public strSoundIMIn As String     'im in sound
Public strSoundIMOut As String    'im out sound

Public strInviteBuddies As String 'list of buddies to send chat invite to
Public strInviteRoom As String    'chat room name for invite (not id)
Public strInviteMessage As String 'message to send with invite

Public MeAway As Boolean

Public Sub Main()
  frmDebug.Show
  frmSignOn.Show
End Sub

Public Sub SendProc(lngFrame As Long, ByVal strData As String)
  'this procedure sends data to the aim server
  Dim lngSeqHi As Long, lngSeqLo As Long, strOut As String
  Dim lngLen As Long, lngLenHi As Long, lngLenLo As Long
  'the flap header is built here. see the protocol documentation for an explanation on this.
  
  If Len(strData) > 2048 Then Exit Sub ' Disconnects if bigger than 2k (client->server)
  
  m_lngLocalSeq& = m_lngLocalSeq& + 1
  If m_lngLocalSeq& > 65535 Then
    m_lngLocalSeq& = 0
  End If
  lngSeqHi& = Hi(m_lngLocalSeq&)
  lngSeqLo& = Lo(m_lngLocalSeq&)
  lngLen& = Len(strData$)
  lngLenHi& = Hi(lngLen&)
  lngLenLo& = Lo(lngLen&)
  strOut$ = "*" & Chr(lngFrame) & Chr(lngSeqLo&) & Chr(lngSeqHi&) & Chr(lngLenLo&) & Chr(lngLenHi&) & strData$
  If frmSignOn.wskAIM.state = sckConnected Then
    'we check here for a connection to avoid checking on each form.
    Call DoDebug("OUT: " & lngFrame & " " & lngSeqLo& & " " & lngSeqHi& & " " & lngLenLo& & " " & lngLenHi& & " " & Replace(Right(strOut$, Len(strOut$) - 6), Chr(0), "\0"))
    frmSignOn.wskAIM.SendData strOut$
  End If
End Sub

Public Sub RTFUpdate(rtfOut As RichTextBox, strUpdate As String)
  Dim strRTF As String
  strRTF$ = RTF_START & strUpdate$ & "}"
  rtfOut.SelStart = Len(rtfOut.Text)
  rtfOut.SelRTF = strRTF$

  
  
  If Asc(Left$(rtfOut.Text, 1)) = 13 Then
  rtfOut.SelStart = 1
  rtfOut.SelLength = 2
  rtfOut.SelText = ""
  
  End If
  
End Sub

Public Sub DoDebug(strData As String)
  'this procedure is made to make displaying to the debug window easier
  If frmDebug.Visible = True Then
    frmDebug.rtfDebug.SelStart = Len(frmDebug.rtfDebug.Text)
    frmDebug.rtfDebug.SelColor = vbWhite
    frmDebug.rtfDebug.SelText = vbCrLf & strData$
  End If
End Sub

'the following three procedures are used to help create the sequence numbers for the
'flap headers being sent to the aim toc server.
Public Function MakeLong(lngHi As Long, lngLo As Long) As Long
  MakeLong& = lngLo& * 256 + lngHi&
End Function

Public Function Lo(lngVal As Long) As Long
  Lo& = Fix(lngVal& / 256)
End Function

Public Function Hi(lngVal As Long) As Long
  Hi& = lngVal& Mod 256
End Function

Public Function EncryptPW(ByRef strPass As String) As String
  'this is a simple xor encryption used to encrypt the aim password. the roasting string
  'is "Tic/Toc"
  Dim arrTable() As Variant, strEncrypted As String
  Dim lngX As Long, strHex As String
  arrTable = Array("84", "105", "99", "47", "84", "111", "99")
  strEncrypted$ = "0x"
  For lngX& = 0 To Len(strPass$) - 1
    strHex$ = Hex(Asc(Mid(strPass$, lngX& + 1, 1)) Xor CLng(arrTable((lngX& Mod 7))))
    If CLng("&H" & strHex$) < 16 Then strEncrypted$ = strEncrypted$ & "0"
    strEncrypted$ = strEncrypted$ & strHex$
  Next
  EncryptPW$ = LCase(strEncrypted$)
End Function

Public Function KillHTML(ByVal strIn As String) As String
  'for the sake of this example, i chose not to try converting html to rtf. this method
  'is not perfect. should this have been a real client, i would have chosen to convert
  'the html to rtf. however, for the sake of this example, i chose just to remove as much
  'html as i could.
  Dim lngLen As Long, lngFound As Long, lngEnd As Long
  Dim strLeft As String, strRight As String
  strIn$ = Replace(strIn$, "<HTML>", "")
  strIn$ = Replace(strIn$, "</HTML>", "")
  strIn$ = Replace(strIn$, "<SUP>", "")
  strIn$ = Replace(strIn$, "</SUP>", "")
  strIn$ = Replace(strIn$, "<HR>", "")
  strIn$ = Replace(strIn$, "<H1>", "")
  strIn$ = Replace(strIn$, "<H2>", "")
  strIn$ = Replace(strIn$, "<H3>", "")
  strIn$ = Replace(strIn$, "<PRE>", "")
  strIn$ = Replace(strIn$, "</PRE>", "")
  strIn$ = Replace(strIn$, "<PRE=", "")
  strIn$ = Replace(strIn$, "<B>", "")
  strIn$ = Replace(strIn$, "</B>", "")
  strIn$ = Replace(strIn$, "<U>", "")
  strIn$ = Replace(strIn$, "</U>", "")
  strIn$ = Replace(strIn$, "<I>", "")
  strIn$ = Replace(strIn$, "</I>", "")
  strIn$ = Replace(strIn$, "<FONT>", "")
  strIn$ = Replace(strIn$, "</FONT>", "")
  strIn$ = Replace(strIn$, "<BODY>", "")
  strIn$ = Replace(strIn$, "</BODY>", "")
  strIn$ = Replace(strIn$, "<BR>", "")
  strIn$ = Replace(strIn$, "</A>", "")
  lngLen& = Len(strIn$)
  lngFound& = InStr(strIn$, "<BODY ")
  Do While lngFound& <> 0
    lngEnd& = InStr(lngFound&, strIn$, ">")
    If lngEnd& <> 0 Then
      strLeft$ = Left(strIn$, lngFound& - 1)
      strRight$ = Right(strIn$, lngLen& - lngEnd&)
      strIn$ = strLeft$ & strRight$
      lngLen& = Len(strIn$)
    End If
    lngFound& = InStr(lngFound& + 1, strIn$, "<BODY ")
  Loop
  lngFound& = InStr(strIn$, "<A ")
  Do While lngFound& <> 0
    lngEnd& = InStr(lngFound&, strIn$, ">")
    If lngEnd& <> 0 Then
      strLeft$ = Left(strIn$, lngFound& - 1)
      strRight$ = Right(strIn$, lngLen& - lngEnd&)
      strIn$ = strLeft$ & strRight$
      lngLen& = Len(strIn$)
    End If
    lngFound& = InStr(lngFound& + 1, strIn$, "<A ")
  Loop
  lngFound& = InStr(strIn$, "<FONT ")
  Do While lngFound& <> 0
    lngEnd& = InStr(lngFound&, strIn$, ">")
    If lngEnd& <> 0 Then
      strLeft$ = Left(strIn$, lngFound& - 1)
      strRight$ = Right(strIn$, lngLen& - lngEnd&)
      strIn$ = strLeft$ & strRight$
      lngLen& = Len(strIn$)
    End If
    lngFound& = InStr(lngFound& + 1, strIn$, "<FONT ")
  Loop
  strIn$ = Replace(strIn$, "&amp;", "&")
  strIn$ = Replace(strIn$, "&lt;", "<")
  KillHTML$ = strIn$
End Function

Public Function Normalize(ByVal strIn As String) As String
  'most strings sent to the aim toc server need to be normalized. this procedure formats
  'the strings as necessary.
  strIn$ = Replace(strIn$, "\", "\\")
  strIn$ = Replace(strIn$, "$", "$")
  strIn$ = Replace(strIn$, Chr(34), "\" & Chr(34))
  strIn$ = Replace(strIn$, "(", "\(")
  strIn$ = Replace(strIn$, ")", "\)")
  strIn$ = Replace(strIn$, "[", "\[")
  strIn$ = Replace(strIn$, "]", "\]")
  strIn$ = Replace(strIn$, "{", "\{")
  strIn$ = Replace(strIn$, "}", "\}")
  Normalize$ = strIn$
End Function

Public Function ExistsInTree(tvw As TreeView, ByVal strItem As String, Optional blnStartEdit As Boolean = False, Optional blnDelete As Boolean = False, Optional strReplaceWith As String = "") As Boolean
  'this procedure is used to handle the buddylist treeviews.
  Dim lngDo As Long, blnExists As Boolean
  blnExists = False
  strItem$ = LCase(Replace(strItem$, " ", ""))
  For lngDo& = 1 To tvw.Nodes.Count
    If strItem$ = LCase(Replace(tvw.Nodes.Item(lngDo&).Text, " ", "")) Then
      blnExists = True
      If blnStartEdit = True Then
        tvw.SetFocus
        tvw.Nodes.Item(lngDo&).Selected = True
        tvw.StartLabelEdit
      End If
      If blnDelete = True Then
        tvw.Nodes.Remove lngDo&
      End If
      If strReplaceWith$ <> "" Then
        tvw.Nodes.Item(lngDo&).Text = strReplaceWith$
      End If
      Exit For
    End If
  Next
  ExistsInTree = blnExists
End Function

Public Function FixRTF(ByVal strRTF As String) As String
  'since we are updating with rtf, it is important to format some of our strings in order
  'to keep our rich text from showing up as rtf code.
  strRTF$ = Replace(strRTF$, "\", "\\")
  strRTF$ = Replace(strRTF$, "}", "\}")
  strRTF$ = Replace(strRTF$, "{", "\{")
  FixRTF$ = strRTF$
End Function

Public Function WriteINIString(strSection As String, strKeyName As String, strValue As String, strFile As String) As Long
  Dim lngStatus As Long
  lngStatus& = WritePrivateProfileString(strSection, strKeyName, strValue, strFile)
  WriteINIString& = (lngStatus& <> 0)
End Function

Public Function GetINIString(strSection As String, strKeyName As String, strFile As String, Optional strDefault As String = "") As String
  Dim strBuffer As String * 256, lngSize As Long
  lngSize& = GetPrivateProfileString(strSection$, strKeyName$, strDefault$, strBuffer$, 256, strFile$)
  GetINIString$ = Left$(strBuffer$, lngSize&)
End Function

Public Function FileExists(strFileName As String) As Boolean
  Dim intLen As Integer
  If strFileName$ <> "" Then
    intLen% = Len(Dir$(strFileName$))
    FileExists = (Not err And intLen% > 0)
  Else
    FileExists = False
  End If
End Function

Public Sub LoadBuddies(strScreenName As String)
  'this procedure is used to load our buddies from the ini to our treeview.
  Dim strBuffer As String * 600, lngSize As Long, arrBuddies() As String, lngDo As Long
  Dim nod() As Node, intGroup As Integer
  lngSize& = GetPrivateProfileString(strScreenName$, "buddylist", "", strBuffer$, 600, App.Path & "\aim.ini")
  arrBuddies$ = Split(Left$(strBuffer$, lngSize&), Chr(1))
  frmBuddyList.tvwSetup.Nodes.Clear
  For lngDo& = LBound(arrBuddies$) To UBound(arrBuddies$)
    ReDim Preserve nod(1 To frmBuddyList.tvwSetup.Nodes.Count + 1)
    If arrBuddies$(lngDo&) <> "" Then
      If Left(arrBuddies$(lngDo&), 1) = "g" Then
        Set nod(frmBuddyList.tvwSetup.Nodes.Count) = frmBuddyList.tvwSetup.Nodes.Add(, , , Right(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 1)
        intGroup% = frmBuddyList.tvwSetup.Nodes.Count
      Else
        If frmBuddyList.tvwSetup.Nodes.Count > 0 Then
          Set nod(frmBuddyList.tvwSetup.Nodes.Count) = frmBuddyList.tvwSetup.Nodes.Add(nod(intGroup%), tvwChild, , Right(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 3)
          nod(frmBuddyList.tvwSetup.Nodes.Count).EnsureVisible
        End If
      End If
    End If
  Next
End Sub

Public Sub SaveBuddies(strScreenName As String)
  'saves the buddies from the treeview to the ini.
  Dim strBuddies As String, lngDo As Long
  For lngDo& = 1 To frmBuddyList.tvwSetup.Nodes.Count
    If frmBuddyList.tvwSetup.Nodes.Item(lngDo&).parent Is Nothing Then
      strBuddies$ = strBuddies$ & "g " & frmBuddyList.tvwSetup.Nodes.Item(lngDo&).Text & Chr(1)
    Else
      strBuddies$ = strBuddies$ & "b " & frmBuddyList.tvwSetup.Nodes.Item(lngDo&).Text & Chr(1)
    End If
  Next
  If InStr(strBuddies$, Chr(1)) Then
    strBuddies$ = Left(strBuddies$, Len(strBuddies$) - 1)
  End If
  Call WritePrivateProfileString(strScreenName$, "buddylist", strBuddies$, App.Path & "\aim.ini")
End Sub

Public Function BuddyConfig() As String
  'this procedure creates the buddies/permit/deny list to be used with the toc_set_config
  'message.
  Dim strBuddies As String, strStart As String, lngDo As Long
  Dim strPermit As String, blnPermitBuddies As Boolean
  Select Case m_strMode$
    Case "3"
      If m_strPDList$ <> "" Then
        strBuddies$ = strBuddies$ & Replace(m_strPDList$, " ", Chr(10) & "p ") & Chr(10)
      End If
    Case "4"
      If m_strPDList$ <> "" Then
        strBuddies$ = strBuddies$ & Replace(m_strPDList$, " ", Chr(10) & "d ") & Chr(10)
      End If
    Case "5"
      blnPermitBuddies = True
  End Select
  For lngDo& = 1 To frmBuddyList.tvwSetup.Nodes.Count
    If frmBuddyList.tvwSetup.Nodes.Item(lngDo&).parent Is Nothing Then
      strBuddies$ = strBuddies$ & "g " & LCase(Replace(frmBuddyList.tvwSetup.Nodes.Item(lngDo&).Text, " ", "")) & Chr(10)
    Else
      strBuddies$ = strBuddies$ & "b " & LCase(Replace(frmBuddyList.tvwSetup.Nodes.Item(lngDo&).Text, " ", "")) & Chr(10)
      If m_strMode$ = "5" Then
        strPermit$ = strPermit$ & "p " & LCase(Replace(frmBuddyList.tvwSetup.Nodes.Item(lngDo&).Text, " ", "")) & Chr(10)
      End If
    End If
  Next
  If m_strMode$ = "5" Then
    strStart$ = "{m 3" & Chr(10)
  Else
    strStart$ = "{m " & m_strMode$ & Chr(10)
  End If
  BuddyConfig$ = strStart$ & strPermit$ & strBuddies$ & "}"
End Function

Public Function GetBuddies() As String
  'retreives only the buddies from our treeview excluding the groups
  Dim strBuddies As String, lngDo As Long
  For lngDo& = 1 To frmBuddyList.tvwSetup.Nodes.Count
    If Not frmBuddyList.tvwSetup.Nodes.Item(lngDo&).parent Is Nothing Then
      strBuddies$ = strBuddies$ & " " & LCase(Replace(frmBuddyList.tvwSetup.Nodes.Item(lngDo&).Text, " ", ""))
    End If
  Next
  GetBuddies$ = Trim(strBuddies$)
End Function

Public Function FormByCaption(strMatch As String) As Long
  'since we are creating many forms dynamically, it is important for us to locate specific
  'forms. this procedure searches by the caption property while the one below searches
  'by the tag
  Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If LCase(Replace(Forms(lngDo&).Caption, " ", "")) = strMatch$ Then
      lngFound& = lngDo&
      Exit For
    End If
  Next
  FormByCaption& = lngFound&
End Function

Public Function FormByTag(strMatch As String) As Long
  Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If Forms(lngDo&).Tag = strMatch$ Then
      lngFound& = lngDo&
      Exit For
    End If
  Next
  FormByTag& = lngFound&
End Function

Public Function HideForm()
Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If MatchShit("*- Instant Message", Forms(lngDo&).Caption) = True Then
        'MsgBox "Found! - HideForm"
        Forms(lngDo&).Hide
    End If
  Next
  'FormByTag& = lngFound&
End Function

Public Function ShowForm()
Dim lngDo As Long, lngFound As Long
  lngFound& = -1
  For lngDo& = 0 To Forms.Count - 1
    If MatchShit("*- Instant Message", Forms(lngDo&).Caption) = True Then
    'MsgBox "Found! - ShowForm"
        Forms(lngDo&).Show
    End If
  Next
  'FormByTag& = lngFound&
End Function

Public Sub PlayWav(strWav As String)
  If FileExists(strWav$) Then Call sndPlaySound(strWav$, SND_ASYNC)
End Sub

Sub ListAdd(TheEvent)
Dim B1, B2, B3, B4, B5
Dim i As Integer
For i = 0 To frmAway.List1.ListCount - 1
    If Right(frmAway.List1.List(i), Len(TheEvent) + 2) = "x " & TheEvent Then
        B1 = frmAway.List1.List(i)
        B2 = InStr(1, B1, "x")
        B3 = Left(B1, B2 - 2)
        B4 = Right(B1, Len(B1) - (B2 + 1))
        B3 = B3 + 1
        frmAway.List1.List(i) = B3 & " x " & B4
        Exit Sub
    End If
    DoEvents
Next i
frmAway.List1.AddItem "1 x " & TheEvent
End Sub

Public Function MatchShit(a As String, b As String) As Boolean
Dim X
Dim ToFind As String
Dim start As Long
    'wild card text matcher so far it only s
    '
    ' upports one * but I'll
    'be trying to get it to work for unlimit
    '
    ' ed numbers of *, i havent found any
    'bugs. I saw a few subs that did somethi
    '
    ' ng like this but they were not as
    'good this is the first of it's kind, as
    '
    ' far as i know..
    '
    'if you can get it to work with unlimite
    '
    ' d number of *'s please email me it!
    'use the freely sub as you please long a
    '
    ' s you give me credit where it's due,
    'thanks! also dont forget to vote :)
    '
    'Using the sub: a is the string that is
    '
    ' the string that has an "*" in it
    'b is the base string meaing a is trying
    '
    ' to be matched to b.
    '
    'ie: a = 127.0.*, b = 127.0.0.1
    '
    'usage: MsgBox IPmatch(Text1, Text2)
    'usage2: if IPmatch(Text1, Text2) = fals
    '
    ' e then : msgbox "error no match!",16,"
    '     er
    ' ror"
    '
    '
    'Contacts:
    'Email: eko2k@hotmail.com
    'Aim:duckee2k
    '
    'oh and, btw open source 4lyfe
    On Error GoTo err
    'check the easy way :)
    If a = b Then: MatchShit = True: Exit Function
    'strings arent identical so here we go :
    '
    ' \
    X = Split(a, "*") 'searches For the *'s
    'defines the vars
    ToFind$ = Mid(a, InStr(a, X(1)))
    start& = InStr(a, X(1))
    'checks the legnth, 0,1,2+ need differnt
    '
    ' subs
    If Len(X(1)) = 1 Then GoTo short
    If Len(X(1)) = 0 Then GoTo none
    'core coding:
    'has more then one chr(s) after the *


    If Left(a, InStr(a, X(1)) - 2) = Left(b, InStr(a, X(1)) - 2) Then
        On Error GoTo err


        If Mid(a, InStr(a, X(1))) = Mid(b, InStr(start&, b, ToFind$)) Then
            MatchShit = True
        Else: GoTo err
        End If
    End If
    'has only one chr(s) after the *
short:


    If Left(a, InStr(a, X(1))) = Left(b, InStr(a, X(1))) Then
        On Error GoTo err


        If Right(a, InStr(a, X(1))) = Left(b, InStr(a, X(1))) Then
            MatchShit = True
        End If
    End If
    'has nothing after the *
none:


    If Len(a) = Len(b) Then
        On Error GoTo err


        If Left(a, Len(a) - 1) = Left(b, Len(a) - 1) Then
            MatchShit = True
        End If
    Else


        If Left(a, Len(a) - 1) = Left(b, Len(a) - 1) Then
            MatchShit = True
        End If
    End If
    Exit Function
err:
    MatchShit = False: Exit Function
End Function

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer


    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
