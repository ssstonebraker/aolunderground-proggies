Attribute VB_Name = "Module1"
'sup
'alot of this stuff in this bas is for
'a mp3 player. To make this work u need Dos32.ocx
'dos32 bas.  and media player ocx. (msxdm.ocx)
'if you do not have this ocx then go to
'www.microsoft.com and dl windows media player.
'Then u can browse for the ocx and load it in
'to ur project any questions email me. Oh and
'by the way.  Its a Cut and paste bas.
'so if you use it i should be put in the greetz
'cause most of this is my coding.  Knk and Arse
'wrote some of the codes so i will give them
'some credit for the bas.
'itz_mage@yahoo.com  or catch me on aim at
'IAmTheGreatMage
' peace
' -mage


Sub Random()

'this is used to make a random function
i% = Int(Rnd * 2) + 1
If i% = 1 Then mssg$ = "blah"

If i% = 2 Then mssg$ = "blah"

Text1.Text = mssg$

End Sub

Sub LCase()

'this is used for a c-chat
If Screen_Name Like GetUser And What_Said Like "trigger" Then
'code here

End Sub


Sub Idle_Min_Sec()

'this can only be used with dos bas i think....
'try without it

Dim thetexT As String
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
aChild& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Something& = FindWindowEx(aChild&, 0, "RICHCNTL", vbNullString)
ourhandle& = FindWindowEx(aChild&, Something&, "RICHCNTL", vbNullString)
thetexT$ = GetText(ourhandle&)
Call SetText(ourhandle&, "")
Text5 = Text5 + 1
idle4 = Text6 & ":" & Text5
ChatSend ("Idle " & idle4 & " -")


Call SetText(ourhandle&, thetexT)
If Text5 = 59 Then
thetexT$ = GetText(ourhandle&)
Call SetText(ourhandle&, "")
Text6 = Text6 + 1
Text5 = 0
idle4 = Text6 & ":" & Text5
ChatSend ("Idle- " & idle4 & " ")
Call SetText(ourhanlde&, thetexT)
End If

End Sub


Sub Mp3GetPath()

'to make this work you need to set the directroy. SetDirofMp3
'then put the following code in form_load.

On Error Resume Next
File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.ini")

End Sub

Sub Mp3SetDir()

'strArgument1$ this is for a c-chat mp3 player
strArgument1$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument1$
      ChatSend " dir set [" & strArgument1$ & "]"
      Call WriteToINI("mp3dir", "dir", strArgument1$, App.Path & "\mp3.ini")

End Sub

Sub Mp3Play()

On Error Resume Next
       File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.ini")
      Dim possibles%, Others$
             strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace)
             thepart$ = strArgument1$
      thepart$ = LCase(ReplaceString(thepart$, " ", ""))
    File1.Path = GetFromINI("mp3dir", "dir", App.Path & "\mp3.ini")
            
    thepart = Mid(What_Said, 7)
    thepart = LCase(ReplaceString(thepart, " ", ""))
For z = 0 To File1.ListCount - 1
Chcky1$ = File1.List(z)
d = InStr(1, Chcky1$, thepart, vbTextCompare)
If d > 0 Then
If Others = "" Then Others = File1.List(z): possibles = (possibles + 1) Else possibles = (possibles + 1)
End If
Next z
If possibles > 1 Then ChatSend "[" & possibles% & "] Possibilitys": Exit Sub
If possibles = 0 Then ChatSend "[" & thepart & "] not found": Exit Sub
fiiles$ = "" & File1.Path & "\" & Others & ""
MediaPlayer1.Open fiiles$
ChatSend " now playing [" & ReplaceString(Others, ".mp3", "") & "]"

End Sub
Sub Mp3RandomPlay()
'to make this work u need to take the
'randomnumber function and put it in the
'bas you are using
On Error Resume Next
            File1.ListIndex = RandomNumber1(File1.ListCount)
            strArgument1$ = File1.Path & "\" & File1.filename
            On Error Resume Next
            MediaPlayer1.Open strArgument1$
            MediaPlayer1.Play
            ChatSend " - Playing [" & ReplaceString(ReplaceString(strArgument1$, File1.Path, ""), "\", "") & "]"

End Sub


 Sub Mp3Pause()
    
    MediaPlayer1.Pause

End Sub


Sub Mp3Stop()

MediaPlayer1.Stop

End Sub

Sub AddCchatNote()

' I just Figured i would put this in here
' for the hell of it.  i think u can figure
' the rest of the options out ur self
 If Len(What_Said$) > 1& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
             List1.AddItem "" + strArgument1$ + ""
               End If

End Sub
 Sub ScrollListBox()
 
 For a = 0 To List1.ListCount - 1
ChatSend "" & List1.List(a)
Pause 0.7
Next a

End Sub

 Sub Cchat_SelectCase()
'this is if u want to make a c-chat
'you mite wanna use this code
'after it use  case " blah"

Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If Screen_Name$ = GetUser$ And InStr(What_Said, ".") = 1& Then
      lngSpace& = InStr(What_Said$, " ")
      If lngSpace& = 0& Then
         strCommand$ = What_Said$
      Else
         strCommand$ = Left(What_Said$, lngSpace& - 1&)
      End If
      Select Case strCommand$

End Sub

Sub keyword_locate()

'aol://3548:" $member"

End Sub

Sub Keyword_PrivateRoom()

'aol://2719:2-2-

End Sub
 Sub Keyword_MemberRoom()

'aol://2719:61-2-

End Sub

 Sub Keyword_PublicRoom()

'aol://2719:21-2-

End Sub

 Sub Im_KeyWord()

'aol://9293:

End Sub

 Sub strArgument1()

'this code is for a c-chat to get the text
'after a command
If Len(What_Said$) > 1& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)

End Sub

 Sub ListToText()
 
 'this code is for making a list a text
Dim X As Long, people As String
For X& = 0 To List1.ListCount - 1
        people$ = people$ & List1.List(X) & ","
        Next X&
    Text1 = people$

End Sub

 Sub SaveText()

CommonDialog1.ShowSave
If CommonDialog1.filename <> "" Then
 Open CommonDialog1.filename For Output As #1
 Print #1, Text1.Text
 Close #1
End If

End Sub

 Sub LoadText()

CommonDialog1.ShowOpen
If CommonDialog1.filename <> "" Then
 Open CommonDialog1.filename For Input As #1
 Do While Not EOF(1)
 Text1.Text = Text1.Text & Input(1, #1)
 Loop
 Text1.Text = Trim(Left(Text1.Text, Len(Text1.Text) - 2))
 Close #1
End If

End Sub

 Sub ScrollMultiLinedText()

text2 = Text1
If Mid(Text1, Len(Text1), 1) <> Chr$(10) Then
    Text1 = Text1 + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text1, Chr$(13)) <> 0)
    Counter = Counter + 1
    ChatSend Mid(Text1, 1, InStr(Text1, Chr(13)) - 1)
    Pause 0.4
    If Counter = 4 Then
        Pause (1.1)
        Counter = 0
    End If
    Text1 = Mid(Text1, InStr(Text1, Chr(13) + Chr(10)) + 2)
Text1 = text2
Loop

End Sub

 Sub MchatCode()

'Put this code in the form with the ocx  i think u can figure
'out the rest if not email me

Text1.SelStart = Len(Text1.Text)
Text1.SelText = vbCrLf & Screen_Name & ":" & Chr(9) & What_Said

End Sub

 Sub PopUpMenu()

'this is from a label
Form2.PopUpMenu Form2.Label, 1

End Sub

 Sub PopUpMenu2()
'pop up from a button

Form2.PopUpMenu Form2.Button, 1

End Sub



Sub Mp3UnPause()

MediaPlayer1.Play

End Sub

 Sub Mp3Volume()

If Len(What_Said$) > 4& Then
               strArgument1$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
                Pause 0.5
            
            MediaPlayer1.Volume = (strArgument1$ * 25) - 2500

End Sub

 Sub Mp3TrackLength()
 
 Dim Minutes As String, Seconds As String
    Minutes = 0
           Seconds = MediaPlayer1.Duration
           Do While Val(Seconds) >= 60
              Seconds = Val(Seconds) - 60
               Minutes = Val(Minutes) + 1
            Loop
            Seconds = Left(Seconds, InStr(Seconds, ".") - 1)
            If Len(Seconds) = 1 Then Seconds = "0" + Seconds
 'to show [" & minutes & ":" & seconds & "]

End Sub

 Sub Mp3_MuteOn()
MediaPlayer1.mute = True
End Sub
 Sub Mp3_MuteOff()
MediaPlayer1.mute = False
End Sub

 Function RandomNumber1(finished)
Randomize
RandomNumber1 = Int((Val(finished) * Rnd) + 1)
End Function
End Function

Sub Mp3PlayOCX_Register()
'Mp3Play1.Authorize "PLAY3326782111", "4796646"
End Sub

Sub Mp3Play_SongScroller()

'this is so when like u press play and there are
'more then 1 strings, it will sroll songs
'you need to put the following codes into your project
'so it will work.

'you need to put the following code in the top so it will work
'this would go where the declartaions are in the project
'not in the bas
Dim Multi As Boolean

'here is the play code

Case ".play"
If Multi <> True Then
    Call PlayIt(Right(What_Said, Len(What_Said) - 6))
    Exit Sub
End If
If Multi = True Then
    On Error GoTo kdmerr
    intt% = Right(What_Said, Len(What_Said) - 6)
    MediaPlayer1.Open File1.Path & "\" & List2.List(intt%)
    ChatSend "  now playing [" & ReplaceString(List2.List(intt%), ".mp3", "") & "]"
    Multi = False
    List2.Clear
    Exit Sub
kdmerr:
    If What_Said <> " [type ''.play'' and 0 - " & List2.ListCount - 1 & "or ''.cancel'']" Then ChatSend "<font color=#ffffff><font face=""arial"">×­› audio¹  [invaled number]"
End If

'this code needs to be put in ur project also to make
'the play function work
Public Sub PlayIt(sString As String)
    For i = 0 To File1.ListCount - 1
        If InStr(LCase(TrimSpaces(File1.List(i))), LCase(TrimSpaces(sString$))) Then
            List2.AddItem (File1.List(i))
        End If
    Next i
    If List2.ListCount = 1 Then
        MediaPlayer1.Open File1.Path & "\" & List2.List(0)
        ChatSend " playing -[" & ReplaceString(List2.List(0), ".mp3", "") & "]"
        List2.Clear
        Exit Sub
    End If
    If List2.ListCount = 0 Then
        ChatSend " [no strings found]"
    End If
    If List2.ListCount > 1 Then
        ChatSend "  [multiple strings found]"
        For i = 0 To List2.ListCount - 1
            ChatSend " [" & i & ") " & List2.List(i) & "]"
            Pause 1
        Next i
        ChatSend "  [type ''.play'' and 0 - " & List2.ListCount - 1 & " or ''.cancel'']"
        Multi = True
    End If
End Sub



