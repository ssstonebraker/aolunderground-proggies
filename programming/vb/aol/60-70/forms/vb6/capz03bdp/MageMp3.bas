Attribute VB_Name = "Module1"
Sub Random()

'this is used to make a random function
i% = Int(Rnd * 2) + 1
If i% = 1 Then mssg$ = "blah"

If i% = 2 Then mssg$ = "blah"

Text1.text = mssg$

End Sub

Sub LCas()

'this is used for a c-chat
If Screen_Name Like GetUser And What_Said Like "trigger" Then
'code here
End If
End Sub

Sub Mp3SetDir()

'strArgument1$ this is for a c-chat mp3 player
strArgumen$ = Mid(What_Said, 6)
      On Error Resume Next
      File1.Path = strArgument$
      ChatSend " dir set [" & strArgument$ & "]"
      Call WriteToINI("Settings", "dir", strArgument$, "c:\windows\system\x982.set")

End Sub

Sub Mp3Play()

On Error Resume Next
       File1.Path = GetFromINI("Settings", "dir", "c:\windows\system\x982.set")
Dim possibles%, Others$
             strArgument$ = Right(What_Said$, Len(What_Said$) - lngSpace)
             thepart$ = strArgument$
      thepart$ = LCase(ReplaceString(thepart$, " ", ""))
    File1.Path = GetFromINI("Settings", "dir", "c:\windows\system\x982.set")

            
    thepart = Mid(What_Said, 7)
    thepart = LCase(ReplaceString(thepart, " ", ""))
For Z = 0 To File1.ListCount - 1
Chcky1$ = File1.List(Z)
d = InStr(1, Chcky1$, thepart, vbTextCompare)
If d > 0 Then
If Others = "" Then Others = File1.List(Z): possibles = (possibles + 1) Else possibles = (possibles + 1)
End If
Next Z
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
            strArgument$ = File1.Path & "\" & File1.FileName
            On Error Resume Next
            MediaPlayer1.Open strArgument$
            MediaPlayer1.Play
            ChatSend "[" & ReplaceString(ReplaceString(strArgument$, File1.Path, ""), "\", "") & "]"

End Sub


 Sub Mp3Pause()
    
    MediaPlayer1.Pause

End Sub


Sub Mp3Stop()

MediaPlayer1.Stop

End Sub

Sub AddCchatNote()
End Sub
 Sub ScrollListBox()
 
 For a = 0 To List1.ListCount - 1
ChatSend "" & List1.List(a)
Pause 0.7
Next a

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


 Sub ScrollMultiLinedText()

Text2 = Text1
If Mid(Text1, Len(Text1), 1) <> Chr$(10) Then
    Text1 = Text1 + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text1, Chr$(13)) <> 0)
    counter = counter + 1
    ChatSend Mid(Text1, 1, InStr(Text1, Chr(13)) - 1)
    Pause 0.4
    If counter = 4 Then
        Pause (1.1)
        counter = 0
    End If
    Text1 = Mid(Text1, InStr(Text1, Chr(13) + Chr(10)) + 2)
Text1 = Text2
Loop

End Sub

 Sub MchatCode()

'Put this code in the form with the ocx  i think u can figure
'out the rest if not email me

Text1.SelStart = Len(Text1.text)
Text1.SelText = vbCrLf & Screen_Name & ":" & Chr(9) & What_Said

End Sub

Sub Mp3UnPause()

MediaPlayer1.Play

End Sub

 Sub Mp3Volume()

If Len(What_Said$) > 4& Then
               strArgument$ = Right(What_Said$, Len(What_Said$) - lngSpace&)
                Pause 0.5
            
            MediaPlayer1.Volume = (strArgument1$ * 25) - 2500
End If
End Sub

 Sub Mp3TrackLength()
 
 Dim minutes As String, Seconds As String
    minutes = 0
           Seconds = MediaPlayer1.Duration
           Do While Val(Seconds) >= 60
              Seconds = Val(Seconds) - 60
               minutes = Val(minutes) + 1
            Loop
            Seconds = Left(Seconds, InStr(Seconds, ".") - 1)
            If Len(Seconds) = 1 Then Seconds = "0" + Seconds
 'to show [" & minutes & ":" & seconds & "]

End Sub

 Sub Mp3_MuteOn()
MediaPlayer1.Mute = True
End Sub
 Sub Mp3_MuteOff()
MediaPlayer1.Mute = False
End Sub

 Function RandomNumber1(finished)
Randomize
RandomNumber1 = Int((Val(finished) * Rnd) + 1)
End Function

Sub Mp3PlayOCX_Register()
'Mp3Play1.Authorize "PLAY3326782111", "4796646"
End Sub
'Public Sub PlayIt(sString As String)
  '  For i = 0 To File1.ListCount - 1
      '  If InStr(LCase(TrimSpaces(File1.List(i))), LCase(TrimSpaces(sString$))) Then
          '  List2.AddItem (File1.List(i))
       ' End If
   ' Next i
    'If List2.ListCount = 1 Then
      '  MediaPlayer1.Open File1.Path & "\" & List2.List(0)
     '   ChatSend " playing -[" & ReplaceString(List2.List(0), ".mp3", "") & "]"
     '   List2.Clear
      '  Exit Sub
   ' End If
    'If List2.ListCount = 0 Then
    '    ChatSend " [no strings found]"
   ' End If
   ' If List2.ListCount > 1 Then
   '     ChatSend "  [multiple strings found]"
    '    For i = 0 To List2.ListCount - 1
    '        ChatSend " [" & i & ") " & List2.List(i) & "]"
    '        Pause 1
    '    Next i
    '    ChatSend "  [type ''.play'' and 0 - " & List2.ListCount - 1 & " or ''.cancel'']"
   '     Multi = True
   ' End If
'end sub


