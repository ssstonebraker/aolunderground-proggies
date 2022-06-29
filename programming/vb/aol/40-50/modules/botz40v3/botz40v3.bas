Attribute VB_Name = "botz40v³"
'Hey, thanks for downloading my bas for bots.  This bas
'would be used for AOL 4.0 or 5.0.  Well this is by FAR the last
'version of botz40.  I don't program for AOL at all but since
'people still bug me about my old versions, and my old versions are
'terrible compared to my abilities now, I feel I must clear my name
'and my programming skills.  I made version 2.0 about a year ago,
'and it does not reflect my skills at all today.  I didnt add any new bots
'but simply rewrote them all better and each one has detailed
'instructions. So here is my last and final version of botz40.
'Enjoy

'sabre {vakattak@yahoo.com}
'aim: vak attak
Option Explicit

Sub STFUBot(NameOfPerson As String)
'You will need to make 1 textbox and 2 command buttons.
'Make one command button's caption be Start and the other's Stop.

'In the Start command button put:

'Call STFUBot(Text1.Text)

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "STFUBot Unloaded"

'The textbox is the name of the person who you want to STFU.
    Do
        If SNFromLastChatLine = NameOfPerson$ Then
          sendchat "STFU " & NameOfPerson$
          TimeOut 7
        End If
    Loop
End Sub


Sub IdleBot(Reason As String, HowLong As Integer)
'Make two textboxes and two command buttons.
'Make one of the command button's caption be Start and the other's Stop
'In the Start command button put:

'Call IdleBot(Text1, Text2)

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "Idlebot Unloaded"

'One of the textboxes is the reason for being idle and the other textbox
'is for how long.

sendchat "Idle Bot Loaded"
TimeOut 1.2
    sendchat "Reason for being idle: " & Reason$
        TimeOut 1
    sendchat "How Long I'll be idle: " & HowLong
        TimeOut 0.7
    sendchat "I'll be back..."
        Do
          If HowLong = 0 Then Exit Do
          HowLong = HowLong - 1
          TimeOut 1
        Loop
    sendchat "I'm back, everybody"
        TimeOut 1.4
    sendchat "Idle Bot Unloaded"
End Sub

Sub RequestBot(WhatYouRequest As String)
'Make one textbox and two command buttons.
'The textbox is what you are requesting.
'Make one command button's caption be Start and the other's be Stop.

sendchat "Request Bot Loaded"
    Do
        sendchat "I request a " & WhatYouRequest$
        TimeOut 4
        SemdChat "Please answer my request!"
    Loop
'In the Start command button put:

'Call RequestBot(Text1.Text)

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "Request Bot Unloaded"
End Sub

Sub AttentionBot()
'Make one command button and make it's caption be AttentionBot.

'In the command button's click event put:

'Call AttentionBot

    sendchat "{S IM"
        TimeOut 3
    sendchat "Attention Bot Loaded"
    sendchat "Gimme Attention!"
        TimeOut 2
    sendchat "{S IM"
        TimeOut 2
    sendchat "{S BUDDYIN"
        TimeOut 1
    sendchat "{S GOTMAIL"
        TimeOut 2
    sendchat "                              "
End Sub

Sub AdvertizeBot(NameOfProgram As String, Handle As String, AttentionGetter As String)
'Make a command button and make it's caption be AdvertizeBot.

'In the command button's click event put:

'Call AdvertizeBot("my prog", "sabre", "this is dah best prog in dah world! =]")
    
    sendchat NameOfProgram$
    TimeOut 1
    sendchat "made by " & Handle$
    TimeOut 1.8
    sendchat AttentionGetter$
End Sub

Sub EchoBot(PersonToEcho As String)
'Make two command buttons and one textbox.
'Make one command button's caption be Start and the other's Stop
'In the Start command button put:

'Call EchoBot(Text1.Text)

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "EchoBot Unloaded"

'The textbox is the person you want to echo.

Dim line As String, SNLine As String

sendchat "Echo Bot Loaded"
  TimeOut 1
sendchat "Now echoing " & PersonToEcho$
  TimeOut 2
    Do
      SNLine$ = SNFromLastChatLine
        If SNLine$ = PersonToEcho$ Then
            line$ = lastchatline
            sendchat line$
        End If
      TimeOut 5
    Loop
End Sub

Sub FightBot(Fighter1 As String, Fighter2 As String)
'Make two textboxes on your form and one command button.
'Make the command button's caption be Fight.

'Insert this in the command button's click event:
'Call Fightbot(Text1, Text2)
Dim Number As Integer
Randomize
Number = Int(Rnd * 11)
If Number = 1 Or Number = 3 Or Number = 5 Or Number = 7 Or Number = 9 Then
    sendchat Fighter1 & " punches " & Fighter2
    TimeOut 1
    sendchat Fighter1 & " kicks " & Fighter2
    TimeOut 1
    sendchat Fighter1 & " kills " & Fighter2
    TimeOut 1
    sendchat "The Victor is: " & Fighter1
  Else
    sendchat Fighter2 & " punches " & Fighter1
    TimeOut 1
    sendchat Fighter2 & " kicks " & Fighter1
    TimeOut 1
    sendchat Fighter2 & " kills " & Fighter1
    TimeOut 1
    sendchat "The Victor is: " & Fighter2
End If
End Sub


Public Sub TimeOut(duration)
Dim Starttime As Long
  
Starttime = Timer
    Do While Timer - Starttime > duration
      DoEvents
    Loop
End Sub

Sub FakeBot(LeftTag As String, ProgName As String, ProgAuthor As String, RightTag As String)
'Make 4 textboxes and 1 command button.
'Two textboxes are what goes on the outsides of the Prog Name and the Prog
'Author.  The other two are the Prog's Author and the Prog Name.

'example:
'Text1 is the left tag, Text2 is the Prog Name, Text3 is the Prog Author,
'and Text4 is the right tag.

'In the command button's click event put:

'Call FakeBot(Text1, Text2, Text3, Text4)

    sendchat LeftTag$ & ProgName$ & RightTag$
    sendchat LeftTag$ & ProgAuthor$ & RightTag$
End Sub

Sub HiBot()
'Make two command buttons.
'Make one of the command button's caption be Start and the other's Stop
'In the Start command button put:

'Call HiBot

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "HiBot Unloaded"

sendchat ("Hi Bot Loaded.  Hi Everybody!  Say Hi to me!")
    Do
      If lastchatline = "Hi" Or lastchatline = "hi" Then
        sendchat "Hi " & SNFromLastChatLine & "!"
        TimeOut 4
      End If
    Loop
End Sub

Sub CustomBot(WhatTheySay As String, WhatYouSay As String)
'This unique bot is like an echo bot, but only responds to a certain word.
'Make two textboxes and two command buttons.
'Make one of the command button's caption be Start and the other's Stop
'In the Start command button put:

'Call CustomBot(Text1, Text2)

'In the Stop command button put:

'Do
'DoEvents
'Loops
'SendChat "CustomBot Unloaded"

'One textbox is what they say and the other textbox is what you say
'whenever someone in the chat says the word in the other textbox.

sendchat "CustomBot Loaded"
    Do
      If lastchatline = WhatTheySay Then sendchat WhatYouSay
      TimeOut 4
    Loop
End Sub

Sub QuizBot(Question As String, Answer As String)
'Make two textboxes and two command buttons.
'One of the textboxes is your question, and the other is the answer.
'Make one of the command button's caption be Start and the other's Stop
'In the Start command button put:

'Call QuizBot(Text1.Text, Text2.Text)

'In the Stop command button put:

'Do
'DoEvents
'Loops

sendchat "QuizBot Loaded"
TimeOut 2
    sendchat "QuizBot question: " & Question$
    Do
        If lastchatline = Answer$ Then
          sendchat SNFromLastChatLine & ", you're right!"
          TimeOut 1
          sendchat "QuizBot Unloaded"
          Exit Sub
        End If
    Loop
End Sub

Sub EightBallBot()
'Make two command buttons.
'Make one of the command button's caption be Start and the other's Stop.
'In the Start command button put:

'Call EightBallBot

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "EightBall Bot Unloaded"

Dim Number As Integer, Result As String, chat As String
Dim Times As Integer
    Randomize
    Number = Int(Rnd * 6)
    Times = 10
        sendchat "EightBall Bot Loaded"
          TimeOut 2
        sendchat "Type /8ball and your question"
Select Case Number
    Case 0
        Result$ = "it is decidedly so"
    Case 1
        Result$ = "keep dreaming.  Not likely"
    Case 2
        Result$ = "no way"
    Case 3
        Result$ = "this is your lucky day!  Yes!"
    Case 4
        Result$ = "try again"
    Case 5
        Result$ = "great chance its true!"
End Select
  Do
    chat$ = lastchatline
      If InStr("/8ball ", chat$) Then
        sendchat SNFromLastChatLine & ", " & Result$
        Times = Times - 1
      End If
      TimeOut 2
      If Times = 0 Then
        sendchat "EightBall Bot Unloaded"
        Exit Sub
      End If
  Loop
End Sub

Sub ScrambleBot()
'To use this, just put Call ScrambleBot in a command button.
Dim aString As String, eString As String, iString As String
Dim Number As Integer
  sendchat "ScrambleBot Loaded"
    TimeOut 0.8
  sendchat "Try to unscramble the words"
Randomize
Number = Int(Rnd * 6)
Select Case Number
    Case 0
        aString$ = "money"
        eString$ = Left(aString$, 2)
        iString$ = right(aString$, 3)
    Case 1
        aString$ = "eggs"
        eString$ = Left(aString$, 2)
        iString$ = right(aString$, 2)
    Case 2
        aString$ = "poop"
        eString$ = Left(aString$, 3)
        iString$ = right(aString$, 1)
    Case 3
        aString$ = "bacon"
        eString$ = Left(aString$, 2)
        iString$ = right(aString$, 3)
    Case 4
        aString$ = "sting"
        eString$ = Left(aString$, 3)
        iString$ = right(aString$, 2)
    Case 5
        aString$ = "sabre"
        eString$ = Left(aString$, 2)
        iString$ = right(aString$, 3)
End Select
  sendchat "Try to unscramble: " & iString$ & eString$
    Do
        If lastchatline = aString$ Then sendchat "Correct, the word is " & aString$
    Loop Until lastchatline = aString$
  sendchat "ScambleBot Unloaded"
End Sub

Sub LuckyNumberBot()
'Make two command buttons.
'Set the one command button's caption be Start and the other's Stop.
'In the Start command button put:

'Call LuckyNumberBot

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "Lucky Number Bot Unloaded"

Dim Number As Integer
    sendchat "Lucky Number Bot Loaded"
    TimeOut 2
    sendchat "Type /luckynumber to see your lucky Number"
    TimeOut 2
    Do
        Randomize
        Number = Int(Rnd * 1001)
        If lastchatline = "/luckynumber" Then sendchat (SNFromLastChatLine & ", your lucky number is " & Number)
    Loop
End Sub

Sub GuessNumberBot(TheNumberToGuess As Integer)
'Make two command buttons and a textbox.
'Set the one command button's caption be Start and the other's Stop.
'In the Start command button put:

'Call GuessNumberBot(Text1.Text)

'In the Stop command button put:

'Do
'DoEvents
'Loop
'SendChat "Guess Number Bot Unloaded"

'For the textbox, change the MaxLength property on the property
'window to 2.  This is the box where the user enters the number
'that they want the chat to guess.

sendchat "Guess Number Bot Loaded"
    If TheNumberToGuess < 10 Or TheNumberToGuess = 10 Then
        sendchat "I'm thinking of a number between 0-10.  What is it?"
      Do
        If lastchatline = TheNumberToGuess Then sendchat SNFromLastChatLine & " is right!"
      Loop Until lastchatline = TheNumberToGuess
    End If
    If TheNumberToGuess = 30 Or TheNumberToGuess < 30 And TheNumberToGuess > 10 Then
        sendchat "I'm thinking of a number between 11-30.  What is it?"
      Do
        If lastchatline = TheNumberToGuess Then sendchat SNFromLastChatLine & " is right!"
      Loop Until lastchatline = TheNumberToGuess
    End If
    If TheNumberToGuess = 50 Or TheNumberToGuess < 50 And TheNumberToGuess > 30 Then
        sendchat "I'm thinking of a number between 31-50.  What is it?"
      Do
        If lastchatline = TheNumberToGuess Then sendchat SNFromLastChatLine & " is right!"
      Loop Until lastchatline = TheNumberToGuess
    End If
    If TheNumberToGuess = 70 Or TheNumberToGuess < 70 And TheNumberToGuess > 50 Then
        sendchat "I'm thinking of a number between 51-70.  What is it?"
      Do
        If lastchatline = TheNumberToGuess Then sendchat SNFromLastChatLine & " is right!"
      Loop Until lastchatline = TheNumberToGuess
    End If
    If TheNumberToGuess < 100 And TheNumberToGuess > 70 Then
        sendchat "I'm thinking of a number between 71-100.  What is it?"
      Do
        If lastchatline = TheNumberToGuess Then sendchat SNFromLastChatLine & " is right!"
      Loop Until lastchatline = TheNumberToGuess
    End If
sendchat "Guess Number Bot Unloaded"

End Sub

'This module was made by:                                                                                                                                                                                                                                                                                                                                                                                                                                                                         'All code written and made by sabre, DO NOT COPY! ! !

'                                                  ______
'      _________________________________________  /
'     /-----------------------------------------\/|||||\)\
'    /___________________________________sabre__/\|||||/)/


