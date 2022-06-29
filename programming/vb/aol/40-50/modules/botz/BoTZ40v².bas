Attribute VB_Name = "BoTZ40v²"
'Hey, thanx for using(or checking) my bas for botz that
'would be used for AOL 4.o.  About half a week went into
'making this so maybe a couple subs won't work.  But i fixed
'a lot of mistakes that were on my previous version.  I made 1
'prog for AOL4.o called Sting Anti Punter. Itz for my own
'personal use so I won't send it out to people.
'If something doesn't work, please don't immediately e-mail me
'about it, mess around with it and u should be able to fix it.
'If u don't, please e-mail me (Vakattak@aol.com) and I will
'help u.  Please use this version over the other because it
'has more botz and most are fixed and work properly.  Thanx
'again for d/ling my bas and i hope u use these botz on
'yor progz.


'Cya

'SaBrE {Vakattak@aol.com}
Option Explicit

Sub MadBot()
'need:
'2 textboxes, 2 command buttons, and 1 timer
'in Command1 put:
'Timer1.Enabled = True
'in Command2 put:
'Do
'DoEvents
'Loop
'Timer1.Enabled = False

'Text1 is the name of the person
'Text2 is something that would tick some1 off

SendChat ("Man, " & Text1.Text & "you " & Text2.Text)
TimeOut (0.2)
SendChat ("You heard me, I said you " & Text2.Text)
Do
SendChat (Text1.Text & "suckz")
TimeOut (5)
Loop Until Timer1.Enabled = False
End Sub

Public Sub SendChat(Text As String)
'here is where u put yor code for sending text to a chat room
'there are many bas's out there and if u ask properly, u kan
'use theirs.  There are other things u kan use besides SendChat
'i just used SendChat cuz i just thought of it.  It kan be
'ChatSend,Chat, AOL4Chat, etc...

'this sub doesn't work but the bots do if u use a sub that
'can send text to a chat room and put it here
End Sub

Sub AFKBot()
'need:
'2 textboxes, 2 command buttons, and a 1 timer
'Put Timer's interval to 5 and in it put:
SendChat ("AFK Bot Activated")
SendChat ("AFK for:" & Text1.Text & " min(z)") 'text1 is how long u r afk
TimeOut (0.1)
SendChat ("Reason: " & Text2.Text) 'text2 is the reason
TimeOut (Text1.Text)
'in Command1 put:
'Timer1.Enabled = True
'in Command2 put:
'Timer1.Enabled = False
'Do
'DoEvents
'Loop
End Sub

Sub IdleBot()
'need:
'1 textboxes, 2 commandbuttons, and 2 labels
'make Label1's caption "Reason:" and put it above the textbox
'make Command1's caption "Start" and Command2's "Stop
'in Command1.Click put: Call IdleBot
'in Command2.Click put:
'Do
'DoEvents
'Loop
'make the text1's text = ""
'make Label2's caption "How Long:" and put it above text2
'make Text2.Text's = ""
'******START OF SUB CODE******'
SendChat ("Idle Bot Activated")
TimeOut (1)
SendChat ("Reason:" & Text1.Text)
TimeOut (0.1)
SendChat ("How Long:" & Text2.Text)
TimeOut (0.1)
SendChat ("I'll be back...")
TimeOut (Text2.Text)
End Sub

Sub RequestBot()
'need:
'2 Textboxes and 2 Command buttons
'Text1 is what you want
'Text2 is yor name

SendChat ("Request Bot Activated")
SendChat ("Request: " & Text1.Text)
TimeOut (0.1)
SendChat ("Better give " & Text2.Text & " what he wants")

'in Command1 put:
'Call RequestBot

'in Command2 put:
'SendChat ("Everything's cool now cuz" & Text2.Text & "got what he wanted")
'Sendchat ("and that is:" & Text1.Text)
'TimeOut (0.3)
'Sendchat ("Request Bot Deactivated")
End Sub

Sub AttentionBot()
'to use it put: Call AttentionBot
SendChat ("{S IM")
TimeOut (1.2)
SendChat ("Attention Bot")
SendChat ("Gimme Attention!")
TimeOut (0.1)
SendChat ("{S IM")
TimeOut (1.1)
SendChat ("{S BUDDYIN")
TimeOut (1.1)
SendChat ("{S GOTMAIL")
TimeOut (1.1)
End Sub

Sub AdvertizeBot()
'to use it put: Call AdvertizeBot
SendChat ("[write the name of your prog really fancy here]")
TimeOut (0.1) ' the timesouts are so u dont get logged off for scrolling
SendChat ("Made By [your handle here]")
TimeOut (0.1)
SendChat ("[Some cocky or attention-getter for yor prog]")
End Sub

Sub EchoBot()
'need:
'1 timer and 2 Command buttons
'With Timer1's interval at 5 and enabled = false put:
Dim LastChatLine As String
Dim SNLastChatLine As String
LastChatLine$ = LastChatLine 'this would be a sub that would get
                          'the last text line in a chat room
SNLastChatLine$ = SNFromLastChatLine 'this would also be a sub that would
                            'get the SN from the last text line
                            'in a chat room
SendChat ("Echo Bot Active")
SendChat ("Echoing: " & SNFromLastChatLine$)
TimeOut (0.3)
Do
SendChat (LastChatLine$)
TimeOut (4)
Loop Until Timer1.Enabled = False
'in Command1 put: Timer1.Enabled = True
'in Command2 put: Timer1.Enabled = False
End Sub

Sub FightBot()
'need:
'2 textboxes and 1 label
'work:
'make Label1 NOT visible, so they kant see it.
'the first textbox is 1 person the other textbox is the other
Label1.Caption = Int(Rnd * 3)
If Label1.Caption = 1 Then
SendChat (Text1.Text & "punches" & Text2.Text)
TimeOut (1)
SendChat (Text1.Text & "kicks" & Text2.Text)
TimeOut (1)
SendChat (Text1.Text & "kills" & Text2.Text)
TimeOut (1)
SendChat ("The Winner is: " & Text1.Text)
Else
SendChat (Text2.Text & "punches" & Text1.Text)
TimeOut (1)
SendChat (Text2.Text & "kicks" & Text1.Text)
TimeOut (1)
SendChat (Text2.Text & "kills" & Text1.Text)
TimeOut (1)
SendChat ("The Winner is: " & Text2.Text)
End If
End Sub


Public Sub TimeOut(Duration)
Dim Starttime As Long
  Starttime = Timer
Do While Timer - Starttime > Duration
DoEvents
Loop
End Sub

Sub FakeBot()
'need:
'_ Command buttons, it depends on how many progs u wanna do it to
'dont call this sub, just use the code given and put in each
'command button this:
'Call SendChat("[name of prog which i will provide a lot of em]")

'these are the opening things that some progs say when they
'are activated.

'********************START OF CODE***************************
'º¯`v´¯¯) PhrostByte By: Progee (¯¯`v´¯º

'.·´¯`·-  gøthíc nightmâres by másta  ­·´¯`·
'·._.--   aøl 4.o punt tools · loaded ---._.·

'· úpr mácro stùdio · másta ·

'-=·Sting Anti Punta 2.o Loaded·=-
'-=·MaDe By SaBrE·=-

'(¯\_ GøDZîLLa³·º _/¯)
'(¯\_ ßy ÇoLd _/¯)

'¢º°¤÷®ÍP§ 2øøø÷¤°º¢
' ¢º°¤÷£øÃdÊD÷¤°º¢

'-•(`(`·•Fate Zero v¹ Loaded•·´)´)•-

'•·.·´).·÷•[ Outlaw Mass Mailer by Twiztid

'(¯`·.····÷• ärméñïå¹ · kðkô
'(¯`·.····÷• îøâdèd

'•·._.·´¯`·>AoL 4.0 TooLz By: X GeNuS X
'•·._.·´¯`·>Status: LoaDeD
'•·._.·´¯`·>Ya'll BeTTa NoT MeSS WiT ThiS NiG!

'^····÷• James Bond Toolz Ver .007
'^····÷• By: Saßan

'(¯`•Prophecy²·° Loaded

'···÷••(¯`·._ CoRn Fader _.·´¯)••÷···
'···÷••(¯`·._Created by :::PooP:::_.·´¯)••÷···

'Blue Ice Punter¹ For AOL 4.0
'By STaNK

'¤-----==America Onfire Platinum
'¤-----==Loaded
'¤-----==Created (²›y Fatal Error

'.­”ˆ”­.•Fí/\/ä£ Få/\/†ä§y \/ïïï•.­”ˆ”­.
'·­„¸„­•·ßy RšZz•
'.­”ˆ”­.•£õàÐëÐ

'¤¤†³¹¹º†¤¤ SANNMEN †oºLz ¤¤†³¹¹º†¤¤
'¤¤†³¹¹º†¤¤    By:Má§†é®MinÐ    ¤¤†³¹¹º†¤¤
'¤¤†³¹¹º†¤¤ LOADED ¤¤†³¹¹º†¤¤

'··¤÷×(Rapier Bronze)×÷¤··
'··¤÷×(By Excalibur)×÷¤··
'··¤÷×(Works for 3.0 and 4.0!!!)×÷¤··

'<-==(`(` Icy Hot 2.0 For AOL 4.0 ')')==->
'<-==(`(` Loaded ')')==->

'(\›•‹ Im Backfire KiLLer ›•‹/)
'(\›•‹ By:phire Status:Loaded ›•‹/)
'(\›•‹ Im Backfire KiLLer ›•‹/)

'[_.·´¯° Indian Invasion Punter Loaded °¯`·._]

'***********************END OF CODE**************************
End Sub

Sub HiBot()
SendChat ("Hi Bot Loaded.  Hi Everybody!")
Do
If LastChatLine = "Hi" Then
SendChat ("Hi " & SNFromLastChatLine & "!")
Loop Until LastChatLine = "Hi"
End If
End Sub


Sub CustomBot()
'make sure u make a stop button and in it put:
'Do
'DoEvents
'Loop

'u need 2 textboxes
'Text1 is what they say and Text2 is what u want to say
Do
If Text1.Text = LastChatLine Then
SendChat (Text2.Text)
Loop Until Text1.Text = LastChatLine
End If
End Sub

Sub QuizBot()
'need:
'1 label
Label1.Caption = Int(Rnd * 11)
If Label1.Caption = "1" Then
SendChat ("What state is Harrisburg in?")
Do
If LastChatLine = "Pennsylvania" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "Pennsylvania"
End If
    If Label1.Caption = "2" Then
    SendChat ("How many inches r in a foot?")
    Do
    If LastChatLine = "12" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "12"
    End If
If Label1.Caption = "3" Then
SendChat ("How many hours r in a day?")
Do
If LastChatLine = "24" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "24"
End If
    If Label1.Caption = "4" Then
    SendChat ("How many days r in a week?")
    Do
    If LastChatLine = "7" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "7"
    End If
If Label1.Caption = "5" Then
SendChat ("What country has the most population?")
Do
If LastChatLine = "China" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "China"
    If Label1.Caption = "6" Then
    SendChat ("Does money suck?")
    Do
    If LastChatLine = "no" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "no"
    End If
If Label1.Caption = "7" Then
SendChat ("Which is bigger, USA or Japan?")
Do
If LastChatLine = "USA" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "USA"
End If
    If Label1.Caption = "8" Then
    SendChat ("How many points is a touchdown?")
    Do
    If LastChatLine = "6" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "6"
    End If
If Label1.Caption = "9" Then
SendChat ("Which is bigger, USA or Japan?")
Do
If LastChatLine = "USA" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "USA"
End If
    If Label1.Caption = "10" Then
    SendChat ("What kind of verb is ran?")
    Do
    If LastChatLine = "action" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "action"
    End If
End Sub

Sub ShhBot()
'need:
'_ Textboxes, it depends on how many people
Do
If SNFromLastChatLine = Text1.Text Then 'u kan add more textboxes for more people
SendChat ("STFU " & SNFromLastChatLine & "!")
Loop Until SNFromLastChatLine = Text1.Text
End If
End Sub

Sub EightBallBot()
'need:
'1 textbox, 1 label
'Text1 is what u r asking, and the label where u randomize
Label1.Caption = Int(Rnd * 9)
SendChat ("8Ball Bot Loaded")
TimeOut (0.4)
SendChat (Text1.Text & " = Question")
    If Label1.Caption = "1" Then
    SendChat ("Excellent Chance!")
    End If
If Label1.Caption = "2" Then
SendChat ("Great Chance!")
End If
    If Label1.Caption = "3" Then
    SendChat ("Good Chance!")
    End If
If Label1.Caption = "4" Then
SendChat ("OK Chance!")
End If
    If Label1.Caption = "5" Then
    SendChat ("Bad Chance!")
    End If
If Label1.Caption = "6" Then
SendChat ("Very Bad Chance!")
End If
    If Label1.Caption = "7" Then
    SendChat ("0% Chance!")
    End If
If Label1.Caption = "8" Then
SendChat ("Ur having a horrible day!")
End If
End Sub

Sub ScrambleBot()
'need:
'1 label
Dim aString As String, eString As String, iString As String
SendChat ("ScrambleBot Loaded")
TimeOut (0.2)
SendChat ("Try to unscramble the words:")
Label1.Caption = Int(Rnd * 6)

    If Label1.Caption = "1" Then
    aString$ = "eggs"
    eString$ = Left(aString$, 2)
    iString$ = Right(aString$, 2)
    TimeOut (0.3)
    SendChat (iString$ & eString$)
    Do
    If LastChatLine = "eggs" Then SendChat ("Correct, the word is eggs")
    Loop Until LastChatLine = "eggs"
    End If
 TimeOut (0.2)
If Label1.Caption = "2" Then
aString$ = "poop"
eString$ = Left(aString$, 3)
iString$ = Right(aString$, 1)
TimeOut (0.3)
SendChat (iString$ & eString$)
Do
If LastChatLine = "poop" Then SendChat ("Correct, the word is poop")
Loop Until LastChatLine = "poop"
End If
 TimeOut (0.2)
    If Label1.Caption = "3" Then
    aString$ = "bacon"
    eString$ = Left(aString$, 2)
    iString$ = Right(aString$, 3)
 TimeOut (0.3)
    SendChat (iString$ & eString$)
    Do
    If LastChatLine = "bacon" Then SendChat ("Correct, the word is bacon")
    Loop Until LastChatLine = "bacon"
    End If
 TimeOut (0.2)
If Label1.Caption = "4" Then
aString$ = "Sting"
eString$ = Left(aString$, 3)
iString$ = Right(aString$, 2)
TimeOut (0.3)
SendChat (iString$ & eString$)
Do
If LastChatLine = "Sting" Then SendChat ("Correct, the word is Sting")
Loop Until LastChatLine = "Sting"
End If
 TimeOut (0.2)
    If Label1.Caption = "5" Then
    aString$ = "SaBrE"
    eString$ = Left(aString$, 2)
    iString$ = Right(aString$, 3)
 TimeOut (0.3)
    SendChat (iString$ & eString$)
    Do
    If LastChatLine = "SaBrE" Then SendChat ("Correct, the word is SaBrE")
    Loop Until LastChatLine = "SaBrE"
End If
 TimeOut (0.2)
End Sub

Sub LuckyNumberBot()
'need:
'1 textbox and 1 label
Label1.Caption = Int(Rnd * 1000)
SendChat ("Lucky # Bot Loaded")
TimeOut (1)
SendChat ("type /luckynumber to see your lucky #")
TimeOut (1)
Do
If LastChatLine = "/luckynumber" Then SendChat (SNFromLastChatLine & " : " & Label1.Caption)
Loop Until LastChatLine = "/luckynumber"
End Sub

Sub GuessNumberBot()
'need:
'1 label
Label1.Caption = Int(Rnd * 6)
    If Label1.Caption = "1" Then
    SendChat ("I'm thinking of a number between 1-3")
    Do
    If LastChatLine = "2" Then SendChat ("2 is right!")
    Loop Until LastChatLine = "2"
    End If
If Label1.Caption = "2" Then
SendChat ("I'm thinking of a number between 10-13")
Do
If LastChatLine = "12" Then SendChat ("12 is right!")
Loop Until LastChatLine = "12"
End If
    If Label1.Caption = "3" Then
    SendChat ("I'm thinking of a number between 30-33")
    Do
    If LastChatLine = "31" Then SendChat ("31 is right!")
    Loop Until LastChatLine = "31"
    End If
If Label1.Caption = "4" Then
SendChat ("I'm thinking of a number between 40-43")
Do
If LastChatLine = "42" Then SendChat ("42 is right!")
Loop Until LastChatLine = "42"
End If
    If Label1.Caption = "5" Then
    SendChat ("I'm thinking of a number between 50-53")
    Do
    If LastChatLine = "51" Then SendChat ("51 is right!")
    Loop Until LastChatLine = "51"
    End If

End Sub

'This bas was made by:'                                                                                                                                                                                                                                                                                                                                                                                                                                                                         'All code written and made by SaBrE, Copyright 1999 : SaBrE.  If this is found on yor bas and u dont have SaBrE's permission, ur in deep crap

'                                               ______
'      ______________________________________  /
'     /--------------------------------------\/|||||\)\
'    /________________________________SaBrE__/\|||||/)/

