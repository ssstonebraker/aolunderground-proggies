 ____________________________________
[Quirk's AIM ActiveX Control By Quirk]
 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
About: This is a ActiveX Control [OCX] file that I made to be used with
¯¯¯¯¯¯ AOL Instant Messenger Version 2.0 [AIM] and AOL Instant Messenger Sysop [Sysop].  
       It does all of the basic functions a user does on AIM and more.

Functions:  
¯¯¯¯¯¯¯¯¯¯
AboutBox 		[Shows my about box]
AddBuddList		[Adds the BuddyList to a ListBox or ComboBox][AIM2 ONLY]
ChatSend		[Send text to the Chat Room]
GetUser			[Gets the current users screen name]
IM			[Sends a Instant Message]
Invitation		[Sends a Buddy Chat Invitation]
Online			[Checks if the user is signed on]
RunMenu			[Runs a AIM Menu Item][AIM2 ONLY]
AddRoom			[Adds the Chat Room to a ListBox or ComboBox][AIM2 ONLY]
ChatName		[Returns the name of the current chat room]
FindBuddyList		[Returns the handle of the Buddy List Window]
FindChat		[Returns the handle of the current chat room]
FindIM			[Returns the handle of the current Instant Message]
Version			[Returns whether the user is on AIM2 or Sysop]
ChatOn			[Turns the Chat Moniter On]
ChatOff			[Turns the Chat Moniter Off]

ChatLastLine:
¯¯¯¯¯¯¯¯¯¯¯¯¯
This is my pride and joy.  It is just like a AIM Version of Dos's AOL4 Chat Control. You use it
same way.

Eample:
¯¯¯¯¯¯¯
Sub Command1_Click()
if Comamnd1.Caption = "On" then
	AIM1.ChatOn
	Command1.Caption = "Off"
Else
	AIM1.ChatOff
	Command1.Caption = "On"
End if

Private Sub AIM1_ChatLastLine(Who As String, what As String)
If LCase(what) = "sup" then
	AIM1.ChatSend "SuP " & who
End If



Quirk [AIM: XX Quirk | Mail: Quirk@NetZero.Net]