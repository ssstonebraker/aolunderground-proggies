What it is: Quirk's AIM ActiveX Control v4
Made By: Quirk
What .ocx's .bas's used: None
What Visual Basic it was made on: Visual Basic 6 Enterprise

Some Codes:

AIM1.AboutBox 			Shows the about box
AIM1.AddBuddList(List, false) 	Adds the buddylist to a combo or a list, true to add yourself to the list, false if not [aim2 only]
AIM1.ChatSend(TheText) 		Sends text to the chat room
X = AIM1.GetUser 			Gets the users SN
AIM1.IM(who, Message) 		Sends a IM
AIM1.IM2(who, Message) 		Sends a IM a better way [aim2.0.996+ only]
AIM1.Invitation(who,message,room) 	Sends a buddy chat invitation
X = AIM1.Online 			Returns true if the user is online, false if not
AIM1.AddRoom(List, false)		Adds the chatroom list to a vombo or list, true to add yourself, false if not
X = AIM1.ChatName		Returns the name of the current chat room
X = AIM1.FindBuddyList		Returns the handle of the buddylist
X = AIM1.FindChat			Returns the handle of the current chat room
X = AIM1.FindIM			Returns the handle of the current IM
X = AIM1.Version			Returns "2" if the user is on AIM2, and "Sysop" if the user is on Sysop
AIM1.ChatOn			Turns the Chat Moniter On[AIM2+ only]
AIM1.ChatOff			Turns the Chat Moniter Off[AIM2+ only]
AIM1.ClearChat			Clears the Current Chat Room Text
AIM1.CloseChat(RoomName)	Closes the current room if RoomName is omited, else closes roomname
AIM1.CloseIM(Person)		Closes the current im if Person is omited, else closes the im from person
X = AIM1.IMLastLine		Returns the last line of the current IM
X = AIM1.IMSender		Returns the sender of the current im
X = AIM1.IMText			Returns the entire IM text
AIM1.SearchBar(String)		This searches the internet via the aim searchbar and can be used for keywords[aim2.0.996+ only]
AIM1.SignOff			Ends the current AIM session [aim2+ only]

ChatLastLine:
¯¯¯¯¯¯¯¯¯¯¯¯¯
This is my pride and joy.  It is just like a AIM Version of Dos's AOL4 Chat Control. You use it same way. I am sorry but this only works for AIM2. I have made this alot faster and more effecient, still does not work for sysop.

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