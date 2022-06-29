	 _______________________________
	|  file name:  "chat.ocx"       |
	|  author:     Top dawG         |
	|  version:    2.09             |
	|  language:   visual basic 5.0 |
	 ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
_____________
introduction: 
¯¯¯¯¯¯¯¯¯¯¯¯¯
this control basically "subclasses" the america online chatroom window
and extracts from it the last line of chat. having done that, you now
have access to the built-in subs of the control to exract only the
screen name and/or message portion of that extracted line. there is
also a built-in sub which sends any desired text to the chatroom.

if there is any problems/feedback you would like to give, please email
me at topdavvg@geocities.com (notice the "w" is actually two "v"'s).

*  i use quotation marks around the word subclass because this control
   doesnt really subclass the chatroom hWnd, but it still accomplishes
   the same task with far greater efficiency and speed.
__________
whats new?
¯¯¯¯¯¯¯¯¯¯
	- a whole new scanning engine for the chatroom.
	- a new custom error-handling built-in procedure.
	- greater resource management and/or control.
	- auto-last-line ignore.
	- lesser but sounder code.
	- a little more user-friendly.
	- much much more...
__________________________________________
example of a normal bot w/ trigger phrase:
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
1) create a timer named "timer1".
2) add the chat.ocx custom control under the name "chat1".

Private Sub Timer1_Timer()
Dim chatsn As String      'you must dim your variables as strings
Dim chatmsg As String     'because other data types arent accepted
Dim chatline As String    'i did this to force people to dim variables!
On Error Resume Next
    chatline = chat1.getline        'retrieve the last line of chat
    If chatline = "" Then Exit Sub  'make sure its a valid string
    chatmsg = chat1.getmessage(chatline)       'get only the message
    If InStr(1, chatmsg, "i love top dawg", vbTextCompare) Then
        chatsn = chat1.getscreenname(chatline) 'get only the sn
        chat1.send chatsn & ", don't we all?"  'send the phrase
    End If
End Sub

Private Sub chat1_errorhandling(message As String)
    MsgBox message, vbCritical
End Sub
_______________
important note:
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
   although this code works fine, this is simply an example. in
   no way is this the only use for this control, or the best way to use
   this control. thank you and enjoy!