Attribute VB_Name = "Enlightenment"
'//welcome to Crypto's shrine of enlightenment. learn well.
'//some of you think you have to put the cursor over the icon to click it...tsk tsk tsk
'//this is simply just incorrect and kind of funny to me. all you need to do is think like a program.
'//yes i know it sounds funny and down right disturbing, but just think for a moment....

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
'//oh yeah and i'll also bestow upon you the proper way to use sendmessage, yes the real way...
'//not the lazy persons way and the way it's been done for years <<shudders>>

'//note this works for SOME msgbox buttons, but it always works for the _AOL_Icon class
Public Sub Click_Button(Icon As Long)   '//clicks either a Button or an _AOL_Icon
    Call SendMessage(Icon, WM_LBUTTONDOWN, 0&, 0&)              '//sets focus to the object (this is to beat some of the stuff on AOL6)
    Call SendMessage(Icon, WM_KEYDOWN, ByVal VK_SPACE, 0&)      '//calls SendMessage and passes the VK_SPACE constant
    Call SendMessage(Icon, WM_KEYUP, ByVal VK_SPACE, 0&)        '//_AOL_Icons behave just like buttons when you hit spacebar or enter
End Sub

'//now here comes the second nugget of goodness upside ya head
'//SendMessage! by far the most used and abused API function ever!..the one with an identity crisis
'//people...you need only ONE DECLARE OF SENDMESSAGE, yes! ONE i say!
'//let's run through the ways people that don't know what they're doing do it (confusing?)
'//Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                               ByVal wMsg As Long, _
                                                                               ByVal wParam As Long, _
                                                                               ByVal lParam As Long)
'//and for a string?
'//Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                                  ByVal wMsg As Long, _
                                                                                  ByVal wParam As Long, _
                                                                                  ByVal lParam As String) As Long
'//now look at them carefully, the last argument is the one that changes
'//lParam as Any, lParam as Long and lParam as String now..notice the keyword that's repeated
'//BYVAL!!! my god use it. see in my code where i have SendMessage(Icon,WM_KEYDOWN, BYVAL VK_SPACE,0&)
'//ByVal means By Value, the value of whatever variable or expression is passed as the parameter
'//this prevents the function from changing the contents of the variable passed ByVal
'//-----------------------------------------------------------------------------------------------------
'//now say you want to send a string instead? no problemo
'//Call SendMessage(Edit,WM_SETTEXT,0&,ByVal sText)
'//for settext you dont need the wParam, its not used by that message and it must be zero, but not just any zero
'//you either have to do SendMessage(Edit,WM_SETTEXT, ByVal Clng(0), ByVal sText) or the way i have it above 0& <--the & means long duh.
'//ok i've provided 3 nuggets of value when i only said 1 in the beginning, i hope you learned something from this
'//declaring multiple variations of functions is just bad programing practice, and is also ineffecient.
'//questions? comments? want to yell at me? drop me a line, i check it once every 2 weeks
'//cryptofx@hotmail.com, yes i have a screen name, no you can't have it
'//some people just don't catch my humor, if i sounded like a prick, sorry it's called dry humor :)
'//your truly - Crypto
