Table of Whatevers:

1. Intro
2. Basic Use
3. Events
4. Other Stuff
5. More Stuff


•Intro:
ok, this is version 1.2 of scrambler score keeper.
i fixed that smart mistake i made. (lol) it now can send scores that 
works with any version of aol, and you can specify the "sendchat"
method you want.

•Basic Use:
the use hasn't changed much. To add a name simply use this code:

score1.addnameandscore "the name", 6

i've noticed that my control acts funny sometimes when you use parenthesis
like:
score1.addnameandscore ("the name",6)

it's will give you an error. just remove the parenthesis. here's the syntax:
score1.addnameandscore Name, Score

easy huh?


•Events:
It responds to numerous events. here's a list of the one's i added:
Event Click()
	Responds when control is clicked
Event DblClick()
	Responds when control is double clicked
Event KeyDown(KeyCode As Integer, Shift As Integer)
	Responds when a key is pressed down
Event KeyPress(KeyAscii As Integer)
	Responds when a key is pressed period
Event KeyUp(KeyCode As Integer, Shift As Integer)
	Responds when a key is let up
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
	Responds when the mouse button is clicked down
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 
	Responds when the mouse moves over the control
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 
	Responds when the mouse button is realesed
Event SendScore(Score As String)
	Responds when the sub SendScore is activated

•Other Stuff:
How to send the score. I forgot to add this in version 1.0 so here's how
i fixed it:

start by activating the sub SendScore:
Score1.SendScore

then in the event SendScore put code like this:
dim b as integer

b%=b%+1
if b% < 4 then
sendchat "yourASCIIº°" & Score
elseif b% = 4 then
timeout 2.6
sendchat "yourASCIIº°" & Score
b% = 0
end if


a little more complicated than adding a name but pretty easy.


some of the properties act really stupid so don't gripe to me if they
don't work right. i don't know why. some do work though, so before
basing an application on a property test it out first.
i don't really provide help with properties unless it's directly related
to the use of the control. sorry.

I've also included an example of use with this control. look at that to see
how it's done, and i don't mind if you just copy and paste the code.  





•Other Stuff:
i'm going to be improving on this control and making the stupid properties 
work ALL of the time. i'm sorry for any trouble or frustration i've caused 
but you must realize something. i've realesed this as freeware and so it
comes with no guarentee or warenty.
if you have any suggestions, questions, comments, or need help
you can e-mail/IM me at: XSuprGeekX@aol.com  

-SuprGeek



ps i can't spel fur beens





 
			a fruity monkey production © ® 1998