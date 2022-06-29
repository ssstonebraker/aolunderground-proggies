                    ArcFade 1.0  By:Arc   3/28/99


Ok.. Now there is a lot you need to know about this .Bas, i will 
probably leave something out, but i'll tell you everyhting i can think 
of right now. 

1. This .Bas can "PulseFade" or "FlashFade" any Object in VB 
that supports the ForeColor and BackColor property.

2.If you want a continuess Flash or Pulse, put the code in a loop
like this (Always use the Call Keyword) 
Call PulseFadeBack_Yellow_Black(Form1,0.0001) or
Call PulseFadeFore_Yellow_Black(Form1,0.0001)
The "PulseFade" Back pulses the backcolor the "PulseFade" Fore
pulses the forecolor.

3. The 0.0001 is the "FlashSpeed". In the Pulse Mode the 
"FlashSpeed" needs to be faster because thats just how the 
code worked out. Hehe... Sorry,The Pulse is not that fast.
In the Flash mode however the "FlashSpeed" should be set
to around 0.1. The Flash subs move Faster.

4. If you want to Pulse Or Flash a lot of items on 1 form
there is a catch. Say you wanted to "PulseFade" a label
a textbox and your forms BackGroung. Well the code would
look like this.
Do
Call PulseFadeBack_Blue_Purple(Label1, 0.00001)
Call PulseFadeBack_Blue_Purple(Text1, 0.00001)
Call PulseFadeBack_Blue_Purple(ArcFadeFrm, 0.00001)
Loop
The problem is, it wouldn't Pulse all 3 Objects at the 
same time. it would Pulse the Label then the TextBox and 
then the Form. One at a time.
I suggest if you are going to do this then you set your
objects Back or Fore Colors to the starting color position,
that way they are not siiting there just being Grey or
whatever until the code Kicks in.

5. You don't have to use these subs in a Loop. You can 
just have a code like this

Private Sub Command1_Click()
Call FlashFadeBack_Purple(Command1, 0.1)
End Sub

This code will Cause the Button to Flash only Once when 
the Button is clicked. Or the Label or whatever.

6. If you are going to have a continuess Pulse or Flash
I do not suggest you put the code in the Form_Paint event.
The reason is, if anythijng passes over it , it will freeze.
Because The Form_Paint repaints the Form every time something
is moved over it, causing an infinite loop. This is not good.

Well that is all i can think of right now. I'm sure as 
soon as i send this out i will think of ten other things
but for now just have fun with this until ArcFade v2.0.

Future versions of this .Bas will be more like
MonK-E-Fade, meaning you will be able to set your own
colors. Plus the next version will have 3 color fades
so you can fade your Objects from Blue to green to red
then back to green and then Blue. This .Bas comes with
a Form example, so if you didn't get it, ask around
it will help you see what this .Bas can do.
Don't let me limit you though, this .Bas has
limitless possibilities.

-Arc