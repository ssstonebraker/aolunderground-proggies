Attribute VB_Name = "ArcFade"
'ArcFade 1.0 By:Arc 3/28/99    E-Mail=lllArc@HotMail.com

'This is the first Fader.Bas that i have seen, that does this.
'This .Bas has two kinds of Fades. The first is the PulseFade.
'The PulseFade,(as will the FlashFade) will fade any Object
'including(Forms,Labels,Buttons,Pictureboxes, any Object
'in VB that has a ForeColor or BackColor property.
'The PulseFade fades your Objects from one color to another
'and then back again. The FlashFade fades the color down a
'little and then starts over again. Creating a Flash effect.
'I have added Both,  Fading  the BackColor of an
'Object or the Forecolor(usually means the Text Of an Object)
'I have also added the option of FadeSpeed, which is the
'pause speed.
'Please refer to the ReadMe file that came with this .Bas
'for specific instructions on how to use this .Bas.
'It gives you many ideas on how to use this.Bas
'Plus instructions on how not to use this .Bas!
'If not used correctly you will run into some problems!
'Future versions of this .Bas will be more like
'MonK-E-Fade, meaning you will be able to set your own
'colors. Plus the next version will have 3 color fades
'so you can fade your Objects from Blue to green to red
'then back to green and then Blue. This .Bas comes with
'a Form example, so if you didn't get it, ask around
'it will help you see what this .Bas can do.
'Don't let me limit you though, this .Bas has
'limitless possibilities.
'-Arc


Public Red As Long, Green As Long, Blue As Long, Color As Long
Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
 DoEvents
Loop
End Sub

    'This is the PulseFade BackColor Section
Sub PulseFadeBack_Red_Black(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Blue_Black(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Green_Black(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Purple_Black(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Yellow_Black(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_White_Black(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Red_White(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Blue_White(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Green_White(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Purple_White(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Yellow_White(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Black_White(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Blue_Red(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Green_Red(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Purple_Red(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Yellow_Red(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_White_Red(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Black_Red(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Red_Green(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Blue_Green(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Purple_Green(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Yellow_Green(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_White_Green(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Black_Green(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Red_Purple(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Blue_Purple(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Green_Purple(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Yellow_Purple(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_White_Purple(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Black_Purple(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Red_Yellow(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Blue_Yellow(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Green_Yellow(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Purple_Yellow(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_White_Yellow(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub
Sub PulseFadeBack_Black_Yellow(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
End Sub

'This is the beggining of the PulseFade ForeColor Section.


Sub PulseFadeFore_Red_Black(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Blue_Black(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Green_Black(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Purple_Black(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Yellow_Black(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_White_Black(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Red_White(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Blue_White(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Green_White(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Purple_White(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Yellow_White(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Black_White(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Blue_Red(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Green_Red(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Purple_Red(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Yellow_Red(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_White_Red(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Black_Red(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Red_Green(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Blue_Green(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Purple_Green(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Yellow_Green(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_White_Green(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Black_Green(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Red_Purple(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Blue_Purple(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Green_Purple(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Yellow_Purple(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_White_Purple(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Black_Purple(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Blue = Blue - 5

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Red_Yellow(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Blue_Yellow(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Green_Yellow(OB As Object, FadeSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Purple_Yellow(OB As Object, FadeSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Green = Green + 5
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Green = Green - 5
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_White_Yellow(OB As Object, FadeSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Blue = Blue + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub
Sub PulseFadeFore_Black_Yellow(OB As Object, FadeSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 50
Pause (FadeSpeed)
Red = Red + 5
Green = Green + 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
For i = 0 To 50
Pause (FadeSpeed)
Red = Red - 5
Green = Green - 5
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
End Sub


'This is the begginig of the FlashFade BackColor Section.

Sub FlashFadeBack_Black(OB As Object, FlashSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)
Red = Red + 10
Green = Green + 10
Blue = Blue + 10
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 0
Green = 0
Blue = 0
End Sub

Sub FlashFadeBack_White(OB As Object, FlashSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10
Green = Green - 10
Blue = Blue - 10
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 255
Green = 255
Blue = 255
End Sub
Sub FlashFadeBack_Red(OB As Object, FlashSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 255
Green = 0
Blue = 0
End Sub
Sub FlashFadeBack_Blue(OB As Object, FlashSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 10
Pause (FlashSpeed)

Blue = Blue - 10
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 0
Green = 0
Blue = 255
End Sub
Sub FlashFadeBack_Green(OB As Object, FlashSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)

Green = Green - 10

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 0
Green = 255
Blue = 0
End Sub
Sub FlashFadeBack_Purple(OB As Object, FlashSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10

Blue = Blue - 10
Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 255
Green = 0
Blue = 255
End Sub
Sub FlashFadeBack_Yellow(OB As Object, FlashSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10
Green = Green - 10

Color = RGB(Red, Green, Blue)
OB.BackColor = Color
Next i
Red = 255
Green = 255
Blue = 0
End Sub

'This is the beggining of the FlashFade ForeColor Section

Sub FlashFadeFore_Black(OB As Object, FlashSpeed)
Red = 0
Green = 0
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)
Red = Red + 10
Green = Green + 10
Blue = Blue + 10
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 0
Green = 0
Blue = 0
End Sub

Sub FlashFadeFore_White(OB As Object, FlashSpeed)
Red = 255
Green = 255
Blue = 255
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10
Green = Green - 10
Blue = Blue - 10
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 255
Green = 255
Blue = 255
End Sub
Sub FlashFadeFore_Red(OB As Object, FlashSpeed)
Red = 255
Green = 0
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 255
Green = 0
Blue = 0
End Sub
Sub FlashFadeFore_Blue(OB As Object, FlashSpeed)
Red = 0
Green = 0
Blue = 255
For i = 0 To 10
Pause (FlashSpeed)

Blue = Blue - 10
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 0
Green = 0
Blue = 255
End Sub
Sub FlashFadeFore_Green(OB As Object, FlashSpeed)
Red = 0
Green = 255
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)

Green = Green - 10

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 0
Green = 255
Blue = 0
End Sub
Sub FlashFadeFore_Purple(OB As Object, FlashSpeed)
Red = 255
Green = 0
Blue = 255
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10

Blue = Blue - 10
Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 255
Green = 0
Blue = 255
End Sub
Sub FlashFadeFore_Yellow(OB As Object, FlashSpeed)
Red = 255
Green = 255
Blue = 0
For i = 0 To 10
Pause (FlashSpeed)
Red = Red - 10
Green = Green - 10

Color = RGB(Red, Green, Blue)
OB.ForeColor = Color
Next i
Red = 255
Green = 255
Blue = 0
End Sub
