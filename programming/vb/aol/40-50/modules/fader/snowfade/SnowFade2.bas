Attribute VB_Name = "SnowFade2"
'•SnowFade2.BAS•
'by Equinox

'Made For: (Visual Basic 4.0, 5.0, 6.0)

'contacts:
'(e-mail me at: nitrogen snow@n2mail.com) or
'(iM Me at:Iiquidsnow,(thats my AIM sn))

'version:2
'updates:
'The updates in this bas iz that i stoped usin the
'cryo type fades, even tho that they are cool as hell
'i found out you can do more like this than like cryo's
'subs. I only took 1 sub from monkefade cause hes ieet as
'hell. well anyways later..
'Equinox

'Please do not steal our subs and functions,
'there is no reason to add them to your
'bas, why not just use my bas too instead
'of being a code thief.  And also please
'add me to your greets, especially if
'your prog is just a fader, I mean with
'this bas you could make a really leet
'fader very very easily.

'i think this .bas will be up there with the ieet
'fader .bas's like monk-e-fade3 and cryofade
'well atleast i hope so..

'this bas is fairly simple to use
'all you have to do is basicly what you do
'in monk-e-fade3.bas except theres more colors to
'work with

'LaTeZ
'-Equinox
Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000

Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type


Function ScrollBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This Sub Was Not Taken From Monk-E-Fade.Bas
'Example On How To Use This Sub
'Label#.BackColor = RGB (Red#.Value, Green#.Value, Blue#.Value)
'And Put The Scroll Bar Max To : 255
Dim ScrollBar As String
ScrollBar = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
End Function
Function ReverseBoldWavy(Text As String, Bold As Boolean, Wavy As Boolean)
'This sub allows u to quickly make text bold or wavy or both
'IT DOES NOT WORK ON FADED TEXT!!!
'Example:
'SendChat ReverseBoldWavy ("the text you want",True,False")
'And you can also put ("true,true") or what ever

Dim B%
Dim W%
Dim X%
Dim T$
B% = 1
W% = 1
For X% = 1 To Len(Text)
If Bold = True Then
Select Case B%
Case 1
T$ = T$ + "<B>"
Case 2
T$ = T$ + "</B>"
End Select
End If
If Wavy = True Then
Select Case W%
Case 1
T$ = T$ + "<SUB>"
Case 2
T$ = T$ + "</SUB>"
Case 3
T$ = T$ + "<SUP>"
Case 4
T$ = T$ + "</SUP>"
End Select
End If
T$ = T$ + Mid$(Text, X%, 1)
Select Case B%
Case 1
B% = 2
Case 2
B% = 1
End Select
Select Case W%
Case 1
W% = 2
Case 2
W% = 3
Case 3
W% = 4
Case 4
W% = 1
End Select
Next X%
RBW = T$
End Function

Function UnderlineWavy(Text As String, Underline As Boolean, Wavy As Boolean)
'just the same as ReverseBoldWavy
Dim U%
Dim W%
Dim X%
Dim T$
S% = 1
W% = 1
For X% = 1 To Len(Text)
If Underline = True Then
Select Case U%
Case 1
T$ = T$ + "<u>"
Case 2
T$ = T$ + "</u>"
End Select
End If
If Wavy = True Then
Select Case W%
Case 1
T$ = T$ + "<SUB>"
Case 2
T$ = T$ + "</SUB>"
Case 3
T$ = T$ + "<SUP>"
Case 4
T$ = T$ + "</SUP>"
End Select
End If
T$ = T$ + Mid$(Text, X%, 1)
Select Case U%
Case 1
U% = 2
Case 2
U% = 1
End Select
Select Case W%
Case 1
W% = 2
Case 2
W% = 3
Case 3
W% = 4
Case 4
W% = 1
End Select
Next X%
UnderlineWavy = T$
End Function

Function WavyStrike(Text As String, Strike As Boolean, Wavy As Boolean)
'Same rules apply for this as the ReverseBoldWavy sub

Dim S%
Dim W%
Dim X%
Dim T$
S% = 1
W% = 1
For X% = 1 To Len(Text)
If Strike = True Then
Select Case S%
Case 1
T$ = T$ + "<S>"
Case 2
T$ = T$ + "</S>"
End Select
End If
If Wavy = True Then
Select Case W%
Case 1
T$ = T$ + "<SUB>"
Case 2
T$ = T$ + "</SUB>"
Case 3
T$ = T$ + "<SUP>"
Case 4
T$ = T$ + "</SUP>"
End Select
End If
T$ = T$ + Mid$(Text, X%, 1)
Select Case S%
Case 1
S% = 2
Case 2
S% = 1
End Select
Select Case W%
Case 1
W% = 2
Case 2
W% = 3
Case 3
W% = 4
Case 4
W% = 1
End Select
Next X%
StrikeWavy = T$
End Function

Function EightColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 7 <> 0 Then
    Do Until Len(Text) Mod 7 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 7
Eight1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Eight2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Eight3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Eight4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Eight5 = TwoColors(Mid(Text, P + P + P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Eight6 = TwoColors(Mid(Text, P + P + P + P + P + 1, P), Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Eight7 = TwoColors(Right(Text, P), Red7, Green7, Blue7, Red8, Green8, Blue8, False)
EightColors = Eight1 + Eight2 + Eight3 + Eight4 + Eight5 + Eight6 + Eight7
If Wavy = True Then
For X% = 1 To Len(EightColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(EightColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(EightColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(EightColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(EightColors, X% + 63, 21)
Next X%
EightColors = TextX$
End If
End Function


Function ElevenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Wavy As Boolean)

If Len(Text) < 6 Then
    Do Until Len(Text) = 6
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 5 <> 0 Then
    Do Until Len(Text) Mod 5 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 5
Eleven1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Eleven2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Eleven3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Eleven4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Eleven5 = ThreeColors(Right(Text, P), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, False)
ElevenColors = Eleven1 + Eleven2 + Eleven3 + Eleven4 + Eleven5
If Wavy = True Then
For X% = 1 To Len(ElevenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(ElevenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(ElevenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(ElevenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(ElevenColors, X% + 63, 21)
Next X%
ElevenColors = TextX$
End If
End Function

Function FiveColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)

If Len(Text) < 3 Then
    Do Until Len(Text) = 3
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 2 <> 0 Then
    Do Until Len(Text) Mod 2 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 2
Five1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Five2 = ThreeColors(Right(Text, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
FiveColors = Five1 + Five2
If Wavy = True Then
For X% = 1 To Len(FiveColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(FiveColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(FiveColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(FiveColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(FiveColors, X% + 63, 21)
Next X%
FiveColors = TextX$
End If
End Function

Function FourColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 3 <> 0 Then
    Do Until Len(Text) Mod 3 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 3
Four1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Four2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Four3 = TwoColors(Right(Text, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
FourColors = Four1 + Four2 + Four3
If Wavy = True Then
For X% = 1 To Len(FourColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(FourColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(FourColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(FourColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(FourColors, X% + 63, 21)
Next X%
FourColors = TextX$
End If
End Function

Function ReverseBold(TextBox As String, Wavy As Boolean)
Dim Text$
Dim X%
Dim Tot%
Tot% = Len(TextBox)
If Wavy = False Then
For X% = 1 To Tot% Step 2
Text$ = Text$ + "<b>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</b>" + Mid(TextBox, X% + 1, 1)
Next X%
End If
If Wavy = True Then
For X% = 1 To Tot% Step 4
Text$ = Text$ + "<sub><b>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</sub></b>" + Mid(TextBox, X% + 1, 1)
Text$ = Text$ + "<sup><b>" + Mid(TextBox, X% + 2, 1)
Text$ = Text$ + "</sup></b>" + Mid(TextBox, X% + 3, 1)
Next X%
End If
ReverseBold = Text$
End Function

Function ReverseBoldFadedEight(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = EightColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedEight = Text$
End Function

Function ReverseBoldFadedFive(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FiveColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedFive = Text$
End Function

Function ReverseBoldFadedFour(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FourColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedFour = Text$
End Function

Function ReverseBoldFadedNine(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = NineColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedNine = Text$
End Function

Function ReverseBoldFadedSeven(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SevenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedSeven = Text$
End Function

Function ReverseBoldFadedSix(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SixColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedSix = Text$
End Function

Function ReverseBoldFadedTen(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedTen = Text$
End Function

Function ReverseBoldFadedThree(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = ThreeColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedThree = Text$
End Function
Function ReverseBoldFadedTwo(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TwoColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldFadedTwo = Text$
End Function

Function ReverseBoldItalic(TextBox As String, Wavy As Boolean)
Dim Text$
Dim X%
Dim Tot%
Tot% = Len(TextBox)
If Wavy = False Then
For X% = 1 To Tot% Step 2
Text$ = Text$ + "<b></i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</b><i>" + Mid(TextBox, X% + 1, 1)
Next X%
End If
If Wavy = True Then
For X% = 1 To Tot% Step 4
Text$ = Text$ + "<sub><b></i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</sub></b><i>" + Mid(TextBox, X% + 1, 1)
Text$ = Text$ + "<sup></i><b>" + Mid(TextBox, X% + 2, 1)
Text$ = Text$ + "</sup></b><i>" + Mid(TextBox, X% + 3, 1)
Next X%
End If
ReverseBoldItalic = Text$
End Function
Function ReverseBoldItalicFadedEight(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = EightColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedEight = Text$
End Function

Function ReverseBoldItalicFadedFive(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FiveColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedFive = Text$
End Function
Function ReverseBoldItalicFadedFour(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FourColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedFour = Text$
End Function
Function ReverseBoldItalicFadedNine(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = NineColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedNine = Text$
End Function

Function ReverseBoldItalicFadedSeven(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SevenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedSeven = Text$
End Function
Function ReverseBoldItalicFadedSix(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SixColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedSix = Text$
End Function

Function ReverseBoldItalicFadedTen(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedTen = Text$
End Function
Function ReverseBoldItalicFadedThree(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = ThreeColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedThree = Text$
End Function

Function ReverseBoldItalicFadedTwo(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TwoColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></b><i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><b></i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></b><i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<b></i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</b><i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseBoldItalicFadedTwo = Text$
End Function

Function ReverseItalic(TextBox As String, Wavy As Boolean)
Dim Text$
Dim X%
Dim Tot%
Tot% = Len(TextBox)
If Wavy = False Then
For X% = 1 To Tot% Step 2
Text$ = Text$ + "<i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</i>" + Mid(TextBox, X% + 1, 1)
Next X%
End If
If Wavy = True Then
For X% = 1 To Tot% Step 4
Text$ = Text$ + "<sub><i>" + Mid(TextBox, X%, 1)
Text$ = Text$ + "</sub></i>" + Mid(TextBox, X% + 1, 1)
Text$ = Text$ + "<sup><i>" + Mid(TextBox, X% + 2, 1)
Text$ = Text$ + "</sup></i>" + Mid(TextBox, X% + 3, 1)
Next X%
End If
ReverseItalic = Text$
End Function

Function ReverseItalicFadedEight(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = EightColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedEight = Text$
End Function
Function ReverseItalicFadedFive(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FiveColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedFive = Text$
End Function

Function ReverseItalicFadedFour(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = FourColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedFour = Text$
End Function
Function ReverseItalicFadedNine(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = NineColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedNine = Text$
End Function

Function ReverseItalicFadedSeven(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SevenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedSeven = Text$
End Function

Function ReverseItalicFadedSix(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = SixColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedSix = Text$
End Function

Function ReverseItalicFadedTen(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TenColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedTen = Text$
End Function

Function ReverseItalicFadedThree(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = ThreeColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedThree = Text$
End Function

Function ReverseItalicFadedTwo(TextBox As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
Dim Text$
Dim X%
Dim TextX$
Dim Tot%
TextX$ = TwoColors(TextBox, Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Tot% = Len(TextX$)
If Wavy = True Then
For X% = 1 To Tot% Step 84
Text$ = Text$ + "<sub><i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</sub></i>" + Mid(TextX$, X% + 21, 21)
Text$ = Text$ + "<sup><i>" + Mid(TextX$, X% + 42, 21)
Text$ = Text$ + "</sup></i>" + Mid(TextX$, X% + 63, 21)
Next X%
End If
If Wavy = False Then
For X% = 1 To Tot% Step 42
Text$ = Text$ + "<i>" + Mid(TextX$, X%, 21)
Text$ = Text$ + "</i>" + Mid(TextX$, X% + 21, 21)
Next X%
End If
ReverseItalicFadedTwo = Text$
End Function

Function RGB2HEX(r, g, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = g
        If X& = 3 Then Color& = r
        For xx& = 1 To 2
            Divide = Color& / 16
            Answer& = Int(Divide)
            Remainder& = (10000 * (Divide - Answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = Answer&
        Next xx&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function RGBtoHEX(RGB)
'i didnt make this one...
'i think it was cryo :)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function

Function SeeFade(R1, G1, B1, R2, B2, G2, pctre)
'i have often found that this will only work once,
'so for this reason i recomend u copy and paste
'the code into the Paint Proc of a picture box.
'This only shows 2 colors faded at a time.

On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double

Static SplitNum(3) As Double
Static DivideNum(3) As Double

Dim FadeW As Integer
Dim Loo As Integer
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100
FadeW = pctre.Width / 100
For Loo = 0 To 100

pctre.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)

Next Loo

End Function

Function SevenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 6 <> 0 Then
    Do Until Len(Text) Mod 6 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 6
Seven1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Seven2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Seven3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Seven4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Seven5 = TwoColors(Mid(Text, P + P + P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
Seven6 = TwoColors(Right(Text, P), Red6, Green6, Blue6, Red7, Green7, Blue7, False)
SevenColors = Seven1 + Seven2 + Seven3 + Seven4 + Seven5 + Seven6
If Wavy = True Then
For X% = 1 To Len(SevenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(SevenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(SevenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(SevenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(SevenColors, X% + 63, 21)
Next X%
SevenColors = TextX$
End If
End Function

Function SixColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Wavy As Boolean)

If Text = "" Then Text = " "
If Len(Text) Mod 5 <> 0 Then
    Do Until Len(Text) Mod 5 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 5
Six1 = TwoColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, False)
Six2 = TwoColors(Mid(Text, P + 1, P), Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Six3 = TwoColors(Mid(Text, P + P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, False)
Six4 = TwoColors(Mid(Text, P + P + P + 1, P), Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Six5 = TwoColors(Right(Text, P), Red5, Green5, Blue5, Red6, Green6, Blue6, False)
SixColors = Six1 + Six2 + Six3 + Six4 + Six5
If Wavy = True Then
For X% = 1 To Len(SixColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(SixColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(SixColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(SixColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(SixColors, X% + 63, 21)
Next X%
SixColors = TextX$
End If
End Function


Function TenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Wavy As Boolean)

If Len(Text) < 6 Then
    Do Until Len(Text) = 6
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 5 <> 0 Then
    Do Until Len(Text) Mod 5 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 5
Ten1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Ten2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Ten3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Ten4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Ten5 = TwoColors(Right(Text, P), Red9, Green9, Blue9, Red10, Green10, Blue10, False)
TenColors = Ten1 + Ten2 + Ten3 + Ten4 + Ten5
If Wavy = True Then
For X% = 1 To Len(TenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(TenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(TenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(TenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(TenColors, X% + 63, 21)
Next X%
TenColors = TextX$
End If
End Function

Function ThirteenColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Red12, Green12, Blue12, Red13, Green13, Blue13, Wavy As Boolean)

If Len(Text) < 7 Then
    Do Until Len(Text) = 7
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 6 <> 0 Then
    Do Until Len(Text) Mod 6 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 6
Thirteen1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Thirteen2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Thirteen3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Thirteen4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Thirteen5 = ThreeColors(Mid(Text, P + P + P + P + 1, P), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, False)
Thirteen6 = ThreeColors(Right(Text, P), Red11, Green11, Blue11, Red12, Green12, Blue12, Red13, Green13, Blue13, False)
ThirteenColors = Thirteen1 + Thirteen2 + Thirteen3 + Thirteen4 + Thirteen5 + Thirteen6
If Wavy = True Then
For X% = 1 To Len(ThirteenColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(ThirteenColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(ThirteenColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(ThirteenColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(ThirteenColors, X% + 63, 21)
Next X%
ThirteenColors = TextX$
End If
End Function

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)

    d = Len(Text)
        If d = 0 Then GoTo theEnd
        If d = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If d = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If d = X Then GoTo Odds
    Next X
Evens:
    C = d \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C)
    GoTo theEnd
Odds:
    C = d \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C + 1)
theEnd:
    LA1 = Fade1
    LA2 = Fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    Msg = FadeA + FadeB
    ThreeColors = Msg
End Function

Function TwelveColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, Red12, Green12, Blue12, Wavy As Boolean)

If Len(Text) < 7 Then
    Do Until Len(Text) = 7
        Text = Text + " "
    Loop
End If
If Len(Text) Mod 6 <> 0 Then
    Do Until Len(Text) Mod 6 = 0
        Text = Text + " "
    Loop
End If
P = Len(Text) / 6
Twelve1 = ThreeColors(Left(Text, P), Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, False)
Twelve2 = ThreeColors(Mid(Text, P + 1, P), Red3, Green3, Blue3, Red4, Green4, Blue4, Red5, Green5, Blue5, False)
Twelve3 = ThreeColors(Mid(Text, P + P + 1, P), Red5, Green5, Blue5, Red6, Green6, Blue6, Red7, Green7, Blue7, False)
Twelve4 = ThreeColors(Mid(Text, P + P + P + 1, P), Red7, Green7, Blue7, Red8, Green8, Blue8, Red9, Green9, Blue9, False)
Twelve5 = ThreeColors(Mid(Text, P + P + P + P + 1, P), Red9, Green9, Blue9, Red10, Green10, Blue10, Red11, Green11, Blue11, False)
Twelve6 = TwoColors(Right(Text, P), Red11, Green11, Blue11, Red12, Green12, Blue12, False)
TwelveColors = Twelve1 + Twelve2 + Twelve3 + Twelve4 + Twelve5 + Twelve6
If Wavy = True Then
For X% = 1 To Len(TwelveColors) Step 84
TextX$ = TextX$ + "<sub>" + Mid$(TwelveColors, X%, 21)
TextX$ = TextX$ + "</sub>" + Mid$(TwelveColors, X% + 21, 21)
TextX$ = TextX$ + "<sup>" + Mid$(TwelveColors, X% + 42, 21)
TextX$ = TextX$ + "</sup>" + Mid$(TwelveColors, X% + 63, 21)
Next X%
TwelveColors = TextX$
End If
End Function

'Variable color fade functions begin here


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
    C = 0
    o = 0
    o2 = 0
    q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(Text) * X) + Red1
        VAL2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
        C1 = RGB2HEX(VAL1, VAL2, VAL3)
        C2 = RGB2HEX(VAL1, VAL2, VAL3)
        C3 = RGB2HEX(VAL1, VAL2, VAL3)
        C4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then C = 1: Msg = Msg & "<FONT COLOR=#" + C1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If C <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
        End If
        
        If Wavy = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(Text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf Wavy = False Then
            Msg = Msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next X
    C1 = C1BAK
    C2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    TwoColors = Msg
End Function

Sub FadePreview(PreTxtMain As Control, FadedText As String, PreTxt As TextBox)
'by monk-e-god
'-FADE PREVIEW-
'To use the fadepreview you need a
'rich textbox (which requires an ocx)
'and a regular text box in which the
'HTML will be interpreted.

'example:
'HTMLbox.Text = FadeByColor4(FADE_RED, FADE_BLACK, FADE_GREY, FADE_GREEN, "Red/Black/Grey/Green Fade Preview", False)
'Call FadePreview(PreviewBox, HTMLbox.Text, InvisBox)

'now in the rich textbox, PreviewBox, you
'will see a Red to Black to Grey to Green
'fade saying "Red/Black/Grey/Green Fade Preview"

'NOTE: You cannot preview wavy fades.
'NOTE: PreTxtMain MUST be a rich textbox!

PreTxtMain.Text = ""
Dim Starts()
Dim Lengths()
Dim Colors()
Dim LastHtml%
Dim CurStart%
Dim CurLen%
Dim CurColor$
Dim NumFades%
PreTxt.Text = FadedText
NumFades% = 0
LastHtml% = 2
findhtml% = 1
While findhtml%
If NumFades% = 0 Then findhtml% = 0

NumFades% = NumFades% + 1
findhtml% = InStr(findhtml% + 1, PreTxt.Text, "<Font Color=#") 'InStr(LastHtml - 1, PreTxt.Text, "<Font Color=#")
If findhtml% = 0 Then GoTo Blah
LastHtml% = InStr(findhtml% + 1, PreTxt.Text, ">")
thecolor = Mid(PreTxt.Text, findhtml% + 13, 6)
htmlblue$ = Right(thecolor, 2)
htmlgreen$ = Mid(thecolor, 3, 2)
htmlred$ = Left(thecolor, 2)
vbcolor = "&H00" + htmlblue$ + htmlgreen$ + htmlred$ + "&"

nexthtml% = InStr(findhtml% + 1, PreTxt.Text, "<Font Color=#")
CurLen% = 1
Firstpart$ = Left(PreTxt.Text, findhtml% - 1)
Secondpart$ = Mid(PreTxt.Text, LastHtml% + 1)
PreTxt.Text = Firstpart$ + Secondpart$
CurStart% = findhtml%
CurColor = vbcolor
ReDim Preserve Starts(NumFades%)
ReDim Preserve Lengths(NumFades%)
ReDim Preserve Colors(NumFades%)
Starts(NumFades%) = CurStart% - 1
Lengths(NumFades%) = CurLen%
Colors(NumFades%) = CurColor

Blah:
H = H
Wend
PreTxtMain.Text = PreTxt.Text
For cc% = 1 To NumFades% - 1
PreTxtMain.SelStart = Starts(cc%)
PreTxtMain.SelLength = Lengths(cc%)
PreTxtMain.SelColor = Val(Colors(cc%))
H = H
Next cc%
PreTxtMain.SelLength = 0

End Sub

