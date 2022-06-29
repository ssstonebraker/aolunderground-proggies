Attribute VB_Name = "FrENzY32Misc"
' These are just some subs i found usefull
'and i didn't really feel like making myself
'but felt i should include with my bas.

#If Win32 Then
    Global SubCount As Long
#Else
    Global SubCount As Integer
#End If

Function Hex2Dec!(ByVal strHex$)
'Monke-God
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function
Function HTMLtoRGB(HTMLColor$)
'Monke-God
If Left(HTMLColor$, 1) = "#" Then HTMLColor$ = Right(HTMLColor$, 6)

RedX$ = Left(HTMLColor$, 2)
GreenX$ = Mid(HTMLColor$, 3, 2)
BlueX$ = Right(HTMLColor$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function InverseColor(OldColor)
'Monke-God
DAColor$ = RGBtoHEX(OldColor)
RedX% = Val("&H" + Right(DAColor$, 2))
GreenX% = Val("&H" + Mid(DAColor$, 3, 2))
BlueX% = Val("&H" + Left(DAColor$, 2))
newred% = 255 - RedX%
newgreen% = 255 - GreenX%
newblue% = 255 - BlueX%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function
Function RGBtoHEX(RGB)
'Monke-God
    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function
Function Trm(ByVal Txt As String)
'Monke-God
Dim X As Integer
Dim Y As String
Dim Z As String
For X = 1 To Len(Txt)
Y = Mid(Txt, X, 1)
If Y = Chr(0) Then Y = ""
Z = Z & Y
Next X
Trm = Z
End Function
Function MultiFade(NumColors%, TheColors(), TheText$, Wavy As Boolean)
'Monke-God
Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NumColors < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = TheText
Exit Function
End If

If NumColors = 1 Then
blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(blah$, 2))
greenpart% = Val("&H" + Mid(blah$, 3, 2))
bluepart% = Val("&H" + Left(blah$, 2))
blah2 = RGB(bluepart%, greenpart%, redpart%)
blah3$ = RGBtoHEX(blah2)

MultiFade = "<Font Color=#" + blah3$ + ">" + TheText
Exit Function
End If

Dim RedList%()
Dim GreenList%()
Dim BlueList%()
Dim DaColors$()
Dim DaLens%()
Dim DaParts$()
Dim Faded$()

ReDim RedList%(NumColors)
ReDim GreenList%(NumColors)
ReDim BlueList%(NumColors)
ReDim DaColors$(NumColors)
ReDim DaLens%(NumColors - 1)
ReDim DaParts$(NumColors - 1)
ReDim Faded$(NumColors - 1)

For q% = 1 To NumColors
DaColors(q%) = RGBtoHEX(TheColors(q%))
Next q%

For w% = 1 To NumColors
RedList(w%) = Val("&H" + Right(DaColors(w%), 2))
GreenList(w%) = Val("&H" + Mid(DaColors(w%), 3, 2))
BlueList(w%) = Val("&H" + Left(DaColors(w%), 2))
Next w%

TextLen% = Len(TheText)
Do: DoEvents
For F% = 1 To (NumColors - 1)
DaLens(F%) = DaLens(F%) + 1: TextLen% = TextLen% - 1
If TextLen% < 1 Then Exit For
Next F%
Loop Until TextLen% < 1
    
DaParts(1) = Left(TheText, DaLens(1))
DaParts(NumColors - 1) = Right(TheText, DaLens(NumColors - 1))
    
dastart% = DaLens(1) + 1

If NumColors > 2 Then
For E% = 2 To NumColors - 2
DaParts(E%) = Mid(TheText, dastart%, DaLens(E%))
dastart% = dastart% + DaLens(E%)
Next E%
End If

For r% = 1 To (NumColors - 1)
TextLen% = Len(DaParts(r%))
For i = 1 To TextLen%
    TextDone$ = Left(DaParts(r%), i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((BlueList(r% + 1) - BlueList(r%)) / TextLen% * i) + BlueList(r%), ((GreenList%(r% + 1) - GreenList(r%)) / TextLen% * i) + GreenList(r%), ((RedList(r% + 1) - RedList(r%)) / TextLen% * i) + RedList(r%))
    colorx2 = RGBtoHEX(ColorX)
        
    If Wavy = True Then
    WaveState = WaveState + 1
    If WaveState > 4 Then WaveState = 1
    If WaveState = 1 Then WaveHTML = "<sup>"
    If WaveState = 2 Then WaveHTML = "</sup>"
    If WaveState = 3 Then WaveHTML = "<sub>"
    If WaveState = 4 Then WaveHTML = "</sub>"
    Else
    WaveHTML = ""
    End If
        
    Faded(r%) = Faded(r%) + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
Next i
Next r%

For qwe% = 1 To (NumColors - 1)
FadedTxtX$ = FadedTxtX$ + Faded(qwe%)
Next qwe%

MultiFade = FadedTxtX$

End Function
Function Rich2HTML(RichTXT As Control, StartPos%, EndPos%)
'Monke-God
Dim Bolded As Boolean
Dim Undered As Boolean
Dim Striked As Boolean
Dim Italiced As Boolean
Dim LastCRL As Long
Dim LastFont As String
Dim HTMLString As String

For posi% = StartPos To EndPos
RichTXT.SelStart = posi%
RichTXT.SelLength = 1

If Bolded <> RichTXT.SelBold Or posi% = StartPos Then
If RichTXT.SelBold = True Then
HTMLString = HTMLString + "<b>"
Bolded = True
Else
HTMLString = HTMLString + "</b>"
Bolded = False
End If
End If

If Undered <> RichTXT.SelUnderline Or posi% = StartPos Then
If RichTXT.SelUnderline = True Then
HTMLString = HTMLString + "<u>"
Undered = True
Else
HTMLString = HTMLString + "</u>"
Undered = False
End If
End If

If Striked <> RichTXT.SelStrikeThru Or posi% = StartPos Then
If RichTXT.SelStrikeThru = True Then
HTMLString = HTMLString + "<s>"
Striked = True
Else
HTMLString = HTMLString + "</s>"
Striked = False
End If
End If

If Italiced <> RichTXT.SelItalic Or posi% = StartPos Then
If RichTXT.SelItalic = True Then
HTMLString = HTMLString + "<i>"
Italiced = True
Else
HTMLString = HTMLString + "</i>"
Italiced = False
End If
End If

If LastCRL <> RichTXT.SelColor Or posi% = StartPos Then
ColorX = RGB(GetRGB(RichTXT.SelColor).Blue, GetRGB(RichTXT.SelColor).Green, GetRGB(RichTXT.SelColor).Red)
colorhex = RGBtoHEX(ColorX)
HTMLString = HTMLString + "<Font Color=#" & colorhex & ">"
LastCRL = RichTXT.SelColor
End If

If LastFont <> RichTXT.SelFontName Then
HTMLString = HTMLString + "<font face=" + Chr(34) + RichTXT.SelFontName + Chr(34) + ">"
LastFont = RichTXT.SelFontName
End If

HTMLString = HTMLString + RichTXT.SelText
Next posi%

Rich2HTML = HTMLString

End Function
Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, TheText$, Wavy As Boolean)
'Monke-God
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, TheText, Wavy)

End Function

Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)
'Monke-God
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, Wavy)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)
'Monke-God
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, TheText$, Wavy As Boolean)
'Monke-God
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, TheText, Wavy)

End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, TheText$, Wavy As Boolean)
'Monke-GoD
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / TextLen% * i) + B3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, TheText$, Wavy As Boolean)
'Monke-God
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, TheText, Wavy)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, TheText$, Wavy As Boolean)
'Monke-God
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(TheText, frthlen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / TextLen% * i) + B3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part4
    TextLen% = Len(part4$)
    For i = 1 To TextLen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / TextLen% * i) + B4, ((G5 - G4) / TextLen% * i) + G4, ((R5 - R4) / TextLen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, TheText$, Wavy As Boolean)
'Monke-God
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(TheText, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(TheText, ninelen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / TextLen% * i) + B3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    TextLen% = Len(part4$)
    For i = 1 To TextLen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / TextLen% * i) + B4, ((G5 - G4) / TextLen% * i) + G4, ((R5 - R4) / TextLen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    TextLen% = Len(part5$)
    For i = 1 To TextLen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B6 - B5) / TextLen% * i) + B5, ((G6 - G5) / TextLen% * i) + G5, ((R6 - R5) / TextLen% * i) + R5)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    TextLen% = Len(part6$)
    For i = 1 To TextLen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B7 - B6) / TextLen% * i) + B6, ((G7 - G6) / TextLen% * i) + G6, ((R7 - R6) / TextLen% * i) + R6)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    TextLen% = Len(part7$)
    For i = 1 To TextLen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B8 - B7) / TextLen% * i) + B7, ((G8 - G7) / TextLen% * i) + G7, ((R8 - R7) / TextLen% * i) + R7)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    TextLen% = Len(part8$)
    For i = 1 To TextLen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B9 - B8) / TextLen% * i) + B8, ((G9 - G8) / TextLen% * i) + G8, ((R9 - R8) / TextLen% * i) + R8)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part9
    TextLen% = Len(part9$)
    For i = 1 To TextLen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B10 - B9) / TextLen% * i) + B9, ((G10 - G9) / TextLen% * i) + G9, ((R10 - R9) / TextLen% * i) + R9)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function
Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, Wavy As Boolean)
'Monke-God
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    fstlen% = (Int(TextLen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, TextLen% - fstlen%)
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Wavy As Boolean)
'Monke-God
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen$ = Len(TheText)
    For i = 1 To TextLen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen$ * i) + B1, ((G2 - G1) / TextLen$ * i) + G1, ((R2 - R1) / TextLen$ * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    FadeTwoColor = Faded$
End Function
Sub FadePreview(PicB As PictureBox, ByVal FadedText As String)
FadedText$ = ReplaceString(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    Select Case t$
      Case "u"
        PicB.FontUnderline = True
      Case "/u"
        PicB.FontUnderline = False
      Case "s"
        PicB.FontStrikethru = True
      Case "/s"
        PicB.FontStrikethru = False
      Case "b"    'start bold
        PicB.FontBold = True
      Case "/b"   'stop bold
        PicB.FontBold = False
      Case "i"    'start italic
        PicB.FontItalic = True
      Case "/i"   'stop italic
        PicB.FontItalic = False
      Case "sup"  'start superscript
        TextOffY = -1
      Case "/sup" 'end superscript
        TextOffY = 0
      Case "sub"  'start subscript
        TextOffY = 1
      Case "/sub" 'end subscript
        TextOffY = 0
      Case Else
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print c$
        TextOffX = TextOffX + PicB.TextWidth(c$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub

Function GetRGB(ByVal CVal As Long)
'Monke-God
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function
Function GETVAL%(ByVal strLetter$)
'Monke-GOD
  Select Case strLetter$
    Case "0"
      GETVAL = 0
    Case "1"
      GETVAL = 1
    Case "2"
      GETVAL = 2
    Case "3"
      GETVAL = 3
    Case "4"
      GETVAL = 4
    Case "5"
      GETVAL = 5
    Case "6"
      GETVAL = 6
    Case "7"
      GETVAL = 7
    Case "8"
      GETVAL = 8
    Case "9"
      GETVAL = 9
    Case "A"
      GETVAL = 10
    Case "B"
      GETVAL = 11
    Case "C"
      GETVAL = 12
    Case "D"
      GETVAL = 13
    Case "E"
      GETVAL = 14
    Case "F"
      GETVAL = 15
  End Select
End Function
Public Sub PhishPhrases(Txt As String)
'By Sasquach
    Dim X As Long, Phrazes As Long
    Randomize X
    Phrazes = Int((Val("140") * Rnd) + 1)
    If Phrazes = "1" Then
    Txt = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
    ElseIf Phrazes = "2" Then
    Txt = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
    ElseIf Phrazes = "3" Then
    Txt = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
    ElseIf Phrazes = "4" Then
    Txt = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
    ElseIf Phrazes = "5" Then
    Txt = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
    ElseIf Phrazes = "6" Then
    Txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
    ElseIf Phrazes = "7" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    ElseIf Phrazes = "8" Then
    Txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
    ElseIf Phrazes = "9" Then
    Txt = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
    ElseIf Phrazes = "10" Then
    Txt = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
    ElseIf Phrazes = "11" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
    ElseIf Phrazes = "12" Then
    Txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
    ElseIf Phrazes = "13" Then
    Txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
    ElseIf Phrazes = "14" Then
    Txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
    ElseIf Phrazes = "15" Then
    Txt = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
    ElseIf Phrazes = "16" Then
    Txt = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
    ElseIf Phrazes = "17" Then
    Txt = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
    ElseIf Phrazes = "18" Then
    Txt = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
    ElseIf Phrazes = "19" Then
    Txt = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
    ElseIf Phrazes = "20" Then
    Txt = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
    ElseIf Phrazes = "21" Then
    Txt = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
    ElseIf Phrazes = "22" Then
    Txt = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
    ElseIf Phrazes = "23" Then
    Txt = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
    ElseIf Phrazes = "24" Then
    Txt = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
    ElseIf Phrazes = "25" Then
    Txt = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
    ElseIf Phrazes = "26" Then
    Txt = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
    ElseIf Phrazes = "27" Then
    Txt = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
    ElseIf Phrazes = "28" Then
    Txt = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
    ElseIf Phrazes = "29" Then
    Txt = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
    ElseIf Phrazes = "30" Then
    Txt = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
    ElseIf Phrazes = "31" Then
    Txt = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
    ElseIf Phrazes = "32" Then
    Txt = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
    ElseIf Phrazes = "33" Then
    Txt = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
    ElseIf Phrazes = "34" Then
    Txt = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
    ElseIf Phrazes = "35" Then
    Txt = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
    ElseIf Phrazes = "36" Then
    Txt = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
    ElseIf Phrazes = "37" Then
    Txt = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
    ElseIf Phrazes = "38" Then
    Txt = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
    ElseIf Phrazes = "39" Then
    Txt = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
    ElseIf Phrazes = "40" Then
    Txt = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
    ElseIf Phrazes = "41" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    ElseIf Phrazes = "42" Then
    Txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
    ElseIf Phrazes = "43" Then
    Txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    ElseIf Phrazes = "44" Then
    Txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
    ElseIf Phrazes = "45" Then
    Txt = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
    ElseIf Phrazes = "46" Then
    Txt = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
    ElseIf Phrazes = "47" Then
    Txt = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
    ElseIf Phrazes = "48" Then
    Txt = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
    ElseIf Phrazes = "49" Then
    Txt = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
    ElseIf Phrazes = "50" Then
    Txt = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
    ElseIf Phrazes = "51" Then
    Txt = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
    ElseIf Phrazes = "52" Then
    Txt = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
    ElseIf Phrazes = "53" Then
    Txt = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
    ElseIf Phrazes = "54" Then
    Txt = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
    ElseIf Phrazes = "55" Then
    Txt = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
    ElseIf Phrazes = "56" Then
    Txt = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
    ElseIf Phrazes = "57" Then
    Txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
    ElseIf Phrazes = "58" Then
    Txt = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
    ElseIf Phrazes = "59" Then
    Txt = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
    ElseIf Phrazes = "60" Then
    Txt = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
    ElseIf Phrazes = "61" Then
    Txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
    ElseIf Phrazes = "62" Then
    Txt = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
    ElseIf Phrazes = "63" Then
    Txt = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
    ElseIf Phrazes = "64" Then
    Txt = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
    ElseIf Phrazes = "65" Then
    Txt = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
    ElseIf Phrazes = "66" Then
    Txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
    ElseIf Phrazes = "67" Then
    Txt = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
    ElseIf Phrazes = "68" Then
    Txt = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
    ElseIf Phrazes = "69" Then
    Txt = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
    ElseIf Phrazes = "70" Then
    Txt = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
    ElseIf Phrazes = "71" Then
    Txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
    ElseIf Phrazes = "72" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
    ElseIf Phrazes = "73" Then
    Txt = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
    ElseIf Phrazes = "74" Then
    Txt = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
    ElseIf Phrazes = "75" Then
    Txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
    ElseIf Phrazes = "76" Then
    Txt = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
    ElseIf Phrazes = "77" Then
    Txt = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
    ElseIf Phrazes = "78" Then
    Txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
    ElseIf Phrazes = "79" Then
    Txt = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
    ElseIf Phrazes = "80" Then
    Txt = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
    ElseIf Phrazes = "81" Then
    Txt = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
    ElseIf Phrazes = "82" Then
    Txt = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
    ElseIf Phrazes = "83" Then
    Txt = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
    ElseIf Phrazes = "84" Then
    Txt = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
    ElseIf Phrazes = "85" Then
    Txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    
    End If
    Txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
End Sub

Public Sub MailToListNew(TheList As ListBox)
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = AOLMailFindBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem AddMails& + 1 & ")   " & MyString$
    Next AddMails&
End Sub



Public Function CheckAlive(ScreenName As String) As Boolean
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call AOLMailSend("*, " & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = hWndText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
    Else
        CheckAlive = True
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        NoWindow& = FindWindow("#32770", "America Online")
        NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
    Loop Until NoWindow& <> 0& And NoButton& <> 0
    Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function


Public Function CheckIfMaster() As Boolean
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim AOL As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, Modal As Long, mStatic As Long
    Dim mString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Call AOLKeyWord("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Parental Controls")
        pButton& = FindWindowEx(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pButton& <> 0&
    ProgramPause 0.3
    Do
        DoEvents
        Call PostMessage(pButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(pButton&, WM_LBUTTONUP, 0&, 0&)
        ProgramPause 0.8
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        mStatic& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
        mString$ = hWndText(mStatic&)
    Loop Until Modal& <> 0 And mStatic& <> 0& And mString$ <> ""
    mString$ = ReplaceString(mString$, Chr(10), "")
    mString$ = ReplaceString(mString$, Chr(13), "")
    If mString$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
End Function


Public Function CheckIMs(Person As String) As Boolean
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call AOLKeyWord("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(IM&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(IM&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = hWndText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
    Else
        CheckIMs = False
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function
Public Function ProfileGet(ScreenName As String) As String
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = hWndText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function
Public Sub SetMailPrefs()
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, mPrefs As Long, mButton As Long
    Dim gStatic As Long, mStatic As Long, fStatic As Long
    Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
    Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 3
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        mPrefs& = FindWindowEx(MDI&, 0&, "AOL Child", "Preferences")
        gStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "General")
        mStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Mail")
        fStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Font")
        maStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Marketing")
    Loop Until mPrefs& <> 0& And gStatic& <> 0& And mStatic& <> 0& And fStatic& <> 0& And maStatic& <> 0&
    mButton& = FindWindowEx(mPrefs&, 0&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call SendMessage(mButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(mButton&, WM_LBUTTONUP, 0&, 0&)
        dMod& = FindWindow("_AOL_Modal", "Mail Preferences")
        ProgramPause 0.6
    Loop Until dMod& <> 0&
    ConfirmCheck& = FindWindowEx(dMod&, 0&, "_AOL_Checkbox", vbNullString)
    CloseCheck& = FindWindowEx(dMod&, ConfirmCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, CloseCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    OKButton& = FindWindowEx(dMod&, 0&, "_AOL_icon", vbNullString)
    Call SendMessage(ConfirmCheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(CloseCheck&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(SpellCheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Call PostMessage(mPrefs&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub MailToListOld(TheList As ListBox)
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = AOLMailFindBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem AddMails& + 1 & ")   " & MyString$
    Next AddMails&
End Sub


Public Sub MailToListFlash(TheList As ListBox)
'Ðø§ - www.hider.com/dos/ - xDoSx@hotmail.com
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim Count As Long, MyString As String, AddMails As Long
    Dim sLength As Long, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    If fMail& = 0& Then Exit Sub
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MyString$ = String(255, 0)
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(fList&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(fList&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        MyString$ = ReplaceString(MyString$, Chr(0), "")
TheList.AddItem AddMails& + 1 & ")   " & MyString$
    Next AddMails&
End Sub
Public Function FindInfoWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindInfoWindow& = child&
End Function
