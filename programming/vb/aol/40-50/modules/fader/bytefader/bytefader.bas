Attribute VB_Name = "bytefader"
' sup everyone this is bytè
' i made this bas for all of you programmers that
' like to program for aol4.0
' to use this bas u need another bas with sendchat
' or aolchatsend or sendtext ( 32 bit only )
' to activate the fade type -
' SendChat "" & LBlue_Yellow("bytefader kix")
'
'                     bytè
'
' o ya this bas has 95 fader options!

Function Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Yellow_LBlue = msg
End Function
    
Function YellowRedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowRedYellow = msg
End Function

Function YellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowRed = msg
End Function
Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPurpleYellow = msg
End Function
Function YellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPurple = msg
End Function
Function YellowPink(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(78, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPink = msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowGreenYellow = msg
End Function
Function YellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowGreen = msg
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlueYellow = msg
End Function
Function YellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlue = msg
End Function
Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlackYellow = msg
End Function
Function YellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlack = msg
End Function
Function WhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    WhitePurple = msg
End Function

Function RGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function



Function RGB2HEX(r, G, b)
    Dim X&
    Dim XX&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = b
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = r
        For XX& = 1 To 2
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
        Next XX&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedYellowRed = msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedYellow = msg
End Function

Function RedPurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedPurpleRed = msg
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedPurple = msg
End Function
Function RedBlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlueRed = msg
End Function
Function RedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlue = msg
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlackRed = msg
End Function
Function RedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlack = msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleYellowPurple = msg
End Function

Function PurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleYellow = msg
End Function

Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleRedPurple = msg
End Function
Function PurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleRed = msg
End Function
Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleGreenPurple = msg
End Function
Function PurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleGreen = msg
End Function
Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBluePurple = msg
End Function

Function PurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlue = msg
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlackPurple = msg
End Function
Function PurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlack = msg
End Function

Function Purple_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Purple_LBlue = msg
End Function

Function LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Yellow = msg
End Function
Function LBlue_Green(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Green = msg
End Function
Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyYellowGrey = msg
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyYellow = msg
End Function
Function GreyRedGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyRedGrey = msg
End Function
Function GreyRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyRed = msg
End Function
Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyPurpleGrey = msg
End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyPurple = msg
End Function

Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyGreenGrey = msg
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyGreen = msg
End Function

Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlueGrey = msg
End Function
Function GreyBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlue = msg
End Function

Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlackGrey = msg
End Function

Function GreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 220 / a
        F = e * b
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlack = msg
End Function
Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenYellowGreen = msg
End Function
Function GreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenYellow = msg
End Function
Function GreenRedGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenRedGreen = msg
End Function
Function GreenRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenRed = msg
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenPurpleGreen = msg
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenPurple = msg
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlueGreen = msg
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlue = msg
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlackGreen = msg
End Function

Function GreenBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlack = msg
End Function
Function DBlue_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    DBlue_Black = msg
End Function
Public Sub CenterFormTop(frm As Form)
' this function will center your form and also keep
' it on top of the screen
' to use type - CenterFormTop Me ( in form_load )
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub
Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueYellowBlue = msg
End Function
Function BlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueYellow = msg
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueRedBlue = msg
End Function


Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BluePurpleBlue = msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueRed = msg
End Function
Function BluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BluePurple = msg
End Function
Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueGreenBlue = msg
End Function
Function BlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueGreen = msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueBlackBlue = msg
End Function


Function BlueBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueBlack = msg
End Function

Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackYellowBlack = msg
End Function
Function BlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackYellow = msg
End Function
Function BlackRedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackRedBlack = msg
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackRed = msg
End Function
Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackPurpleBlack = msg
End Function
Function BlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackPurple = msg
End Function
Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGreyBlack = msg
End Function
Function BlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        F = e * b
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGrey = msg
End Function
Function Black_LBlue_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Black_LBlue_Black = msg
End Function

Function Black_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, F, F - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Black_LBlue = msg
End Function



Function YellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPink = msg
End Function

Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    WhitePurpleWhite = msg
End Function

Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Green_LBlue = msg
End Function

Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Yellow_LBlue = msg
End Function

Function Purple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Purple_LBlue = msg
End Function

Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 450 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    DBlue_Black_DBlue = msg
End Function

Function DGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    DGreen_Black = msg
End Function



Function LBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 155, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange = msg
End Function



Function LBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange_LBlue = msg
End Function

Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 220 / a
        F = e * b
        G = RGB(0, 375 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LGreen_DGreen = msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LGreen_DGreen_LGreen = msg
End Function

Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(355, 255 - F, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_DBlue = msg
End Function

Function LBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_DBlue_LBlue = msg
End Function

Function PinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        F = e * b
        G = RGB(255 - F, 167, 510)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PinkOrange = msg
End Function

Function PinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PinkOrangePink = msg
End Function

Function PurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 200 / a
        F = e * b
        G = RGB(255, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleWhite = msg
End Function

Function PurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleWhitePurple = msg
End Function

Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        d = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Yellow_LBlue_Yellow = msg
End Function


