Attribute VB_Name = "CryoFade"
'CryoFade.bas by VvVCryoVvV
'Feel free so pass this bas around,
'Just please refrain from changing anything
'without sending mail to VvVCryoVvV@aol.com
'Created 1998
'Enjoy


'Pre-set 2 color fade combinations begin here


Function BlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlue = msg
End Function

Function BlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreen = msg
End Function

Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 220 / a
        F = E * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGrey = msg
End Function

Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurple = msg
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRed = msg
End Function

Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellow = msg
End Function

Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlack = msg
End Function

Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreen = msg
End Function

Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurple = msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRed = msg
End Function

Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellow = msg
End Function

Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlack = msg
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlue = msg
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurple = msg
End Function

Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRed = msg
End Function

Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellow = msg
End Function

Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 220 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlack = msg
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlue = msg
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreen = msg
End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurple = msg
End Function

Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRed = msg
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellow = msg
End Function

Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlack = msg
End Function

Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlue = msg
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreen = msg
End Function

Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRed = msg
End Function

Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellow = msg
End Function

Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlack = msg
End Function

Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlue = msg
End Function

Function RedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreen = msg
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurple = msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellow = msg
End Function

Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlack = msg
End Function

Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlue = msg
End Function

Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreen = msg
End Function

Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurple = msg
End Function

Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRed = msg
End Function


'Pre-set 3 Color fade combinations begin here


Function BlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlueBlack = msg
End Function

Function BlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreenBlack = msg
End Function

Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreyBlack = msg
End Function

Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurpleBlack = msg
End Function

Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRedBlack = msg
End Function

Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellowBlack = msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlackBlue = msg
End Function

Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreenBlue = msg
End Function

Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurpleBlue = msg
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRedBlue = msg
End Function

Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellowBlue = msg
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlackGreen = msg
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlueGreen = msg
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurpleGreen = msg
End Function

Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRedGreen = msg
End Function

Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellowGreen = msg
End Function

Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlackGrey = msg
End Function

Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlueGrey = msg
End Function

Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreenGrey = msg
End Function

Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurpleGrey = msg
End Function

Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRedGrey = msg
End Function

Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellowGrey = msg
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlackPurple = msg
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBluePurple = msg
End Function

Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreenPurple = msg
End Function

Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRedPurple = msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellowPurple = msg
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlackRed = msg
End Function

Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlueRed = msg
End Function

Function RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreenRed = msg
End Function

Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurpleRed = msg
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellowRed = msg
End Function

Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlackYellow = msg
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlueYellow = msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreenYellow = msg
End Function

Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurpleYellow = msg
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRedYellow = msg
End Function


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function


'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub


'Variable color fade functions begin here


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
    c = 0
    o = 0
    o2 = 0
    Q = 1
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
        
        If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: msg = msg & "<FONT COLOR=#" + C1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        
        If Wavy = True Then
            If o2 = 1 Then msg = msg + "<SUB>"
            If o2 = 3 Then msg = msg + "<SUP>"
            msg = msg + Mid$(Text, X, 1)
            If o2 = 1 Then msg = msg + "</SUB>"
            If o2 = 3 Then msg = msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
                If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
                If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf Wavy = False Then
            msg = msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next X
    C1 = C1BAK
    C2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    TwoColors = msg
End Function

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)

'This code is still buggy, use at your own risk

    D = Len(Text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    c = D \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c)
    GoTo TheEnd
Odds:
    c = D \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    msg = FadeA + FadeB
    ThreeColors = msg
End Function

Function RGB2HEX(R, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
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

Function TrimSpaces(Text)
    If InStr(Text, " ") = 0 Then
    TrimSpaces = Text
    Exit Function
    End If
    For TrimSpace = 1 To Len(Text)
    thechar$ = Mid(Text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function

