Attribute VB_Name = "VK_Fade"

'this version  9-1-98

'coded by KRhyME and SkaFia
'|¯|    |¯|\¯\ |¯|    |_|/¯//¯/|¯|\¯\ /¯/|¯|
'| |/¯/ | |/ / | |\¯\   / / | ||_|| ||  |/_/_
'|_|\_\ |_|\_\ |_| |_| /_/  |_|   |_| \______|
'HEAD of the [Voltron Kru]
'Voltron Kru '98
'www.voltronkru.com
'voltronkru@juno.com

'This Bas file requires Voltron.bas to work
'Voltron.Bas is the core bas file for the
'Voltron Kru. You can get all our bas files at
'www.voltronkru.com

'many ideas for this bas came from other Voltron Kru
'members. I would Like to thank SkaFia for all the
'things he did for the series of Bas files.
'Please do not steal our codes without giving us
'credit. I would like to say thank you to KnK for
'making so many files avaible to the public, The makers
'of DiVe32.bas (the first bas i used), Toast, Magus,
'and all the other great programmers out there who
'have infuinced us

'Please join our VB mailing list
'www.voltronkru.com

Function BlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackGreen = Msg
End Function

Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 220 / a
        F = E * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
End Function

Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackPurple = Msg
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackRed = Msg
End Function

Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackYellow = Msg
End Function

Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueBlack = Msg
End Function

Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueGreen = Msg
End Function

Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BluePurple = Msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueRed = Msg
End Function

Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueYellow = Msg
End Function

Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenBlack = Msg
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenBlue = Msg
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenPurple = Msg
End Function

Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenRed = Msg
End Function

Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenYellow = Msg
End Function

Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 220 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyBlack = Msg
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyBlue = Msg
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyGreen = Msg

End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyPurple = Msg
End Function

Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyRed = Msg
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyYellow = Msg
End Function

Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleBlack = Msg
End Function

Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
       Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleBlue = Msg
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleGreen = Msg
End Function

Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleRed = Msg
End Function

Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleYellow = Msg
End Function

Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
       F = E * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlack = Msg
End Function

Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlue = Msg
End Function

Function RedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedGreen = Msg
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
RedPurple = (Msg)
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedYellow = Msg
End Function

Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowBlack = Msg
End Function

Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowBlue = Msg
End Function

Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowGreen = Msg
End Function

Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowPurple = Msg
End Function

Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowRed = Msg
End Function


Function BlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackBlueBlack = Msg
End Function

Function BlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackGreenBlack = Msg
End Function

Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackGreyBlack = Msg
End Function

Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackPurpleBlack = Msg
End Function

Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackRedBlack = Msg
End Function

Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackYellowBlack = Msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueBlackBlue = Msg
End Function

Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueGreenBlue = Msg
End Function

Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BluePurpleBlue = Msg
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueRedBlue = Msg
End Function

Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueYellowBlue = Msg
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   GreenBlackGreen = Msg
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenBlueGreen = Msg
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenPurpleGreen = Msg
End Function

Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenRedGreen = Msg
End Function

Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenYellowGreen = Msg
End Function

Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyBlackGrey = Msg
End Function

Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyBlueGrey = Msg
End Function

Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyGreenGrey = Msg
End Function

Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   GreyPurpleGrey = Msg
End Function

Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyRedGrey = Msg
End Function

Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyYellowGrey = Msg
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleBlackPurple = Msg
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleBluePurple = Msg
End Function

Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleGreenPurple = Msg
End Function

Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleRedPurple = Msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
   For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleYellowPurple = Msg
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlackRed = Msg
End Function

Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlueRed = Msg
End Function

Function RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedGreenRed = Msg
End Function

Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedPurpleRed = Msg
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedYellowRed = Msg
End Function

Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowBlackYellow = Msg
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowBlueYellow = Msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowGreenYellow = Msg
End Function

Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowPurpleYellow = Msg
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowRedYellow = Msg
End Function

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

Function FadeBlack(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & t$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & L$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & B$ & "<FONT COLOR=#888888>" & C$ & "<FONT COLOR=#999999>" & d$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & m$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next w
FadeBlack = pc$

End Function

Function FadeGreen(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & t$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & L$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & C$ & "<FONT COLOR=#00FF00>" & d$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & m$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next w
FadeGreen = pc$
End Function

Function FadeYellow(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & t$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & L$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & C$ & "<FONT COLOR=#111100>" & d$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & m$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next w
FadeYellow = pc$
End Function

Function FadeCustom(thetext As String, Hx1 As Integer, Hx2 As Integer, Hx3 As Integer, Hx4 As Integer, Hx5 As Integer, Hx6 As Integer, Hx7 As Integer, Hx8 As Integer, Hx9 As Integer, Hx10 As Integer)
'Dont worry this is 18 hexes that can
'Be entered but I made it 10 CuZ
'it goes: 1,2,3,4,5,6,7,8,9,10,9,8,7,6,5,4,3,2
'this is so when it fades it will loop
'I entered this so you wouldnt have to delete and
'myne and edit your own or figure out how to
'a new sub
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#" & Hex1 & ">" & ab$ & "<FONT COLOR=#" & Hx2 & ">" & u$ & "<FONT COLOR=#" & Hx3 & ">" & S$ & "<FONT COLOR=#" & Hx4 & ">" & t$ & "<FONT COLOR=#" & Hx5 & ">" & Y$ & "<FONT COLOR=#" & Hx6 & ">" & L$ & "<FONT COLOR=#" & Hx7 & ">" & F$ & "<FONT COLOR=#" & Hx8 & ">" & B$ & "<FONT COLOR=#" & Hx9 & ">" & C$ & "<FONT COLOR=#" & Hx10 & ">" & d$ & "<FONT COLOR=#" & Hx9 & ">" & H$ & "<FONT COLOR=#" & Hx8 & ">" & j$ & "<FONT COLOR=#" & Hx7 & ">" & k$ & "<FONT COLOR=#" & Hx6 & ">" & m$ & "<FONT COLOR=#" & Hx5 & ">" & n$ & "<FONT COLOR=#" & Hx4 & ">" & q$ & "<FONT COLOR=#" & Hx3 & ">" & V$ & "<FONT COLOR=#" & Hx2 & ">" & Z$
Next w
FadeCustom = pc$

End Function

' The following are the color HTML codes
' for these colors --------------
'
' Red = FE0000
' Blue=0000FE
' Green=00FE00
' dkBlue=000066
' orange=FE7C00
' White=FEFEFE
' purple=C200C2
' yellow=FEFE00
' DkRed=660000
'
Function AddCustomColorToText(text As String, Blend As Boolean, Wavey As Boolean, Lagger As Boolean, Bold As Boolean, Italics As Boolean, Strikeout As Boolean, UnderLine As Boolean, CC1 As Long, CC2 As Long) As String
If text = "" Then Exit Function
wavestep = 1
MaxColor = &HFE
MinColor = &H0
txtsize = Len(text)
' set colors...
Dim RED, GREEN, BLUE, ERED, EGREEN, EBLUE
Dim REDBL, GREENBL, BLUEBL
Dim SRED As String, SGREEN As String, SBLUE As String
Dim CustomC1Str As String, CustomC2Str As String
CustomC1Str = Hex$(CC1)
CustomC2Str = Hex$(CC2)
If Len(CustomC1Str) < 6 Then
    If Len(CustomC1Str) = 5 Then CustomC1Str = "0" & CustomC1Str
    If Len(CustomC1Str) = 4 Then CustomC1Str = "00" & CustomC1Str
    If Len(CustomC1Str) = 3 Then CustomC1Str = "000" & CustomC1Str
    If Len(CustomC1Str) = 2 Then CustomC1Str = "0000" & CustomC1Str
    If Len(CustomC1Str) = 1 Then CustomC1Str = "00000" & CustomC1Str
End If

If Len(CustomC2Str) < 6 Then
    If Len(CustomC2Str) = 5 Then CustomC2Str = "0" & CustomC2Str
    If Len(CustomC2Str) = 4 Then CustomC2Str = "00" & CustomC2Str
    If Len(CustomC2Str) = 3 Then CustomC2Str = "000" & CustomC2Str
    If Len(CustomC2Str) = 2 Then CustomC2Str = "0000" & CustomC2Str
    If Len(CustomC2Str) = 1 Then CustomC2Str = "00000" & CustomC2Str
End If

RED = Right(CustomC1Str, 2)
GREEN = Mid(CustomC1Str, 3, 2)
BLUE = Left(CustomC1Str, 2)
ERED = Right(CustomC2Str, 2)
EGREEN = Mid(CustomC2Str, 3, 2)
EBLUE = Left(CustomC2Str, 2)

RED = Val("&H" & RED)
GREEN = Val("&H" & GREEN)
BLUE = Val("&H" & BLUE)
ERED = Val("&H" & ERED)
EGREEN = Val("&H" & EGREEN)
EBLUE = Val("&H" & EBLUE)


' If Blend is true then takes you to the
' blending function

FinalTxt$ = ""

If Blend = True Then GoTo BlendTxt

' Blends colors from startcolor to endcolor

REDBL = -(Int(((RED - ERED)) / txtsize))
GREENBL = -(Int(((GREEN - EGREEN)) / txtsize))
BLUEBL = -(Int(((BLUE - EBLUE)) / txtsize))

For X = 1 To txtsize

SRED = MakeHexString(RED)
SGREEN = MakeHexString(GREEN)
SBLUE = MakeHexString(BLUE)

FinalTxt$ = FinalTxt$ + "<FONT COLOR=""#" & SRED & SGREEN & SBLUE & """>" & Mid(text, X, 1)
If wavestep = 5 Then wavestep = 1
If Wavey = True And wavestep = 1 Then FinalTxt$ = FinalTxt$ + "<SUP>"
If Wavey = True And wavestep = 2 Then FinalTxt$ = FinalTxt$ + "</SUP>"
If Wavey = True And wavestep = 3 Then FinalTxt$ = FinalTxt$ + "<SUB>"
If Wavey = True And wavestep = 4 Then FinalTxt$ = FinalTxt$ + "</SUB>"
wavestep = wavestep + 1
If Lagger = True Then
FinalTxt$ = FinalTxt$ + "<HTML><HTML><HTML><HTML>"
If Bold = True Then FinalTxt$ = FinalTxt$ + "<B>"
If Italics = True Then FinalTxt$ = FinalTxt$ + "<I>"
If Strikeout = True Then FinalTxt$ = FinalTxt$ + "<S>"
If UnderLine = True Then FinalTxt$ = FinalTxt$ + "<U>"
End If
RED = RED + REDBL
GREEN = GREEN + GREENBL
BLUE = BLUE + BLUEBL
If RED > 254 Then RED = 254
If RED < 0 Then RED = 0
If GREEN > 254 Then GREEN = 254
If GREEN < 0 Then GREEN = 0
If BLUE > 254 Then BLUE = 254
If BLUE < 0 Then BLUE = 0

Next X

AddCustomColorToText = FinalTxt$

Exit Function


BlendTxt:

If (Len(text) / 2) <> (Abs(Len(text) / 2)) Then txtsize = txtsize - 1


REDBL = -(Int(((RED - ERED)) / txtsize))
GREENBL = -(Int(((GREEN - EGREEN)) / txtsize))
BLUEBL = -(Int(((BLUE - EBLUE)) / txtsize))

REDBL = (Int(REDBL * 2))
GREENBL = (Int(GREENBL * 2))
BLUEBL = (Int(BLUEBL * 2))

For X = 1 To Int(txtsize / 2)

SRED = MakeHexString(RED)
SGREEN = MakeHexString(GREEN)
SBLUE = MakeHexString(BLUE)

FinalTxt$ = FinalTxt$ + "<FONT COLOR=""#" & SRED & SGREEN & SBLUE & """>" & Mid(text, X, 1)
If Lagger = True Then
FinalTxt$ = FinalTxt$ + "<HTML><HTML><HTML><HTML>"
If Bold = True Then FinalTxt$ = FinalTxt$ + "<B>"
If Italics = True Then FinalTxt$ = FinalTxt$ + "<I>"
If Strikeout = True Then FinalTxt$ = FinalTxt$ + "<S>"
If UnderLine = True Then FinalTxt$ = FinalTxt$ + "<U>"
End If
If wavestep = 5 Then wavestep = 1
If Wavey = True And wavestep = 1 Then FinalTxt$ = FinalTxt$ + "<SUP>"
If Wavey = True And wavestep = 2 Then FinalTxt$ = FinalTxt$ + "</SUP>"
If Wavey = True And wavestep = 3 Then FinalTxt$ = FinalTxt$ + "<SUB>"
If Wavey = True And wavestep = 4 Then FinalTxt$ = FinalTxt$ + "</SUB>"
wavestep = wavestep + 1

RED = RED + REDBL
GREEN = GREEN + GREENBL
BLUE = BLUE + BLUEBL
If RED > 254 Then RED = 254
If RED < 0 Then RED = 0
If GREEN > 254 Then GREEN = 254
If GREEN < 0 Then GREEN = 0
If BLUE > 254 Then BLUE = 254
If BLUE < 0 Then BLUE = 0

Next X

For X = (Int(txtsize / 2)) + 1 To txtsize
If Lagger = True Then
FinalTxt$ = FinalTxt$ + "<HTML><HTML><HTML><HTML>"
If Bold = True Then FinalTxt$ = FinalTxt$ + "<B>"
If Italics = True Then FinalTxt$ = FinalTxt$ + "<I>"
If Strikeout = True Then FinalTxt$ = FinalTxt$ + "<S>"
If UnderLine = True Then FinalTxt$ = FinalTxt$ + "<U>"
End If
SRED = MakeHexString(RED)
SGREEN = MakeHexString(GREEN)
SBLUE = MakeHexString(BLUE)
If wavestep = 5 Then wavestep = 1
If Wavey = True And wavestep = 1 Then FinalTxt$ = FinalTxt$ + "<SUP>"
If Wavey = True And wavestep = 2 Then FinalTxt$ = FinalTxt$ + "</SUP>"
If Wavey = True And wavestep = 3 Then FinalTxt$ = FinalTxt$ + "<SUB>"
If Wavey = True And wavestep = 4 Then FinalTxt$ = FinalTxt$ + "</SUB>"
wavestep = wavestep + 1

FinalTxt$ = FinalTxt$ + "<FONT COLOR=""#" & SRED & SGREEN & SBLUE & """>" & Mid(text, X, 1)

RED = RED - REDBL
GREEN = GREEN - GREENBL
BLUE = BLUE - BLUEBL
If RED > 254 Then RED = 254
If RED < 0 Then RED = 0
If GREEN > 254 Then GREEN = 254
If GREEN < 0 Then GREEN = 0
If BLUE > 254 Then BLUE = 254
If BLUE < 0 Then BLUE = 0

Next X

AddCustomColorToText = FinalTxt$

End Function

Function MakeHexString(Number As Variant) As String

' Takes a number and Makes a Hexadecimal
' 2 letter string out of it

HxNumber$ = Hex$(Number)

Select Case (HxNumber$)

    Case "0"
    HxNumber$ = "00"
    
    Case "1"
    HxNumber$ = "01"
    
    Case "2"
    HxNumber$ = "02"
    
    Case "3"
    HxNumber$ = "03"
    
    Case "4"
    HxNumber$ = "04"
    
    Case "5"
    HxNumber$ = "05"
    
    Case "6"
    HxNumber$ = "06"
    
    Case "7"
    HxNumber$ = "07"
    
    Case "8"
    HxNumber$ = "08"
    
    Case "9"
    HxNumber$ = "09"
    
    Case "A"
    HxNumber$ = "0A"
    
    Case "B"
    HxNumber$ = "0B"
    
    Case "C"
    HxNumber$ = "0C"
    
    Case "D"
    HxNumber$ = "0D"
    
    Case "E"
    HxNumber$ = "0E"
    
    Case "F"
    HxNumber$ = "0F"
    
'    Case Else
'    HxNumber$ = HexNumber$
    
End Select

MakeHexString = HxNumber$

End Function

'Pre-set 2 color fade combinations begin here


Function BlackBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackBlue = Msg
End Function

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

Function WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    ua$ = Mid$(G$, w + 1, 1)
    Sa$ = Mid$(G$, w + 2, 1)
    ta$ = Mid$(G$, w + 3, 1)
    pa$ = pa$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & ua$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & Sa$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & ta$
Next w
WavyChatBlueBlack = pa$
End Function

Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavYChaTRB = p$
End Function

Function WavYChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & t$
Next w
WavYChaTRG = p$
End Function

Function WavyFader(txt)
'wavs, and fades.
a = Len(txt)
For w = 1 To a Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    B$ = Mid$(txt, w + 7, 1)
    C$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    H$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<b><FONT COLOR=#000019><sup>" & ab$ & "<FONT COLOR=#000026></sup>" & u$ & "<FONT COLOR=#00003F><sub>" & S$ & "<FONT COLOR=#000058></sub>" & t$ & "<FONT COLOR=#000072><sup>" & Y$ & "<FONT COLOR=#00008B></sup>" & L$ & "<FONT COLOR=#0000A5><sub>" & F$ & "<FONT COLOR=#0000BE></sub>" & B$ & "<FONT COLOR=#0000D7><sup>" & C$ & "<FONT COLOR=#0000F1></sup>" & d$ & "<FONT COLOR=#0000D7><sub>" & H$ & "<FONT COLOR=#0000BE></sub>" & j$ & "<FONT COLOR=#0000A5><sup>" & k$ & "<FONT COLOR=#00008B></sup>" & m$ & "<FONT COLOR=#000072><sub>" & n$ & "<FONT COLOR=#000058></sub>" & q$ & "<FONT COLOR=#00003F><sup>" & V$ & "<FONT COLOR=#000026></sup>" & Z$
Next w
GeM3 = p$
WavyFader = p$
End Function

Function WavyFader2(txt)
'wavs, and fades.
a = Len(txt)
For w = 1 To a Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    B$ = Mid$(txt, w + 7, 1)
    C$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    H$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = "<FONT FACE=""Arial Black"">" & ab$ & "<FONT FACE=""Bell MT"">" & u$ & "<FONT FACE=""Cataneo BT"">" & S$ & "<FONT FACE=""Copperplate Gothic Bold"">" & t$ & "<FONT FACE=""Desdemona"">" & Y$ & "<FONT FACE=""Engravers MT"">" & L$ & "<FONT FACE=""Eras Bold ITC"">" & F$ & "<FONT FACE=""Placard Condensed"">" & B$ & "<FONT FACE=""Tw Cen MT"">" & C$ & "<FONT FACE=""Viner Hand ITC"">" & d$ & "<FONT FACE=""Stencil"">" & H$ & "<FONT FACE=""Script MT Bold"">" & j$ & "<FONT FACE=""Placard Condensed"">" & k$ & "<FONT FACE=""Perpetuna Titling MT"">" & m$ & "<FONT COLOR=#000072><sub>" & n$ & "<FONT COLOR=#000058></sub>" & q$ & "<FONT COLOR=#00003F><sup>" & V$ & "<FONT COLOR=#000026></sup>" & Z$
Next w
GeM3 = p$
WavyFader2 = p$
End Function

Function WavyFader3(txt)
a = Len(txt)
For w = 1 To a Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    B$ = Mid$(txt, w + 7, 1)
    C$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    H$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<b><FONT COLOR=#000019><sup>" & ab$ & "<FONT COLOR=#000026></sup>" & u$ & "<FONT COLOR=#00003F><sub>" & S$ & "<FONT COLOR=#000058></sub>" & t$ & "<FONT COLOR=#000072><sup>" & Y$ & "<FONT COLOR=#00008B></sup>" & L$ & "<FONT COLOR=#0000A5><sub>" & F$ & "<FONT COLOR=#0000BE></sub>" & B$ & "<FONT COLOR=#0000D7><sup>" & C$ & "<FONT COLOR=#0000F1></sup>" & d$ & "<FONT COLOR=#0000D7><sub>" & H$ & "<FONT COLOR=#0000BE></sub>" & j$ & "<FONT COLOR=#0000A5><sup>" & k$ & "<FONT COLOR=#00008B></sup>" & m$ & "<FONT COLOR=#000072><sub>" & n$ & "<FONT COLOR=#000058></sub>" & q$ & "<FONT COLOR=#00003F><sup>" & V$ & "<FONT COLOR=#000026></sup>" & Z$
Next w
WavyFader3 = p$
End Function

Function Fade(YourMessage, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)

   C1BAK = c1
   C2BAK = c2
   C3BAK = c3
   C4BAK = c4
   C = 0
   O = 0
   o2 = 0
   q = 1
   Q2 = 1
   For X = 1 To Len(YourMessage)
            BVAL1 = Red2 - Red1
            BVAL2 = Green2 - Green1
            BVAL3 = Blue2 - Blue1
            VAL1 = (BVAL1 / Len(YourMessage) * X) + Red1
            VAL2 = (BVAL2 / Len(YourMessage) * X) + Green1
            VAL3 = (BVAL3 / Len(YourMessage) * X) + Blue1
            c1 = RGB2HEX(VAL1, VAL2, VAL3)
            c2 = RGB2HEX(VAL1, VAL2, VAL3)
            c3 = RGB2HEX(VAL1, VAL2, VAL3)
            c4 = RGB2HEX(VAL1, VAL2, VAL3)
            If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then C = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If C <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
            If WavY = True Then
                If o2 = 1 Then Msg = Msg + "<sub>"
                If o2 = 3 Then Msg = Msg + "<sup>"
                Msg = Msg + Mid$(YourMessage, X, 1)
                If o2 = 1 Then Msg = Msg + "</sub>"
                If o2 = 3 Then Msg = Msg + "</sup>"
                If Q2 = 2 Then
                    q = 1
                    Q2 = 1

                    If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                    If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                    If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                    If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
                End If
            ElseIf WavY = False Then
                Msg = Msg + Mid$(YourMessage, X, 1)
                If Q2 = 2 Then
                    q = 1
                    Q2 = 1

                    If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                    If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                    If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                    If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
                End If
        
            End If
nc: Next X
c1 = C1BAK
c2 = C2BAK
c3 = C3BAK
c4 = C4BAK
Fade = Msg
End Function

'Pre-set 2 color fade combinations begin here
Function BoldFadeBlack(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & t$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & L$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & B$ & "<FONT COLOR=#888888>" & C$ & "<FONT COLOR=#999999>" & d$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & m$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next w
BoldFadeBlack = pc$


'Code for the room will be
'Call Fadeblack(Text1.text)


'to make any of the subs werk in ims
'You will need 2 text boxes and a button
'Do the change below and copy that to your send button
   ' a = Len(Text2.text)
    'For B = 1 To a
        'c = Left(Text2.text, B)
        'D = Right(c, 1)
        'e = 255 / a
        'F = e * B
        'G = RGB(F, 0, 0)
        'H = RGBtoHEX(G)
    ' Dim msg
    ' msg=msg & "<B><Font Color=#" & H & ">" & D
    'Next B
   ' Call IMKeyword(Text1.text, msg)
'u can do it for mail too but
'that is harder and I will leave that to u
'to figure out
End Function

Function BoldFadeGreen(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & t$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & L$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & C$ & "<FONT COLOR=#00FF00>" & d$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & m$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next w
BoldFadeGreen = pc$
End Function

Function BoldFadeRed(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & u$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & t$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & L$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & B$ & "<FONT COLOR=#220000>" & C$ & "<FONT COLOR=#110000>" & d$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & k$ & "<FONT COLOR=#550000>" & m$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next w
BoldFadeRed = pc$
End Function

Function BoldFadeBlue(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & u$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & t$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & L$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & B$ & "<FONT COLOR=#0000D7>" & C$ & "<FONT COLOR=#0000F1>" & d$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & k$ & "<FONT COLOR=#00008B>" & m$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next w
BoldFadeBlue = pc$
End Function

Function BoldFadeYellow(thetext As String)
a = Len(thetext)
For w = 1 To a Step 18
    ab$ = Mid$(thetext, w, 1)
    u$ = Mid$(thetext, w + 1, 1)
    S$ = Mid$(thetext, w + 2, 1)
    t$ = Mid$(thetext, w + 3, 1)
    Y$ = Mid$(thetext, w + 4, 1)
    L$ = Mid$(thetext, w + 5, 1)
    F$ = Mid$(thetext, w + 6, 1)
    B$ = Mid$(thetext, w + 7, 1)
    C$ = Mid$(thetext, w + 8, 1)
    d$ = Mid$(thetext, w + 9, 1)
    H$ = Mid$(thetext, w + 10, 1)
    j$ = Mid$(thetext, w + 11, 1)
    k$ = Mid$(thetext, w + 12, 1)
    m$ = Mid$(thetext, w + 13, 1)
    n$ = Mid$(thetext, w + 14, 1)
    q$ = Mid$(thetext, w + 15, 1)
    V$ = Mid$(thetext, w + 16, 1)
    Z$ = Mid$(thetext, w + 17, 1)
    pc$ = pc$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & t$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & L$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & C$ & "<FONT COLOR=#111100>" & d$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & m$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next w
BoldFadeYellow = pc$
End Function


Function BoldBlackBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)

End Function

Function BoldBlackGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
  End Function

Function BoldBlackGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 220 / a
        F = E * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlackPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlackRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldBlackYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldBlueBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlueGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldBluePurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlueRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldBlueYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreenBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldGreenBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldGreenPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldGreenRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldGreenYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreyBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 220 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldRedBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldRedBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldRedGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldRedPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldRedYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldYellowBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellowBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function


'Pre-set 3 Color fade combinations begin here


Function BoldBlackBlueBlack2(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><U><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function
Function BoldBlackBlueBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function
Function BoldBlackGreenBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlackGreyBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function Bolditalic_BlackPurpleBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlackRedBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlackYellowBlack(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlueBlackBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlueGreenBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function Bolditalic_BluePurpleBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldBlueRedBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlueYellowBlue(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreenBlackGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreenBlueGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldGreenPurpleGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreenRedGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function


Function BoldGreenYellowGreen(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreyBlackGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyBlueGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyGreenGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyPurpleGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyRedGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyYellowGrey(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldPurpleBlackPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBluePurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreenPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleRedPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleYellowPurple(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function RedBlackRed2(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><U><Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function
Function BoldRedBlackRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function
Function BoldRedBlueRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldRedGreenRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldRedPurpleRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldRedYellowRed(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellowBlackYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowBlueYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellowGreenYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowPurpleYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowRedYellow(text As String)
    a = Len(text)
    For B = 1 To a
        C = Left(text, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function TwoColors(text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    C = 0
    O = 0
    o2 = 0
    q = 1
    Q2 = 1
    For X = 1 To Len(text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(text) * X) + Red1
        VAL2 = (BVAL2 / Len(text) * X) + Green1
        VAL3 = (BVAL3 / Len(text) * X) + Blue1
        
        c1 = RGB2HEX(VAL1, VAL2, VAL3)
        c2 = RGB2HEX(VAL1, VAL2, VAL3)
        c3 = RGB2HEX(VAL1, VAL2, VAL3)
        c4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then C = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If C <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
        ElseIf WavY = False Then
            Msg = Msg + Mid$(text, X, 1)
            If Q2 = 2 Then
            q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        End If
nc:     Next X
    c1 = C1BAK
    c2 = C2BAK
    c3 = C3BAK
    c4 = C4BAK
    BoldSendChat (Msg)
End Function

Function ThreeColors(text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)

'This code is still buggy, use at your own risk

    d = Len(text)
        If d = 0 Then GoTo TheEnd
        If d = 1 Then Fade1 = text
    For X = 2 To 500 Step 2
        If d = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If d = X Then GoTo Odds
    Next X
Evens:
    C = d \ 2
    Fade1 = Left(text, C)
    Fade2 = Right(text, C)
    GoTo TheEnd
Odds:
    C = d \ 2
    Fade1 = Left(text, C)
    Fade2 = Right(text, C + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If WavY = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If WavY = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If WavY = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If WavY = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    Msg = FadeA + FadeB
  BoldSendChat (Msg)
End Function

Function RGB2HEX(r, G, B)
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

Function Bold_italic_colorR_Backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function

Function Bold_WavColors(Text1 As String)
G$ = Text1
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & t$
Next w
Bold_WavColors = p$
End Function

Function BoldBlack_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function



Function BoldYellowPinkYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldWhitePurpleWhite(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function Blue_Green_Blue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    Blue_Green_Blue = Msg
End Function

Function Blue_Yellow_Blue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    Blue_Yellow_Blue = Msg
End Function

Function BoldPurple_LBlue_Purple()
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldDBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 450 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldDGreen_Black(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function



Function BoldLBlue_Orange(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function



Function BoldLBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldLGreen_DGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 220 / a
        F = E * B
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldLGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldLBlue_DBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 255 / a
        F = E * B
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 BoldSendChat (Msg)
End Function

Function BoldLBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
BoldSendChat (Msg)
End Function

Function BoldPinkOrange(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 200 / a
        F = E * B
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPinkOrangePink(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhite(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 200 / a
        F = E * B
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhitePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        C = Left(Text1, B)
        d = Right(C, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
   BoldSendChat (Msg)
End Function

Function WaveTalker(strin2, F As ComboBox, c1 As ComboBox, c2 As ComboBox, c3 As ComboBox, c4 As ComboBox)
tixt = F
Color1 = c1
Color2 = c2
Color3 = c3
Color4 = c4
If Color1 = "Navy" Then Color1 = "000080"
If Color1 = "Maroon" Then Color1 = "800000"
If Color1 = "Lime" Then Color1 = "00FF00"
If Color1 = "Teal" Then Color1 = "008080"
If Color1 = "Red" Then Color1 = "F0000"
If Color1 = "Blue" Then Color1 = "0000FF"
If Color1 = "Siler" Then Color1 = "C0C0C0"
If Color1 = "Yellow" Then Color1 = "FFFF00"
If Color1 = "Aqua" Then Color1 = "00FFFF"
If Color1 = "Purple" Then Color1 = "800080"
If Color1 = "Black" Then Color1 = "000000"

If Color2 = "Navy" Then Color2 = "000080"
If Color2 = "Maroon" Then Color2 = "800000"
If Color2 = "Lime" Then Color2 = "00FF00"
If Color2 = "Teal" Then Color2 = "008080"
If Color2 = "Red" Then Color2 = "F0000"
If Color2 = "Blue" Then Color2 = "0000FF"
If Color2 = "Siler" Then Color2 = "C0C0C0"
If Color2 = "Yellow" Then Color2 = "FFFF00"
If Color2 = "Aqua" Then Color2 = "00FFFF"
If Color2 = "Purple" Then Color2 = "800080"
If Color1 = "Black" Then Color2 = "000000"

If Color3 = "Navy" Then Color3 = "000080"
If Color3 = "Maroon" Then Color3 = "800000"
If Color3 = "Lime" Then Color3 = "00FF00"
If Color3 = "Teal" Then Color3 = "008080"
If Color3 = "Red" Then Color3 = "F0000"
If Color3 = "Blue" Then Color3 = "0000FF"
If Color3 = "Siler" Then Color3 = "C0C0C0"
If Color3 = "Yellow" Then Color3 = "FFFF00"
If Color3 = "Aqua" Then Color3 = "00FFFF"
If Color3 = "Purple" Then Color3 = "800080"
If Color1 = "Black" Then Color3 = "000000"

If Color4 = "Navy" Then Color4 = "000080"
If Color4 = "Maroon" Then Color4 = "800000"
If Color4 = "Lime" Then Color4 = "00FF00"
If Color4 = "Teal" Then Color4 = "008080"
If Color4 = "Red" Then Color4 = "F0000"
If Color4 = "Blue" Then Color4 = "0000FF"
If Color4 = "Siler" Then Color4 = "C0C0C0"
If Color4 = "Yellow" Then Color4 = "FFFF00"
If Color4 = "Aqua" Then Color4 = "00FFFF"
If Color4 = "Purple" Then Color4 = "800080"
If Color1 = "Black" Then Color4 = "000000"

Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Loop
WavyTalker = newsent2$
End Function

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
AOL4_SendChat ("<u>" & UnderLineChat & "</u>")
End Sub
Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
AOL4_SendChat ("<i>" & ItalicChat & "</i>")
End Sub
Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
AOL4_SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub BoldWavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
BoldSendChat (p$)
End Sub
Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & t$
Next w
BoldSendChat (p$)
End Function
Sub BoldWavyColorbluegree(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & t$
Next w
BoldSendChat (p$)
End Sub
Sub BoldWavyColorredandblack(thetext)

G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & t$
Next w
BoldSendChat (p$)
End Sub
Sub BoldWavyColorredandblue(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    ra$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    t$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & ra$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & t$
Next w
BoldSendChat (p$)
End Sub


Sub ColorIM(Person)
'This sends someone blank IMs with a different colors
'in each one. It sends 5 IMs
Call AOL4_InstantMessage(Person, "<body bgcolor=#000000>")
Call AOL4_InstantMessage(Person, "<body bgcolor=#0000FF>")
Call AOL4_InstantMessage(Person, "<body bgcolor=#FF0000>")
Call AOL4_InstantMessage(Person, "<body bgcolor=#00FF00>")
Call AOL4_InstantMessage(Person, "<body bgcolor=#C0C0C0>")
End Sub




Sub FormFireFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub FormPlatinumFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub FormIceFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub

Sub FormFadeBG(TheForm As Form)
'another form fade

Dim hBrush%
    Dim formHeight%, RED%, StepInterval%, X%, retval%, OldMode%
    Dim FillArea As rect
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3  'Pixel
    formHeight = TheForm.ScaleHeight
' Divide the form into 63 regions
    StepInterval = formHeight \ 63
    RED = 255
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For X = 1 To 63
        hBrush% = CreateSolidBrush(RGB(0, 0, RED))
        retval% = FillRect(TheForm.hdc, FillArea, hBrush)
     
        RED = RED - 4
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
' Fill the remainder of the form with black
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush% = CreateSolidBrush(RGB(0, 0, 0))
    retval% = FillRect(TheForm.hdc, FillArea, hBrush)

    TheForm.ScaleMode = OldMode
End Sub

