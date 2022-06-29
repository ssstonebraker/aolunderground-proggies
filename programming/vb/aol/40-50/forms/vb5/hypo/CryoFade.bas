Attribute VB_Name = "CryoFade"
'CryoFade.bas by VvVCryoVvV
'Feel free so pass this bas around,
'Just please refrain from changing anything
'without sending mail to VvVCryoVvV@aol.com
'Created 1998
'Enjoy


'Pre-set 2 color fade combinations begin here


Function BlackBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlue = Msg
End Function

Function BlackGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreen = Msg
End Function

Function BlackGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 220 / A
        F = E * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGrey = Msg
End Function

Function BlackPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurple = Msg
End Function

Function BlackRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRed = Msg
End Function

Function BlackYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellow = Msg
End Function

Function BlueBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlack = Msg
End Function

Function BlueGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreen = Msg
End Function

Function BluePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurple = Msg
End Function

Function BlueRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRed = Msg
End Function

Function BlueYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellow = Msg
End Function

Function GreenBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlack = Msg
End Function

Function GreenBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlue = Msg
End Function

Function GreenPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurple = Msg
End Function

Function GreenRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRed = Msg
End Function

Function GreenYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellow = Msg
End Function

Function GreyBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 220 / A
        F = E * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlack = Msg
End Function

Function GreyBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlue = Msg
End Function

Function GreyGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreen = Msg
End Function

Function GreyPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurple = Msg
End Function

Function GreyRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRed = Msg
End Function

Function GreyYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellow = Msg
End Function

Function PurpleBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlack = Msg
End Function

Function PurpleBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlue = Msg
End Function

Function PurpleGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreen = Msg
End Function

Function PurpleRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRed = Msg
End Function

Function PurpleYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellow = Msg
End Function

Function RedBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlack = Msg
End Function

Function RedBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlue = Msg
End Function

Function RedGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreen = Msg
End Function

Function RedPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurple = Msg
End Function

Function RedYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellow = Msg
End Function

Function YellowBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlack = Msg
End Function

Function YellowBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlue = Msg
End Function

Function YellowGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreen = Msg
End Function

Function YellowPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurple = Msg
End Function

Function YellowRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRed = Msg
End Function


'Pre-set 3 Color fade combinations begin here


Function BlackBlueBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlueBlack = Msg
End Function

Function BlackGreenBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreenBlack = Msg
End Function

Function BlackGreyBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreyBlack = Msg
End Function

Function BlackPurpleBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurpleBlack = Msg
End Function

Function BlackRedBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRedBlack = Msg
End Function

Function BlackYellowBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellowBlack = Msg
End Function

Function BlueBlackBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlackBlue = Msg
End Function

Function BlueGreenBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreenBlue = Msg
End Function

Function BluePurpleBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurpleBlue = Msg
End Function

Function BlueRedBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRedBlue = Msg
End Function

Function BlueYellowBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellowBlue = Msg
End Function

Function GreenBlackGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlackGreen = Msg
End Function

Function GreenBlueGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlueGreen = Msg
End Function

Function GreenPurpleGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurpleGreen = Msg
End Function

Function GreenRedGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRedGreen = Msg
End Function

Function GreenYellowGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellowGreen = Msg
End Function

Function GreyBlackGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlackGrey = Msg
End Function

Function GreyBlueGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlueGrey = Msg
End Function

Function GreyGreenGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreenGrey = Msg
End Function

Function GreyPurpleGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurpleGrey = Msg
End Function

Function GreyRedGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRedGrey = Msg
End Function

Function GreyYellowGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellowGrey = Msg
End Function

Function PurpleBlackPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlackPurple = Msg
End Function

Function PurpleBluePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBluePurple = Msg
End Function

Function PurpleGreenPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreenPurple = Msg
End Function

Function PurpleRedPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRedPurple = Msg
End Function

Function PurpleYellowPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellowPurple = Msg
End Function

Function RedBlackRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlackRed = Msg
End Function

Function RedBlueRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlueRed = Msg
End Function

Function RedGreenRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreenRed = Msg
End Function

Function RedPurpleRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurpleRed = Msg
End Function

Function RedYellowRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellowRed = Msg
End Function

Function YellowBlackYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlackYellow = Msg
End Function

Function YellowBlueYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlueYellow = Msg
End Function

Function YellowGreenYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreenYellow = Msg
End Function

Function YellowPurpleYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurpleYellow = Msg
End Function

Function YellowRedYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRedYellow = Msg
End Function


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    A = Hex(RGB)
    B = Len(A)
    If B = 5 Then A = "0" & A
    If B = 4 Then A = "00" & A
    If B = 3 Then A = "000" & A
    If B = 2 Then A = "0000" & A
    If B = 1 Then A = "00000" & A
    RGBtoHEX = A
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
