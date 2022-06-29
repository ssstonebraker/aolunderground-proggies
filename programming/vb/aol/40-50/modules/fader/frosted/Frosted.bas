Attribute VB_Name = "Module1"
'Fro$ted.bas By F®o§§Deé
'Yo I wanna give thanx to ravage bas
'bytefader , and cryofader for makin me wanna
'make a fader bas. I used some stuff from
'those bas's and added some stuff of my own
'This bas is tight I have to say
'If any questions or comments Email me at
'FroSSDee@hotmail.com
'This bas is made for 32 bit Aol 4.0 programmers

Function BlackBlueBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackBlueBlack = msg
End Function

Function BlackGreenBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGreenBlack = msg
End Function
Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGreyBlack = msg
End Function
Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackPurpleBlack = msg
End Function
Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackYellowBlack = msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueBlackBlue = msg
End Function
Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueGreenBlue = msg
End Function
Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BluePurpleBlue = msg
End Function
Function BlueRedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueRedBlue = msg
End Function
Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueYellowBlue = msg
End Function
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
Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlackGreen = msg
End Function
Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlueGreen = msg
End Function
Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenPurpleGreen = msg
End Function
Function GreenRedGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenRedGreen = msg
End Function
Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenYellowGreen = msg
End Function
Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlackGrey = msg
End Function
Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlueGrey = msg
End Function
Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyGreenGrey = msg
End Function
Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyPurpleGrey = msg
End Function
Function GreyRedGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyRedGrey = msg
End Function
Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyYellowGrey = msg
End Function
Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlackPurple = msg
End Function
Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBluePurple = msg
End Function
Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleGreenPurple = msg
End Function
Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleRedPurple = msg
End Function
Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleYellowPurple = msg
End Function
Function RedBlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlackRed = msg
End Function
Function RedBlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlueRed = msg
End Function
Function RedGreenRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedGreenRed = msg
End Function
Function RedPurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedPurpleRed = msg
End Function
Function RedYellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    RedYellowRed = msg
End Function
Function RGB2HEX(r, G, b)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = b
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
Function TrimSpaces(text)
    If InStr(text, " ") = 0 Then
    TrimSpaces = text
    Exit Function
    End If
    For TrimSpace = 1 To Len(text)
    thechar$ = Mid(text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function
Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlackYellow = msg
End Function
Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlueYellow = msg
End Function
Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowGreenYellow = msg
End Function
Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPurpleYellow = msg
End Function
Function YellowRedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowRedYellow = msg
End Function

















Function WavY(thetext As String)
G$ = thetext
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & u$ & "<sub>" & s$ & "</sub>" & T$
Next w
WavY = P$
End Function

Function DBlue_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    DBlue_Black = msg
End Function
Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
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
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    DGreen_Black = msg
End Function
Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
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
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_DBlue_LBlue = msg
End Function
Function LBlue_Green(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Green = msg
End Function
Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Green_LBlue = msg
End Function
Function LBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
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
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange_LBlue = msg
End Function
Function LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Yellow = msg
End Function
Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Yellow_LBlue = msg
End Function
Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
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
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LGreen_DGreen_LGreen = msg
End Function



Function Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Yellow_LBlue = msg
End Function
    Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    Yellow_LBlue_Yellow = msg
End Function
Sub FadeFormHorizon(TheForm As Form)

TheForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
TheForm.Line (0, b)-(TheForm.Width, b + 2), RGB(a + 3, a, a * 3), BF
b = b + 2
Next a
End Sub



             Sub Fire(PiSS As Object)
Dim X
Dim Y
Dim Red
Dim Green
Dim Blue
X = PiSS.Width
Y = PiSS.Height
Red = 255
Green = 255
Blue = 255
Do Until Red = 0
Y = Y - PiSS.Height / 255 * 1
Red = Red - 1
PiSS.Line (0, 0)-(X, Y), RGB(255, Red, 0), BF
Loop
End Sub























Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
ChatSend (P$)
End Sub
Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRB = P$
End Function
Function WavyChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavYChaTRG = P$
End Function
Function WhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    WhitePurple = msg
End Function
Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    WhitePurpleWhite = msg
End Function

Sub UnderlineSendChat(UnderlineChat)
ChatSend ("<U>" & UnderlineChat & "</U>")
End Sub
Sub StikeOutSendChat(StrikeOutChat)
ChatSend ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub ItalicSendChat(ItalicChat)
ChatSend ("<I>" & ItalicChat & "</I>")
End Sub
Sub BoldItalicSendChat(BoldChat)
ChatSend ("<I><B>" & BoldChat & "</I></B>")
End Sub

Sub BoldStrikeOutSendChat(BoldChat)
ChatSend ("<b><S>" & BoldChat & "</b></S>")
End Sub

Sub BoldUnderlineSendChat(BoldChat)
ChatSend ("<b><U>" & BoldChat & "</b></U>")
End Sub
Sub StrikeoutUnderlineSendChat(BoldChat)
ChatSend ("<S><U>" & BoldChat & "</S></U>")
End Sub
Sub StrikeOutItalicSendChat(BoldChat)
ChatSend ("<S><I>" & BoldChat & "</S></I>")
End Sub

Sub ItalicUnderlineSendChat(BoldChat)
ChatSend ("<I><U>" & BoldChat & "</I></U>")
End Sub
Sub AllChat(BoldChat)
ChatSend ("<b><I><S><U>" & BoldChat & "</b></I></S></U>")
End Sub

Sub BoldSendChat(BoldChat)
ChatSend ("<b>" & BoldChat & "</b>")
End Sub

Sub BoldWavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
BoldSendChat (P$)
End Sub

Sub BoldWavyColorbluegree(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next w
BoldSendChat (P$)
End Sub
Sub BoldWavyColorredandblack(thetext)

G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & T$
Next w
BoldSendChat (P$)
End Sub
Sub BoldWavyColorredandblue(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & T$
Next w
BoldSendChat (P$)
End Sub
Function BoldWhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    WhitePurpleWhite (msg)
End Function
Function BoldYellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldYellowBlackYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldYellowBlueYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldYellowGreenYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldYellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BoldSendChat (msg)
End Function
Function BoldYellowPurpleYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldYellowRedYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function Elite(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "]V["
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let newsent$ = newsent$ + NextChr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function








Function Elite2(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "]V["
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let newsent$ = newsent$ + NextChr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function



Function Hacker(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
Loop
BoldBlackBlueBlack (newsent$)


End Function
Function Spaced(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
Loop
 BoldRedBlackRed (newsent$)

End Function





Function Bold_italic_colorR_Backwards(strin As String)

Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = NextChr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function
Function BoldAOL4_WavColors(Text1 As String)
G$ = Text1
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next w
ChatSend (P$)
End Function
Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next w
BoldSendChat (P$)
End Function
Function BoldBlack_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, F, F - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BoldSendChat (msg)
End Function
Function BoldBlackBlueBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldBlackBlueBlack2(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<B><U><Font Color=#" & h & ">" & d
    Next b
  ChatSend (msg)
End Function
Function BoldBlackGreenBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldBlackGreyBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldBlackRedBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldBlackYellowBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldBlueBlackBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldBlueGreenBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldBlueRedBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldBlueYellowBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldDBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 450 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BoldSendChat (msg)
End Function
Function BoldDGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldGreenBlackGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
ChatSend (msg)
End Function
Function BoldGreenBlueGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
  ChatSend (msg)
End Function
Function BoldGreenPurpleGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
  ChatSend (msg)
End Function
Function BoldGreenRedGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
 ChatSend (msg)
End Function
Function BoldGreenYellowGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
  ChatSend (msg)
End Function
Function BoldGreyBlackGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldGreyBlueGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
 BoldSendChat (msg)
End Function
Function BoldGreyGreenGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldGreyPurpleGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldGreyRedGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldGreyYellowGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function Bolditalic_BlackPurpleBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<B><I><Font Color=#" & h & ">" & d
    Next b
   ChatSend (msg)
End Function
Function Bolditalic_BluePurpleBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<B><I><Font Color=#" & h & ">" & d
    Next b
ChatSend (msg)
End Function
Function BoldLBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(355, 255 - F, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
 BoldSendChat (msg)
End Function
Function BoldLBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldLBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Green_LBlue (msg)
End Function
Function BoldLBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange_LBlue (msg)
End Function
Function BoldLBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 155, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange (msg)
End Function
Function BoldLBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Yellow_LBlue (msg)
End Function
Function BoldLGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    LGreen_DGreen_LGreen (msg)
End Function
Function BoldPinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
    BoldSendChat (msg)
End Function
Function BoldPurple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldPurpleBlackPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldPurpleBluePurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
 BoldSendChat (msg)
End Function
Function BoldPurpleGreenPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & h & ">" & d
    Next b
 BoldSendChat (msg)
End Function
Function BoldPurpleRedPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldPurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldPurpleYellowPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
 BoldSendChat (msg)
End Function
Function BoldRedBlackRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldRedBlueRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldSendChat (msg)
End Function
Function BoldRedGreenRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function
Function BoldRedPurpleRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (msg)
End Function
Function BoldRedYellowRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (msg)
End Function














































































Function BlackRedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
BlackRedBlack = msg
End Function
Function PinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 200 / a
        F = e * b
        G = RGB(255 - F, 167, 510)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & d
    Next b
PinkOrange = msg
End Function



Sub Form_BlueFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub
Sub Form_CircleFire(frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = frm.Width
Y = frm.Height
frm.FillStyle = 0
Red = 0
Blue = frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - frm.Width / 255 * 1
frm.FillColor = RGB(255, Red, 0)
If Blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), Blue, RGB(255, Red, 0)
Loop
End Sub
Sub Form_CircleRedFlare(frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = frm.Width
Y = frm.Height
frm.FillStyle = 0
Red = 0
Blue = frm.Width
Do Until Red = 255
Red = Red + 5
Blue = Blue - frm.Width / 255 * 10
frm.FillColor = RGB(Red, 0, 0)
If Blue < 0 Then Exit Do
frm.Circle (frm.Width / 2, frm.Height / 2), Blue, RGB(255, Red, 0)
Loop
End Sub
Sub Form_FadeBlue(frm As Object)
Dim X
Dim Y
Dim Red
Dim Green
Dim Blue
X = frm.Width
Y = frm.Height
Red = 255
Green = 255
Blue = 255
Do Until Red = 0
Y = Y - frm.Height / 255 * 1
Red = Red - 1
frm.Line (0, 0)-(X, Y), RGB(0, 0, Red), BF
Loop
End Sub
Sub Form_FadeFire(frm As Object)
Dim X
Dim Y
Dim Red
Dim Green
Dim Blue
X = frm.Width
Y = frm.Height
Red = 255
Green = 255
Blue = 255
Do Until Red = 0
Y = Y - frm.Height / 255 * 1
Red = Red - 1
frm.Line (0, 0)-(X, Y), RGB(255, Red, 0), BF
Loop
End Sub
Sub Form_Flash(frm As Form)
frm.Show
frm.BackColor = &H0&
pause (".1")
frm.BackColor = &HFF&
pause (".1")
frm.BackColor = &HFF0000
pause (".1")
frm.BackColor = &HFF00&
pause (".1")
frm.BackColor = &H8080FF
pause (".1")
frm.BackColor = &HFFFF00
pause (".1")
frm.BackColor = &H80FF&
pause (".1")
frm.BackColor = &HC0C0C0
End Sub

Sub Form_GreenFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub Form_IceFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B
    Next intLoop
End Sub

Sub Form_RedFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub
Sub Form_SilverFade(frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 2
    frm.ScaleHeight = 256
    For intLoop = 0 To 255
    frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub


