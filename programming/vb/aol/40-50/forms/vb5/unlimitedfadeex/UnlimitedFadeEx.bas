Attribute VB_Name = "Module1"
Function GetR(A)
C = InStr(A, " ")
GetR = Left(A, C - 1)

End Function
Function GetG(A)
C = InStr(A, " ")
R = Left(A, C - 1)
A = Mid(A, C + 1)
C = InStr(A, " ")
GetG = Left(A, C - 1)

End Function
Function GetB(A)
C = InStr(A, " ")
R = Left(A, C - 1)
A = Mid(A, C + 1)
C = InStr(A, " ")
G = Left(A, C - 1)
GetB = Mid(A, C + 1)

End Function
Function Color(Red As Integer, Green As Integer, Blue As Integer)
Color = Red & " " & Green & " " & Blue

End Function

Public Function ChatFade(what$, R1 As Integer, R2 As Integer, G1 As Integer, G2 As Integer, B1 As Integer, B2 As Integer, MakeWavy As Boolean) As String
textlen = Len(what$)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(what$, MakeFade, 1)
    RGBVal = RGB((B2 - B1) / textlen * MakeFade + B1, (G2 - G1) / textlen * MakeFade + G1, (R2 - R1) / textlen * MakeFade + R1)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy
        End If
    End If
AfterWavy:
FadedText = FadedText + "<FONT  COLOR=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
ChatFade = FadedText

End Function
Function RGB2HEX(RedGreenBlue) As String
'LiviD Rocks =]
HexVal$ = Hex(RedGreenBlue)
ZeroFact% = Len(HexVal$)
HexVal$ = String(6 - ZeroFact%, "0") + HexVal$
RGB2HEX = HexVal$
End Function

