' EL354.FRM
Option Explicit
Dim m001A As Variant

Sub Command3D2_Click ()
End Sub

Sub Command3D3_Click ()
Dim l002C As Variant
Dim l0032 As Variant
Dim l0036 As String
Dim l003A As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004E As Variant
Dim l0054 As Variant
Dim l005A As Variant
Dim l0060 As Variant
Dim l0066 As Variant
Dim l006A As Variant
l002C = extfn1018("AOL FRAME25", 0&)
l0032 = extfn10F8(l002C, "Welcome")
l0036$ = String(30, 0)
l003A = extfnFA8(l0032, l0036$, 250)
If  l003A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You Must Enter In Something", 16
Exit Sub
End If
l0044 = extfn1018("AOL FRAME25", 0&)
l0048 = extfn10F8(l0044, "AOL Live!")
l004E = extfn1088(l0048)
l0054 = extfn10C0(l004E, "_AOL_EDIT")
l005A = extfn11A0(l0054, 12, 0, Text2)
l0060 = extfn1168(l0054, 2)
l0066 = extfnFE0(l0060, &H201, 0, 0)
l006A = extfnFE0(l0060, &H202, 0, 0)
Text2.Text = ""
DoEvents
End Sub

Sub Command3D4_Click ()
Text2 = ""
End Sub

Sub Form_GotFocus ()
Text1.SetFocus
End Sub

Sub Form_Load ()
extsub1050 el354.hWnd, -1, 0, 0, 0, 0, 3
Combo1.AddItem "ELiTe"
Combo1.AddItem "Backwards"
Combo1.AddItem "Case"
Combo1.ListIndex = 0
End Sub

Sub Form_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.SetFocus
End Sub

Sub text1_KeyPress (KeyAscii As Integer)
Dim l0088 As Variant
Dim l008C As Variant
Text1 = ""
l0088 = Chr(KeyAscii)
If  l0088 = Chr(13) Then
Command3D3_Click
GoTo L5F6C8
End If
If  Combo1.Text = "ELiTe" Then
Select Case l0088
Case "A"
Text2 = Text2 + "À"
Text1 = ""
Case "B"
Text1 = ""
Text2 = Text2 + "ß"
Case "C"
Text1 = ""
Text2 = Text2 + "Ç"
Case "D"
Text1 = ""
Text2 = Text2 + "Ð"
Case "E"
Text1 = ""
Text2 = Text2 + "Ë"
Case "F"
Text1 = ""
Text2 = Text2 + "ƒ"
Case "G"
Text1 = ""
Text2 = Text2 + "G"
Case "H"
Text1 = ""
Text2 = Text2 + "H"
Case "I"
Text1 = ""
Text2 = Text2 + "ï"
Case "J"
Text1 = ""
Text2 = Text2 + "J"
Case "K"
Text1 = ""
Text2 = Text2 + "k"
Case "L"
Text1 = ""
Text2 = Text2 + "L"
Case "M"
Text1 = ""
Text2 = Text2 + "M"
Case "N"
Text1 = ""
Text2 = Text2 + "ñ"
Case "O"
Text1 = ""
Text2 = Text2 + "ð"
Case "P"
Text1 = ""
Text2 = Text2 + "þ"
Case "Q"
Text1 = ""
Text2 = Text2 + "Ø"
Case "R"
Text1 = ""
Text2 = Text2 + "R"
Case "S"
Text1 = ""
Text2 = Text2 + "$"
Case "T"
Text1 = ""
Text2 = Text2 + "†"
Case "U"
Text1 = ""
Text2 = Text2 + "Û"
Case "V"
Text1 = ""
Text2 = Text2 + "V"
Case "W"
Text1 = ""
Text2 = Text2 + "w"
Case "X"
Text1 = ""
Text2 = Text2 + "×"
Case "Y"
Text1 = ""
Text2 = Text2 + "¥"
Case "Z"
Text1 = ""
Text2 = Text2 + "z"
Case "a"
Text1 = ""
Text2 = Text2 + "À"
Case "b"
Text1 = ""
Text2 = Text2 + "ß"
Case "c"
Text1 = ""
Text2 = Text2 + "Ç"
Case "d"
Text1 = ""
Text2 = Text2 + "Ð"
Case "e"
Text1 = ""
Text2 = Text2 + "Ë"
Case "f"
Text1 = ""
Text2 = Text2 + "ƒ"
Case "g"
Text1 = ""
Text2 = Text2 + "g"
Case "h"
Text1 = ""
Text2 = Text2 + "h"
Case "i"
Text1 = ""
Text2 = Text2 + "ï"
Case "j"
Text1 = ""
Text2 = Text2 + "J"
Case "k"
Text1 = ""
Text2 = Text2 + "k"
Case "l"
Text1 = ""
Text2 = Text2 + "L"
Case "m"
Text1 = ""
Text2 = Text2 + "m"
Case "n"
Text1 = ""
Text2 = Text2 + "ñ"
Case "o"
Text1 = ""
Text2 = Text2 + "ð"
Case "p"
Text1 = ""
Text2 = Text2 + "þ"
Case "q"
Text1 = ""
Text2 = Text2 + "Ø"
Case "r"
Text1 = ""
Text2 = Text2 + "R"
Case "s"
Text1 = ""
Text2 = Text2 + "$"
Case "t"
Text1 = ""
Text2 = Text2 + "†"
Case "u"
Text1 = ""
Text2 = Text2 + "Û"
Case "v"
Text1 = ""
Text2 = Text2 + "V"
Case "w"
Text1 = ""
Text2 = Text2 + "w"
Case "x"
Text1 = ""
Text2 = Text2 + "×"
Case "y"
Text1 = ""
Text2 = Text2 + "¥"
Case "z"
Text1 = ""
Text2 = Text2 + "z"
Case "1"
Text1 = ""
Text2 = Text2 + "¹"
Case "2"
Text1 = ""
Text2 = Text2 + "²"
Case "3"
Text1 = ""
Text2 = Text2 + "³"
Case "4"
Text1 = ""
Text2 = Text2 + "4"
Case "5"
Text1 = ""
Text2 = Text2 + "5"
Case "6"
Text1 = ""
Text2 = Text2 + "6"
Case "7"
Text1 = ""
Text2 = Text2 + "7"
Case "8"
Text1 = ""
Text2 = Text2 + "8"
Case "9"
Text1 = ""
Text2 = Text2 + "9"
Case "0"
Text1 = ""
Text2 = Text2 + "º"
Case "!"
Text1 = ""
Text2 = Text2 + "¡"
Case "@"
Text1 = ""
Text2 = Text2 + "@"
Case "#"
Text1 = ""
Text2 = Text2 + "#"
Case "$"
Text1 = ""
Text2 = Text2 + "S"
Case "%"
Text1 = ""
Text2 = Text2 + "%"
Case "^"
Text1 = ""
Text2 = Text2 + "^"
Case "&"
Text1 = ""
Text2 = Text2 + "&"
Case "*"
Text1 = ""
Text2 = Text2 + "*"
Case "("
Text1 = ""
Text2 = Text2 + "("
Case ")"
Text1 = ""
Text2 = Text2 + ")"
Case "-"
Text1 = ""
Text2 = Text2 + "-"
Case "_"
Text1 = ""
Text2 = Text2 + "_"
Case "="
Text1 = ""
Text2 = Text2 + "="
Case "+"
Text1 = ""
Text2 = Text2 + "+"
Case "|"
Text1 = ""
Text2 = Text2 + "|"
Case "\"
Text1 = ""
Text2 = Text2 + "\"
Case "/"
Text1 = ""
Text2 = Text2 + "/"
Case "`"
Text1 = ""
Text2 = Text2 + "`"
Case "~"
Text1 = ""
Text2 = Text2 + "~"
Case ","
Text1 = ""
Text2 = Text2 + ","
Case "<"
Text1 = ""
Text2 = Text2 + "<"
Case "."
Text1 = ""
Text2 = Text2 + "."
Case ">"
Text1 = ""
Text2 = Text2 + ">"
Case ";"
Text1 = ""
Text2 = Text2 + ";"
Case ":"
Text1 = ""
Text2 = Text2 + ":"
Case "'"
Text1 = ""
Text2 = Text2 + "'"
Case " "
Text1 = ""
Text2 = Text2 + " "
Case "["
Text1 = ""
Text2 = Text2 + "["
Case "{"
Text1 = ""
Text2 = Text2 + "{"
Case "]"
Text1 = ""
Text2 = Text2 + "]"
Case "}"
Text1 = ""
Text2 = Text2 + "}"
Case "?"
Text1 = ""
Text2 = Text2 + "¿"
Case Chr(8)
If  Len(Text2) = 0 Then Exit Sub
Text2.Text = Mid(Text2, 1, Len(Text2) - 1)
If  Len(Text2) = 0 Then Exit Sub
Case Else
Text2.Text = Text2.Text + l0088
End Select
End If
If  Combo1.Text = "Case" Then
If  l0088 = Chr(8) Then
On Error Resume Next
Text2.Text = Mid(Text2.Text, 1, Len(Text2.Text) - 1)
GoTo L5F6C8
End If
If  m001A = 0 Then
l008C = UCase(l0088)
Text2.Text = Text2.Text + l008C
m001A = 1
GoTo L5F6C8
End If
If  m001A = 1 Then
l008C = LCase(l0088)
Text2.Text = Text2.Text + l008C
m001A = 0
GoTo L5F6C8
End If
End If
If  Combo1.Text = "Backwards" Then
If  l0088 = Chr(8) Then
Text2.Text = Mid(Text2.Text, 2, Len(Text2.Text))
GoTo L5F6C8
End If
Text2.Text = l0088 + Text2.Text
End If

L5F6C8:
Text1 = ""
Text2.SelStart = Len(Text2.Text)
End Sub

Sub Text2_GotFocus ()
Text1.SetFocus
End Sub

Sub Text2_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.SetFocus
End Sub

Sub sub9980 ()
Unload el354
End Sub
