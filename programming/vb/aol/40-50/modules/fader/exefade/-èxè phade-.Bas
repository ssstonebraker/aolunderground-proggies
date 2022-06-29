Attribute VB_Name = "exe"
Sub AboutThis()
'hey this is exe you can im me @
'aimexe, with any questions or comments.
'oh ya & one more thing, this is way better than
'FireBall.bas ;x - sorry dude..., but it is.. ;]
'i'd like to thank FireBall.bas for starting me on
'this. this is for aol4.o & version¹
End Sub
Sub Blue(PiSS As Object)
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
PiSS.Line (0, 0)-(X, Y), RGB(0, 0, Red), BF
Loop
End Sub
Sub SideRed(PiSS As Object)
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
X = X - PiSS.Width / 255 * 1
Red = Red - 1
PiSS.Line (0, 0)-(X, Y), RGB(Red, 0, 0), BF
Loop
End Sub
Sub SideGreen(PiSS As Object)
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
X = X - PiSS.Width / 255 * 1
Red = Red - 1
PiSS.Line (0, 0)-(X, Y), RGB(0, Red, 0), BF
Loop
End Sub
Sub SideFire(PiSS As Object)
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
X = X - PiSS.Width / 255 * 1
Red = Red - 1
PiSS.Line (0, 0)-(X, Y), RGB(255, Red, 0), BF
Loop
End Sub
Sub SideBlue(PiSS As Object)
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
X = X - PiSS.Width / 255 * 1
Red = Red - 1
PiSS.Line (0, 0)-(X, Y), RGB(0, 0, Red), BF
Loop
End Sub
Sub Red(PiSS As Object)
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
PiSS.Line (0, 0)-(X, Y), RGB(Red, 0, 0), BF
Loop
End Sub
Sub Green(PiSS As Object)
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
PiSS.Line (0, 0)-(X, Y), RGB(0, Red, 0), BF
Loop
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
Sub CircleRed(PiSS As Object)
Dim X
Dim Y
Dim Red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
Red = 0
TuRd = PiSS.Width
Do Until Red = 255
Red = Red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(Red, 0, 0)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(Red, 0, 0)
Loop
End Sub
Sub CircleGreen(PiSS As Object)
Dim X
Dim Y
Dim Red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
Red = 0
TuRd = PiSS.Width
Do Until Red = 255
Red = Red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(0, Red, 0)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(0, Red, 0)
Loop
End Sub
Sub CircleFire(PiSS As Object)
Dim X
Dim Y
Dim Red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
Red = 0
TuRd = PiSS.Width
Do Until Red = 255
Red = Red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(255, Red, 0)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(255, Red, 0)
Loop
End Sub
Sub CircleBlue(PiSS As Object)
Dim X
Dim Y
Dim Red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
Red = 0
TuRd = PiSS.Width
Do Until Red = 255
Red = Red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(0, 0, Red)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(0, 0, Red)
Loop
End Sub



Sub BarFadeFrm(frm, Style)
'This fade makes a pipe like effect on the form
'Code example:
'BarFadeFrm Me, 1


Dim cx, cy, f, F1, F2, i    ' Declare variables
frm.AutoRedraw = True
frm.Cls
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2

DrawWidth = 2   ' Set DrawWidth.
    
For i = 255 To 0 Step -2

f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.

If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black

frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF
Next i
End Sub

Sub CFadeFrm(frm, Style)
'This fade creates a radial fade of circles
'Code example:
'CFadeFrm Me, 1

frm.AutoRedraw = True
frm.Cls

Dim cx, cy, i
    frm.ScaleMode = 3
    cx = frm.ScaleWidth \ 2
    cy = frm.ScaleHeight \ 2
'    If cx > cy Then limit = cy Else limit = cx
frm.DrawWidth = 2

For i = 0 To 255
If Style = 1 Then frm.Circle (cx, cy), i, RGB(i, i, i)  'Black to white
If Style = 2 Then frm.Circle (cx, cy), i, RGB(0, i, i)  'Black to Cyan
If Style = 3 Then frm.Circle (cx, cy), i, RGB(i, 0, i)  'Black to Purple
If Style = 4 Then frm.Circle (cx, cy), i, RGB(i, i, 0)  'Black to Yellow
If Style = 5 Then frm.Circle (cx, cy), i, RGB(0, 0, i)  'Black to Blue
If Style = 6 Then frm.Circle (cx, cy), i, RGB(i, 0, 0)  'Black to Red
If Style = 7 Then frm.Circle (cx, cy), i, RGB(0, i, 0)  'Black to Green
If Style = 8 Then frm.Circle (cx, cy), i, RGB(0, i, 255)  'Blue to Green
If Style = 9 Then frm.Circle (cx, cy), i, RGB(i, i, 255)  'Blue to White
If Style = 11 Then frm.Circle (cx, cy), i, RGB(i, 0, 255)  'Blue to Purple
If Style = 12 Then frm.Circle (cx, cy), i, RGB(0, 0, 255 - i)  'Blue to Black
If Style = 13 Then frm.Circle (cx, cy), i, RGB(255, 0, i)  'Red to Purple
If Style = 14 Then frm.Circle (cx, cy), i, RGB(255, i, i)  'Red to White
If Style = 15 Then frm.Circle (cx, cy), i, RGB(255, i, 0)  'Red to Yellow
If Style = 16 Then frm.Circle (cx, cy), i, RGB(255 - i, 0, 0)  'Red to Black
If Style = 17 Then frm.Circle (cx, cy), i, RGB(i, 255, i)  'Green to White
If Style = 18 Then frm.Circle (cx, cy), i, RGB(0, 255, i)  'Green to Blue
If Style = 19 Then frm.Circle (cx, cy), i, RGB(i, 255, 0)  'Green to Yellow
If Style = 20 Then frm.Circle (cx, cy), i, RGB(0, 255 - i, 0)  'Green to Black
If Style = 21 Then frm.Circle (cx, cy), i, RGB(255 - i, 255 - i, 255 - i)  'White to Black
If Style = 22 Then frm.Circle (cx, cy), i, RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.Circle (cx, cy), i, RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.Circle (cx, cy), i, RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.Circle (cx, cy), i, RGB(255, 255, i)  'Yellow to White
If Style = 26 Then frm.Circle (cx, cy), i, RGB(255, i, 255)  'Purple to White
If Style = 27 Then frm.Circle (cx, cy), i, RGB(i, 255, 255)  'Cyan to White
If Style = 28 Then frm.Circle (cx, cy), i, RGB(255 - i, 255 - i, 0)  'Yellow to Black
If Style = 29 Then frm.Circle (cx, cy), i, RGB(255 - i, 0, 255 - i)  'Purple to Black
If Style = 30 Then frm.Circle (cx, cy), i, RGB(0, 255 - i, 255 - i)  'Cyan to Black
If Style = 31 Then frm.Circle (cx, cy), i, RGB(S1 - i, S2 - i, S3 - i)  'Selected color to black

Next i
End Sub
Sub DoubleFade(frm, Style)
'This fade combines the explosive fade and the rectangular fade
'Code example:
'Double Me, 1

frm.AutoRedraw = True
frm.Cls

Dim cx, cy, f, F1, F2, i    ' Declare variables
frm.AutoRedraw = True
frm.Cls
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2

DrawWidth = 2   ' Set DrawWidth.
    
For i = 255 To 0 Step -2

f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.

If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black

frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF

Next i
    frm.ScaleMode = 3   ' Set ScaleMode to pixels.
    cx = frm.ScaleWidth / 2 ' Get horizontal center.
    cy = frm.ScaleHeight / 2    ' Get vertical center.
    frm.DrawWidth = 2

For i = 0 To 255
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black
f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.
        
frm.Line (cx * F1, cy)-(cx, cy * F1)   ' Draw upper-left.
frm.Line -(cx * F2, cy) ' Draw upper-right.
frm.Line -(cx, cy * F2) ' Draw lower-right.
frm.Line -(cx * F1, cy) ' Draw lower-left.
Next i
End Sub

Sub ExplosiveFade(frm, Style)
'This fade creates a star like explosion on the form
'Code example:
'ExplosiveFade Me, 1

frm.AutoRedraw = True
frm.Cls
Dim cx, cy, f, F1, F2, i    ' Declare variables
    frm.ScaleMode = 3   ' Set ScaleMode to pixels.
    cx = frm.ScaleWidth / 2 ' Get horizontal center.
    cy = frm.ScaleHeight / 2    ' Get vertical center.
    frm.DrawWidth = 2

For i = 0 To 255
If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black
f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.
        
frm.Line (cx * F1, cy)-(cx, cy * F1)   ' Draw upper-left.
frm.Line -(cx * F2, cy) ' Draw upper-right.
frm.Line -(cx, cy * F2) ' Draw lower-right.
frm.Line -(cx * F1, cy) ' Draw lower-left.

Next i
End Sub
Sub FadeFrm(frm, Style)
'This fade goes from the top to the bottom of the form
'However, you can change this by using the inverse stlye
' For instance, to go from blue to back, the style is 5
' To go from back to blue, the style is 12

'Code example:
'FadeFrm Me, 1

frm.ScaleMode = vbPixels
frm.AutoRedraw = True
frm.DrawStyle = vbInsideSolid 'Set Form Modes
frm.Cls
frm.DrawWidth = 2
frm.DrawMode = 13

frm.ScaleHeight = 256

For i = 0 To 255
If Style = 1 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, i), BF ' Black to white
If Style = 2 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, i), BF ' Black to Cyan
If Style = 3 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, i), BF ' Black to Purple
If Style = 4 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, 0), BF ' Black to Yellow
If Style = 5 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 0, i), BF ' Black to Blue
If Style = 6 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, 0), BF ' Black to Red
If Style = 7 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, 0), BF ' Black to Green
If Style = 8 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, i, 255), BF ' Blue to Green
If Style = 9 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, 255), BF ' Blue to White
If Style = 11 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 0, 255), BF ' Blue to Purple
If Style = 12 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 0, 255 - i), BF ' Blue to Black
If Style = 13 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 0, i), BF ' Red to Purple
If Style = 14 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, i), BF ' Red to White
If Style = 15 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, 0), BF ' Red to Yellow
If Style = 16 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 0, 0), BF ' Red to Black
If Style = 17 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, i), BF ' Green to White
If Style = 18 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255, i), BF ' Green to Blue
If Style = 19 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, 0), BF ' Green to Yellow
If Style = 20 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255 - i, 0), BF ' Green to Black
If Style = 21 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255 - i, 255 - i), BF ' White to Black
If Style = 22 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255, 255 - i), BF 'White to Yellow
If Style = 23 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255 - i, 255), BF 'White to Purple
If Style = 24 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255, 255), BF 'White to Cyan
If Style = 25 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, 255, i), BF ' Yellow to White
If Style = 26 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255, i, 255), BF ' Purple to White
If Style = 27 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, 255, 255), BF ' Cyan to White
If Style = 28 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 255 - i, 0), BF ' Yellow to Black
If Style = 29 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(255 - i, 0, 255 - i), BF ' Purple to Black
If Style = 30 Then frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(0, 255 - i, 255 - i), BF ' Cyan to Black
If Style = 31 Then If i = 193 Then Exit Sub: frm.Line (0, i)-(frm.ScaleWidth, i - 1), RGB(i, i, i), BF ' black to Gray
Next i
End Sub


Sub RFadeFrm(frm, Style)
'This creates a rectangular fade on the form. I like this one a lot
'Code example:
'RFadeFrm Me, 1

Dim cx, cy, f, F1, F2, i    ' Declare variables
frm.AutoRedraw = True
frm.Cls
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2

DrawWidth = 2   ' Set DrawWidth.
    
For i = 255 To 0 Step -2

f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.

If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black

frm.Line (cx * F1, cy * F1)-(cx * F2, cy * F2), , BF

Next i
End Sub
Sub SideFade(frm, Style)
'This fade goes from the left to the right side of the form
'However, you can change this by using the inverse stlye
' For instance, to go from blue to back, the style is 5
' To go from back to blue, the style is 12

'Example of code:
'SideFade Me, 1
Dim cx, cy, f, F1, F2, i    ' Declare variables
frm.AutoRedraw = True
frm.Cls
    frm.ScaleMode = 3
    cx = frm.ScaleWidth
    cy = frm.ScaleHeight

DrawWidth = 2   ' Set DrawWidth.
    
For i = 255 To 0 Step -2

f = i / 255  ' Perform interim
F1 = 1 - f: F2 = 1 + f  ' calculations.

If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black

frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF

Next i
End Sub


Sub Text3D(Ctrl As Control, text, bevel, Style, Font)
'This fades the text so it looks as though it slides out of the background!
'Code Example:
'Text3d Picture1,"3D Text!", 8,12, "Arial"

Ctrl.AutoRedraw = True
Ctrl.FontSize = bevel * 1.4
Ctrl.Font = Font

For i = 0 To bevel * 10
If Style = 1 Then Ctrl.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then Ctrl.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then Ctrl.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then Ctrl.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then Ctrl.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then Ctrl.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then Ctrl.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then Ctrl.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then Ctrl.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then Ctrl.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then Ctrl.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then Ctrl.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then Ctrl.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then Ctrl.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then Ctrl.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then Ctrl.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then Ctrl.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then Ctrl.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then Ctrl.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then Ctrl.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then Ctrl.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then Ctrl.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then Ctrl.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then Ctrl.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then Ctrl.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then Ctrl.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then Ctrl.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then Ctrl.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then Ctrl.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then Ctrl.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black
Ctrl.CurrentY = i \ 2
Ctrl.CurrentX = i \ 2
Ctrl.Print text
Next i

End Sub

