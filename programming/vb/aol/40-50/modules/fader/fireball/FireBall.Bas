Attribute VB_Name = "FireBall"
'FiRE·BaLL
'This is a Form Fading .Bas file so the version
'of AOL you intend to program for doesn't matter
'example :
'SideFire Me
'that's all it takes, choose one of the sub forb this .bas and
'stick that code in your form's paint
'AND FOR THE LOVE OF GOD USE AT LEAST HIGH COLOR!!!!!
'OR IT WILL LOOK HORRIBLE!!!!!!!!!!!!!!!!!!!!!!!!!
'WARNING: the circle fades on this .Bas are EXTREMLEY slow
'but they still are pretty fancy.

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
