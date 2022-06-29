Attribute VB_Name = "TheBird"
Const MaxBirds = 6 '1 to 12

Const FOLLOW = &H1    'Instruct Bird to Flock with closest friend.
Const NOFOLLOW = &H10 'Instruct Bird to Fly to Objective Coords.
Const NONE = &H0      'Null

Const RESETSPACE = -&H20 'Far place for dead birds, untill bird is recycled.
Const LifeTime = 4000    'If the bird is not visible for this interval, Kill It.

Declare Function timeGetTime Lib "winmm.dll" () As Long 'Keep track of birds lifetime.

Public Type Point2D 'Coords.
X As Integer
Y As Integer
End Type

Public Type Bird
coord As Point2D
a As Integer
v As Integer
frame As Integer
Entrypoint As Point2D
Exitpoint As Point2D
flying As Boolean
StartTime As Long
Flag As Long
End Type: Public B(MaxBirds - 1) As Bird

Public Const Pi = 3.14159265358979
Public Sine(359) As Single, CoSn(35) As Single, LI As Integer, LII As Integer, LR As Integer

Sub Math_BTT()
For i = 0 To 35
Sine(i) = Sin(i * (Pi / 18))
CoSn(i) = Cos(i * (Pi / 18))
Next
End Sub

Sub NewBird(): Dim BNum As Integer
For LI = 0 To MaxBirds - 1
If B(LI).flying = False Then BNum = LI: GoTo BirdReady
Next
Exit Sub
BirdReady:
Randomize Rnd * 255 'SuperSeed
LR = Rnd * 6
Select Case LR
Case 1
B(BNum).Entrypoint.X = Rnd * 325
B(BNum).Entrypoint.Y = -20
B(BNum).Exitpoint.X = Rnd * 325
B(BNum).Exitpoint.Y = 377
Case 2
B(BNum).Entrypoint.X = Rnd * 325
B(BNum).Entrypoint.Y = -20
B(BNum).Exitpoint.X = 377
B(BNum).Exitpoint.Y = Rnd * 325
Case 4
B(BNum).Entrypoint.X = -20
B(BNum).Entrypoint.Y = Rnd * 377
B(BNum).Exitpoint.X = 350
B(BNum).Exitpoint.Y = Rnd * 377
Case 5
B(BNum).Entrypoint.X = -20
B(BNum).Entrypoint.Y = Rnd * 377
B(BNum).Exitpoint.X = Rnd * 325
B(BNum).Exitpoint.Y = 377
Case Else
B(BNum).Entrypoint.X = Rnd * 325
B(BNum).Entrypoint.Y = -20
B(BNum).Exitpoint.X = -20
B(BNum).Exitpoint.Y = Rnd * 325
End Select

B(BNum).v = (Rnd * 1) + 1
B(BNum).flying = True
B(BNum).frame = 1
B(BNum).a = 0
B(BNum).coord.X = B(BNum).Entrypoint.X
B(BNum).coord.Y = B(BNum).Entrypoint.Y
B(BNum).StartTime = timeGetTime

LR = Rnd * 100
If LR >= 25 Then
B(BNum).Flag = NOFOLLOW
Else
B(BNum).Flag = FOLLOW
End If
End Sub

Sub DoBird(): Dim testX As Single, testY As Single, TestAngle As Integer, TestDis As Single, MinDis As Single, TestDis2 As Single, MinDis2 As Single, CDir As Integer, CDir2 As Integer, Nearestbuddy As Integer
On Error Resume Next
MinDis = 10000: MinDis2 = 10000
For LI = 0 To MaxBirds - 1
If B(LI).flying Then

If B(LI).coord.X < -10 Or B(LI).coord.X > W + 10 Or B(LI).coord.Y < -10 Or B(LI).coord.Y > H + 10 Then
If timeGetTime - B(LI).StartTime > LifeTime Then KillBird LI
End If

TestAngle = B(LI).a
testX = B(LI).coord.X + ((B(LI).v) * Sine(TestAngle))
testY = B(LI).coord.Y + ((B(LI).v) * CoSn(TestAngle))
TestDis = Sqr(((B(LI).Exitpoint.X - testX) ^ 2) + ((B(LI).Entrypoint.Y - testY) ^ 2))
If TestDis < MinDis Then MinDis = TestDis: CDir = 1
TestAngle = B(LI).a + 1
If TestAngle = 36 Then TestAngle = 0
testX = B(LI).coord.X + ((B(LI).v) * Sine(TestAngle))
testY = B(LI).coord.Y + ((B(LI).v) * CoSn(TestAngle))
TestDis = Sqr(((B(LI).Exitpoint.X - testX) ^ 2) + ((B(LI).Exitpoint.Y - testY) ^ 2))
If TestDis < MinDis Then MinDis = TestDis: CDir = 2
TestAngle = B(LI).a - 1
If TestAngle = -1 Then TestAngle = 35
testX = B(LI).coord.X + ((B(LI).v) * Sine(TestAngle))
testY = B(LI).coord.Y + ((B(LI).v) * CoSn(TestAngle))
TestDis = Sqr(((B(LI).Exitpoint.X - testX) ^ 2) + ((B(LI).Exitpoint.Y - testY) ^ 2))
If TestDis < MinDis Then MinDis = TestDis: CDir = 3

If B(LI).Flag = FOLLOW And B(LI).flying Then
For LII = 0 To MaxBirds - 1
If Not LII = LI Then
TestAngle = B(LI).a
testX = B(LI).coord.X + ((B(LI).v) * Sine(TestAngle))
testY = B(LI).coord.Y + ((B(LI).v) * CoSn(TestAngle))
TestDis2 = Sqr(((B(LII).coord.X - testX) ^ 2) + ((B(LII).coord.Y - testY) ^ 2))
If TestDis2 < MinDis2 Then MinDis2 = TestDis2: CDir2 = 1
TestAngle = B(LI).a + 1
If TestAngle = 36 Then TestAngle = 0
testX = B(LI).coord.X + ((B(LI).v) * Sine(TestAngle))
testY = B(LI).coord.Y + ((B(LI).v) * CoSn(TestAngle))
TestDis2 = Sqr(((B(LII).coord.X - testX) ^ 2) + ((B(LII).coord.Y - testY) ^ 2))
If TestDis2 < MinDis2 Then MinDis2 = TestDis2: CDir2 = 2
TestAngle = B(LI).a - 1
If TestAngle = -1 Then TestAngle = 35
testX = B(LI).coord.X + ((B(LI).v) * Sine(TestAngle))
testY = B(LI).coord.Y + ((B(LI).v) * CoSn(TestAngle))
TestDis2 = Sqr(((B(LII).coord.X - testX) ^ 2) + ((B(LII).coord.Y - testY) ^ 2))
If TestDis2 < MinDis2 Then MinDis2 = TestDis2: CDir2 = 3
End If
Next
End If

If MinDis / 2 < MinDis2 Then
Select Case CDir
Case 2
TestAngle = B(LI).a + 1
If TestAngle = 36 Then TestAngle = 0
B(LI).a = TestAngle
Case 3
TestAngle = B(LI).a - 1
If TestAngle = -1 Then TestAngle = 35
B(LI).a = TestAngle
End Select
Else
B(LI).v = 2
Select Case CDir2
Case 2
TestAngle = B(LI).a + 1
If TestAngle = 36 Then TestAngle = 0
B(LI).a = TestAngle
Case 3
TestAngle = B(LI).a - 1
If TestAngle = -1 Then TestAngle = 35
B(LI).a = TestAngle
End Select
End If
B(LI).coord.X = B(LI).coord.X + (B(LI).v * Sine(B(LI).a))
B(LI).coord.Y = B(LI).coord.Y + (B(LI).v * CoSn(B(LI).a))

Select Case B(LI).frame
Case 1
bmp_rotate GFX.BirdM1, Form1.BirdMask(LI), (B(LI).a * 10) * (Pi / 180)
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BirdMask(LI).hDC, 0, 0, SRCAND
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BlackBird.hDC, 0, 0, SRCPAINT
Case 2
bmp_rotate GFX.BirdM2, Form1.BirdMask(LI), (B(LI).a * 10) * (Pi / 180)
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BirdMask(LI).hDC, 0, 0, SRCAND
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BlackBird.hDC, 0, 0, SRCPAINT
Case 3
bmp_rotate GFX.BirdM3, Form1.BirdMask(LI), (B(LI).a * 10) * (Pi / 180)
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BirdMask(LI).hDC, 0, 0, SRCAND
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BlackBird.hDC, 0, 0, SRCPAINT
Case 4
bmp_rotate GFX.BirdM2, Form1.BirdMask(LI), (B(LI).a * 10) * (Pi / 180)
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BirdMask(LI).hDC, 0, 0, SRCAND
BitBlt Form1.Buffer2.hDC, B(LI).coord.X, B(LI).coord.Y, 29, 29, Form1.BlackBird.hDC, 0, 0, SRCPAINT
End Select

B(LI).frame = B(LI).frame + 1
If B(LI).frame = 5 Then B(LI).frame = 1

End If
Next
End Sub
Sub KillBird(BirdNum As Integer)
B(BirdNum).flying = False

B(BirdNum).Entrypoint.X = RESETSPACE
B(BirdNum).Entrypoint.Y = RESETSPACE
B(BirdNum).Exitpoint.X = RESETSPACE
B(BirdNum).Exitpoint.Y = RESETSPACE
B(BirdNum).coord.X = B(BirdNum).Entrypoint.X
B(BirdNum).coord.Y = B(BirdNum).Entrypoint.Y

B(BirdNum).v = 0
B(BirdNum).frame = 1
B(BirdNum).a = 0
B(BirdNum).Flag = NONE

End Sub
Sub FreakChangeInDirection(): Dim BNum2 As Integer, GC As Integer
Randomize Rnd * 255 'SuperSeed
Grounded: DoEvents
BNum2 = Rnd * (MaxBirds - 1): GC = GC + 1
If GC = 50 Then NewBird
If Not B(BNum).flying Then GoTo Grounded
GC = 0
LR = Rnd * 6
Select Case LR
Case 1
B(BNum2).Exitpoint.X = Rnd * 325
B(BNum2).Exitpoint.Y = 377
Case 2
B(BNum2).Exitpoint.X = 377
B(BNum2).Exitpoint.Y = Rnd * 325
Case 4
B(BNum2).Exitpoint.X = 350
B(BNum2).Exitpoint.Y = Rnd * 377
Case 5
B(BNum2).Exitpoint.X = Rnd * 325
B(BNum2).Exitpoint.Y = 377
Case Else
B(BNum2).Exitpoint.X = Rnd * 325
B(BNum2).Exitpoint.Y = -20
End Select

rn2 = Rnd * 100
If rn2 >= 50 Then
B(BNum2).Flag = NOFOLLOW
Else
B(BNum2).Flag = FOLLOW
End If
End Sub
