Attribute VB_Name = "Movin"
'Hey Whatup,
'This is MegaSouL and I wanted to make this bas becuase
'I know so many things to do with a form and
'And I want to put this for the people that dont
'know how to fade a form
'*I am not saying thier dum, but they people should learn
'More About it*
'On this bas I put how to exit a form in any directions
'And I put on a couple of good things with a label
' *Useful for Greetz
'Well Enjoy
'And Thanx For Usin dis Bas

'Peace
'MegaSouL




' Special Thanx Goes To My *MOM & DAD*



' Do Not Edit This bas
'LoL
Declare Sub ReleaseCapture Lib "user32" ()
Public Sub Form_Move(TheForm As Form)
'Put this Function in the Mouse down
'It will help a lot
       ReleaseCapture
       Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
       
End Sub

Sub BigLabel(lbl As Label)
'this could be useful if your making a cool
'greetz
lbl.FontSize = 65
Call Pause(0.1)
lbl.FontSize = 63
Call Pause(0.1)
lbl.FontSize = 61
Call Pause(0.1)
lbl.FontSize = 59
Call Pause(0.1)
lbl.FontSize = 57
Call Pause(0.1)
lbl.FontSize = 55
Call Pause(0.1)
lbl.FontSize = 53
Call Pause(0.1)
lbl.FontSize = 51
Call Pause(0.1)
lbl.FontSize = 49
Call Pause(0.1)
lbl.FontSize = 47
Call Pause(0.1)
lbl.FontSize = 45
Call Pause(0.1)
lbl.FontSize = 43
Call Pause(0.1)
lbl.FontSize = 41
Call Pause(0.1)
lbl.FontSize = 39
Call Pause(0.1)
lbl.FontSize = 37
Call Pause(0.1)
lbl.FontSize = 35
Call Pause(0.1)
lbl.FontSize = 33
Call Pause(0.1)
lbl.FontSize = 31
Call Pause(0.1)
lbl.FontSize = 29
Call Pause(0.1)
lbl.FontSize = 27
Call Pause(0.1)
lbl.FontSize = 25
Call Pause(0.1)
lbl.FontSize = 23
Call Pause(0.1)
lbl.FontSize = 21
Call Pause(0.1)
lbl.FontSize = 19
Call Pause(0.1)
lbl.FontSize = 17
Call Pause(0.1)
lbl.FontSize = 15
Call Pause(0.1)
lbl.FontSize = 13
Call Pause(0.1)
lbl.FontSize = 11
Call Pause(0.1)
lbl.FontSize = 9
Call Pause(0.1)
lbl.FontSize = 7
Call Pause(0.1)
lbl.FontSize = 5
Call Pause(0.1)
lbl.FontSize = 3
Call Pause(0.1)
lbl.FontSize = 1
End Sub

Sub CenterForm(Frm As Form)
Frm.Left = 3800
Frm.Top = 3000
End Sub

Sub Centerformtop(Frm As Form)
Frm.Left = 3800
Frm.Top = 0
End Sub

Sub DanceAndUpRoll(Frm As Form)
Frm.Visible = True
Frm.Enabled = True
Frm.Left = "-30"
Frm.Top = "1170"
Frm.Height = "0"
Pause (0.1)
Frm.Left = "30"
Frm.Top = "1250"
Frm.Height = "100"
Pause (0.1)
Frm.Left = "130"
Frm.Top = "1300"
Frm.Height = "150"
Pause (0.1)
Frm.Left = "200"
Frm.Top = "1350"
Frm.Height = "350"
Pause (0.1)
Frm.Left = "400"
Frm.Top = "1400"
Frm.Height = "400"
Pause (0.1)
Frm.Left = "430"
Frm.Top = "1450"
Frm.Height = "450"
Pause (0.1)
Frm.Left = "530"
Frm.Top = "1500"
Frm.Height = "500"
Pause (0.1)
Frm.Left = "630"
Frm.Top = "1550"
Frm.Height = "550"
Pause (0.1)
Frm.Left = "830"
Frm.Top = "1600"
Frm.Height = "600"
Pause (0.1)
Frm.Top = "1650"
Frm.Left = "1030"
Frm.Height = "650"
Pause (0.1)
Frm.Top = "1700"
Frm.Left = "1230"
Frm.Height = "700"
Pause (0.1)
Frm.Top = "1750"
Frm.Left = "1430"
Frm.Height = "750"
Pause (0.1)
Frm.Top = "1800"
Frm.Left = "1630"
Frm.Height = "800"
Pause (0.1)
Frm.Top = "1850"
Frm.Left = "1830"
Frm.Height = "850"
Pause (0.1)
Frm.Top = "1900"
Frm.Left = "2030"
Frm.Height = "900"
Pause (0.1)
Frm.Top = "1910"
Frm.Left = "2230"
Frm.Height = "950"
Pause (0.1)
Frm.Top = "1920"
Frm.Left = "2430"
Frm.Height = "1000"
Pause (0.1)
Frm.Top = "1930"
Frm.Left = "2630"
Frm.Height = "1020"
Pause (0.1)
Frm.Top = "1940"
Frm.Left = "2830"
Frm.Height = "1040"
Pause (0.1)
Frm.Top = "1950"
Frm.Left = "3030"
Frm.Height = "1060"
Pause (0.1)
Frm.Height = "1080"
Pause (0.1)
Frm.Enabled = True
End Sub

Sub FormExitCool(m As Form)

'  This makes a form dance across the screen
'Then Exits
m.Left = 5
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 2000
Pause (0.1)
m.Left = 3000
Pause (0.1)
m.Left = 4000
Pause (0.1)
m.Left = 5000
Pause (0.1)
m.Left = 4000
Pause (0.1)
m.Left = 3000
Pause (0.1)
m.Left = 2000
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 5
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 2000
End
End Sub

Sub FormExit_Down(Frm As Form)
'The Form Moves Downward on the screen trhe Exits
Do
Frm.Top = Trim(Str(Int(Frm.Top) + 300))
DoEvents
Loop Until Frm.Top > 7200
If Frm.Top > 7200 Then End
End Sub

Sub FormExit_Left(Frm As Form)
'the form files left off the screen, and
'ends the program
Do
Frm.Left = Trim(Str(Int(Frm.Left) - 300))
DoEvents
Loop Until Frm.Left < -6300
If Frm.Left < -6300 Then End
End Sub

Sub FormExit_right(Frm As Form)
'The Form moves to your right side and exits
Do
Frm.Left = Trim(Str(Int(Frm.Left) + 300))
DoEvents
Loop Until Frm.Left > 9600
If Frm.Left > 9600 Then End
End Sub

Sub FormExit_up(winform As Form)
'form moves upward then eixts
Do
Frm.Top = Trim(Str(Int(Frm.Top) - 300))
DoEvents
Loop Until Frm.Top < -4500
If Frm.Top < -4500 Then End
End Sub

Sub Formheight(Frm As Form)
Frm.Height = 5
Pause (0.1)
Frm.Height = 400
Pause (0.1)
Frm.Height = 700
Pause (0.1)
Frm.Height = 1000
Pause (0.1)
Frm.Height = 2000
Pause (0.1)
Frm.Height = 3000
Pause (0.1)
Frm.Height = 4000
Pause (0.1)
Frm.Height = 5000
Pause (0.1)
Frm.Height = 4000
Pause (0.1)
Frm.Height = 3000
Pause (0.1)
Frm.Height = 2000
Pause (0.1)
Frm.Height = 1000
Pause (0.1)
Frm.Height = 700
Pause (0.1)
Frm.Height = 400
Pause (0.1)
Frm.Height = 5
Pause (0.1)
Frm.Height = 400
Pause (0.1)
Frm.Height = 700
Pause (0.1)
Frm.Height = 1000
Pause (0.1)
Frm.Height = 2000

End Sub

Sub Formhieghtandwidth(m As Form)

m.Height = 5
m.Width = 5
Pause (0.1)
m.Height = 400
m.Width = 400
Pause (0.1)
m.Height = 700
m.Width = 700
Pause (0.1)
m.Height = 1000
m.Width = 1000
Pause (0.1)
m.Height = 2000
m.Width = 2000
Pause (0.1)
m.Height = 3000
m.Width = 3000
Pause (0.1)
m.Height = 4000
m.Width = 4000
Pause (0.1)
m.Height = 5000
m.Width = 5000
Pause (0.1)
m.Height = 4000
m.Width = 4000
Pause (0.1)
m.Height = 3000
m.Width = 3000
Pause (0.1)
m.Height = 2000
m.Width = 2000
Pause (0.1)
m.Height = 1000
m.Width = 1000
Pause (0.1)
m.Height = 700
m.Width = 700
Pause (0.1)
m.Height = 400
m.Width = 400
Pause (0.1)
m.Height = 5
m.Width = 5
Pause (0.1)
m.Height = 400
Pause (0.1)
m.Height = 700
Pause (0.1)
m.Height = 1000
Pause (0.1)
m.Height = 2000

End Sub

Sub Formwidth(Frm As Form)

Frm.Width = 5
Pause (0.1)
Frm.Width = 400
Pause (0.1)
Frm.Width = 700
Pause (0.1)
Frm.Width = 1000
Pause (0.1)
Frm.Width = 2000
Pause (0.1)
Frm.Width = 3000
Pause (0.1)
Frm.Width = 4000
Pause (0.1)
Frm.Width = 5000
Pause (0.1)
Frm.Width = 4000
Pause (0.1)
Frm.Width = 3000
Pause (0.1)
Frm.Width = 2000
Pause (0.1)
Frm.Width = 1000
Pause (0.1)
Frm.Width = 700
Pause (0.1)
Frm.Width = 400
Pause (0.1)
Frm.Width = 5
Pause (0.1)
Frm.Width = 400
Pause (0.1)
Frm.Width = 700
Pause (0.1)
Frm.Width = 1000
Pause (0.1)
Frm.Width = 2190

End Sub

Function Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Function

Sub textmove(l As Label)
'U can change this to no matter what
l.Caption = " Pepsi v4 By CpRide And Dc"
Call Pause(0.1)
l.Caption = " Pepsi v4 By CpRide and d"
Call Pause(0.1)
l.Caption = "  Pepsi v4 By CpRide and "
Call Pause(0.1)
l.Caption = "   Pepsi v4 By CpRide and"
Call Pause(0.1)
l.Caption = "    Pepsi v4 By CpRide an"
Call Pause(0.1)
l.Caption = "     Pepsi v4 By CpRide a"
Call Pause(0.1)
l.Caption = "      Pepsi v4 By CpRide "
Call Pause(0.1)
l.Caption = "       Pepsi v4 By CpRide"
Call Pause(0.1)
l.Caption = "        Pepsi v4 By CpRid"
Call Pause(0.1)
l.Caption = "         Pepsi v4 By CpRi"
Call Pause(0.1)
l.Caption = "          Pepsi v4 By CpR"
Call Pause(0.1)
l.Caption = "           Pepsi v4 By Cp"
Call Pause(0.1)
l.Caption = "            Pepsi v4 By C"
Call Pause(0.1)
l.Caption = "             Pepsi v4 By "
Call Pause(0.1)
l.Caption = "              Pepsi v4 By"
Call Pause(0.1)
l.Caption = "               Pepsi v4 B"
Call Pause(0.1)
l.Caption = "                Pepsi v4 "
Call Pause(0.1)
l.Caption = "                 Pepsi v4"
Call Pause(0.1)
l.Caption = "                  Pepsi v"
Call Pause(0.1)
l.Caption = "                   Pepsi "
Call Pause(0.1)
l.Caption = "                    Pepsi"
Call Pause(0.1)
l.Caption = "                     Peps"
Call Pause(0.1)
l.Caption = "                      Pep"
Call Pause(0.1)
l.Caption = "                       Pe"
Call Pause(0.1)
l.Caption = "                        P"
Call Pause(0.1)
l.Caption = "                         "
End Sub
Sub FormFlash(Frm As Form)
'This make a form flash into colorful colors
'Kinda Like the First page in MacroShop3
Do
Frm.Show
Frm.BackColor = &H0&
Pause (".1")
Frm.BackColor = &HFF&
Pause (".1")
Frm.BackColor = &HFF0000
Pause (".1")
Frm.BackColor = &HFF00&
Pause (".1")
Frm.BackColor = &H8080FF
Pause (".1")
Frm.BackColor = &HFFFF00
Pause (".1")
Frm.BackColor = &H80FF&
Pause (".1")
Frm.BackColor = &HC0C0C0
Loop
End Sub
Sub MidBlueBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(0, 0, 255 - i), B
Next i
End Sub

Sub MidRedBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(255 - i, 0, 0), B
Next i
End Sub

Sub MidGreenBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(0, 255 - i, 0), B
Next i
End Sub

Sub MidCyanBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(0, 255 - i, 255 - i), B
Next i
End Sub

Sub MidYellowBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(255 - i, 255 - i, 0), B
Next i
End Sub

Sub MidPurpleBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(255 - i, 0, 255 - i), B
Next i
End Sub

Sub MidGrayBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (0, i)-(Screen.Width, 255 - i), RGB(255 - i, 255 - i, 255 - i), B
Next i
End Sub

Sub RedBuild(Frm As Form)
'This one is hard to explain.
'it builds up from the bottom left of your
'form and fades out
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(255 - i, 0, 0), BF
Next i
End Sub

Sub GreenBuild(Frm As Form)
'This one is hard to explain.
'it builds up from the bottom left of your
'form and fades out
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(0, 255 - i, 0), BF
Next i
End Sub

Sub BlueBuild(Frm As Form)
'This one is hard to explain.
'it builds up from the bottom left of your
'form and fades out
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(0, 0, 255 - i), BF
Next i
End Sub

Sub YellowBuild(Frm As Form)
'Fades into yellow from the bottom of your form
'Looks like a fire ball
'Sweet like hell
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(255 - i, 255 - i, 0), BF
Next i
End Sub

Sub CyanBuild(Frm As Form)
'This one is hard to explain.
'it builds up from the bottom left of your
'form and fades out
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(0, 255 - i, 255 - i), BF
Next i
End Sub

Sub PurpleBuild(Frm As Form)
'This one is hard to explain.
'it builds up from the bottom left of your
'form and fades out
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(255 - i, 0, 255 - i), BF
Next i
End Sub

Sub GrayBuild(Frm As Form)
'This one is hard to explain.
'it builds up from the bottom left of your
'form and fades out
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(i, 255 - i), RGB(255 - i, 255 - i, 255 - i), BF
Next i
End Sub

Sub RedBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(255 - i, 0, 0), B
Next i
End Sub

Sub GreenBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(0, 255 - i, 0), B
Next i
End Sub

Sub BlueBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(0, 0, 255 - i), B
Next i
End Sub

Sub YellowBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(255 - i, 255 - i, 0), B
Next i
End Sub

Sub CyanBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(0, 255 - i, 255 - i), B
Next i
End Sub

Sub PurpleBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(255 - i, 0, 255 - i), B
Next i
End Sub

Sub GrayBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i - 255, 255)-(Frm.Width + i, i), RGB(255 - i, 255 - i, 255 - i), B
Next i
End Sub

Sub HalfRedBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(255 - i, 0, 0), B
Next i
End Sub

Sub HalfGreenBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(0, 255 - i, 0), B
Next i
End Sub

Sub HalfBlueBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(0, 0, 255 - i), B
Next i
End Sub

Sub HalfYellowBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(255 - i, 255 - i, 0), B
Next i
End Sub

Sub HalfCyanBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(0, 255 - i, 255 - i), B
Next i
End Sub

Sub HalfPurpleBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(255 - i, 0, 255 - i), B
Next i
End Sub

Sub HalfGrayBlack(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (255, i - 255)-(i, Frm.Width + i), RGB(255 - i, 255 - i, 255 - i), B
Next i
End Sub

Sub RedBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(255 - i, 0, 0), B
Next i
End Sub

Sub GreenBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(0, 255 - i, 0), B
Next i
End Sub

Sub BlueBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(0, 0, 255 - i), B
Next i
End Sub

Sub YellowBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(255 - i, 255 - i, 0), B
Next i
End Sub

Sub PurpleBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(255 - i, 0, 255 - i), B
Next i
End Sub

Sub CyanBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(0, 255 - i, 255 - i), B
Next i
End Sub

Sub GrayBlackDown(Frm As Form)
With Frm
.DrawStyle = vbInsideSolid
.DrawMode = vbCopyPen
.ScaleMode = vbPixels
.DrawWidth = 2
.ScaleHeight = 256
End With
For i = 0 To 255
Frm.Line (i, 255)-(Frm.Width + i, i), RGB(255 - i, 255 - i, 255 - i), B
Next i
End Sub
Sub FormExit2(Frm As Form)
'Exits A Form Up And Down
Do
        DoEvents
  Frm.Top = Trim(Str(Int(Frm.Top) + 300))
        Loop Until Frm.Top > 7200
  Do
        DoEvents
  Frm.Top = Trim(Str(Int(Frm.Top) - 300))
        Loop Until Frm.Top < -Frm.Width
    End
End Sub
Sub dancewitme(m As Form)

'  This makes a form dance across the screen
m.Left = 5
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 2000
Pause (0.1)
m.Left = 3000
Pause (0.1)
m.Left = 4000
Pause (0.1)
m.Left = 5000
Pause (0.1)
m.Left = 4000
Pause (0.1)
m.Left = 3000
Pause (0.1)
m.Left = 2000
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 5
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 2000

End Sub
Sub Dancewitme2(Frm As Form)
'Dancing up and down
Do
        DoEvents
  Frm.Top = Trim(Str(Int(Frm.Top) + 300))
        Loop Until Frm.Top > 7200
  Do
        DoEvents
  Frm.Top = Trim(Str(Int(Frm.Top) - 300))
        Loop Until Frm.Top < -Frm.Width
    End
End Sub
