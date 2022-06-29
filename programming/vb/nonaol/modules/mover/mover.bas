'Email To : chillz2000@hotmail.com
'this was made in Vb 5 Pro but works with all versions
'this is for aol 4.o all thought there is not may subs
'about aol 4.0 in here this is just a cool graffics bas
'so far but more will be added

'Peace out

'Chillz

Declare Sub ReleaseCapture Lib "user32" ()

Public Sub Form_Move(theForm As Form)
'WILL HELP YOU MOVE A FORM WITHOUT
'A TITLE BAR, PLACE IN MOUSEDOWN
       ReleaseCapture
       Call SendMessage(theForm.hWnd, &HA1, 2, 0&)
       '
End Sub

Sub bigtext (l As Label)
l.FontSize = 33
Call Pause(.1)
l.FontSize = 31
Call Pause(.1)
l.FontSize = 29
Call Pause(.1)
l.FontSize = 27
Call Pause(.1)
l.FontSize = 25
Call Pause(.1)
l.FontSize = 23
Call Pause(.1)
l.FontSize = 21
Call Pause(.1)
l.FontSize = 19
Call Pause(.1)
l.FontSize = 17
Call Pause(.1)
l.FontSize = 15
Call Pause(.1)
l.FontSize = 13
Call Pause(.1)
l.FontSize = 11
Call Pause(.1)
l.FontSize = 9
Call Pause(.1)
l.FontSize = 7
Call Pause(.1)
l.FontSize = 5
Call Pause(.1)
l.FontSize = 3
Call Pause(.1)
l.FontSize = 1
Call Pause(.2)
l.FontSize = 33
Call Pause(.1)
l.FontSize = 31
Call Pause(.1)
l.FontSize = 29
Call Pause(.1)
l.FontSize = 27
Call Pause(.1)
l.FontSize = 25
Call Pause(.1)
l.FontSize = 23
Call Pause(.1)
l.FontSize = 21
Call Pause(.1)
l.FontSize = 19
Call Pause(.1)
l.FontSize = 17
Call Pause(.1)
l.FontSize = 15
Call Pause(.1)
l.FontSize = 13
Call Pause(.1)
l.FontSize = 11
Call Pause(.1)
l.FontSize = 9
Call Pause(.1)
l.FontSize = 7
Call Pause(.1)
l.FontSize = 5
Call Pause(.1)
l.FontSize = 3
Call Pause(.1)
l.FontSize = 1
End Sub

Sub Centerform (f As Form)
f.Left = 3800
f.Top = 3000
End Sub

Sub Centerformtop (f As Form)
f.Left = 3800
f.Top = 0
End Sub

Sub coolform (e As Form)
e.Visible = True
e.Enabled = True
e.Left = "-30"
e.Top = "1170"
e.Height = "0"
Pause (.1)
e.Left = "30"
e.Top = "1250"
e.Height = "100"
Pause (.1)
e.Left = "130"
e.Top = "1300"
e.Height = "150"
Pause (.1)
e.Left = "200"
e.Top = "1350"
e.Height = "350"
Pause (.1)
e.Left = "400"
e.Top = "1400"
e.Height = "400"
Pause (.1)
e.Left = "430"
e.Top = "1450"
e.Height = "450"
Pause (.1)
e.Left = "530"
e.Top = "1500"
e.Height = "500"
Pause (.1)
e.Left = "630"
e.Top = "1550"
e.Height = "550"
Pause (.1)
e.Left = "830"
e.Top = "1600"
e.Height = "600"
Pause (.1)
e.Top = "1650"
e.Left = "1030"
e.Height = "650"
Pause (.1)
e.Top = "1700"
e.Left = "1230"
e.Height = "700"
Pause (.1)
e.Top = "1750"
e.Left = "1430"
e.Height = "750"
Pause (.1)
e.Top = "1800"
e.Left = "1630"
e.Height = "800"
Pause (.1)
e.Top = "1850"
e.Left = "1830"
e.Height = "850"
Pause (.1)
e.Top = "1900"
e.Left = "2030"
e.Height = "900"
Pause (.1)
e.Top = "1910"
e.Left = "2230"
e.Height = "950"
Pause (.1)
e.Top = "1920"
e.Left = "2430"
e.Height = "1000"
Pause (.1)
e.Top = "1930"
e.Left = "2630"
e.Height = "1020"
Pause (.1)
e.Top = "1940"
e.Left = "2830"
e.Height = "1040"
Pause (.1)
e.Top = "1950"
e.Left = "3030"
e.Height = "1060"
Pause (.1)
e.Height = "1080"
Pause (.1)
e.Enabled = True
End Sub

Sub FormDance (m As Form)

'  This makes a form dance across the screen
m.Left = 5
Pause (.1)
m.Left = 400
Pause (.1)
m.Left = 700
Pause (.1)
m.Left = 1000
Pause (.1)
m.Left = 2000
Pause (.1)
m.Left = 3000
Pause (.1)
m.Left = 4000
Pause (.1)
m.Left = 5000
Pause (.1)
m.Left = 4000
Pause (.1)
m.Left = 3000
Pause (.1)
m.Left = 2000
Pause (.1)
m.Left = 1000
Pause (.1)
m.Left = 700
Pause (.1)
m.Left = 400
Pause (.1)
m.Left = 5
Pause (.1)
m.Left = 400
Pause (.1)
m.Left = 700
Pause (.1)
m.Left = 1000
Pause (.1)
m.Left = 2000

End Sub

Sub FormExit_Down (winform As Form)
'the form files down off the screen, and
'ends the program
Do
winform.Top = Trim(Str(Int(winform.Top) + 300))
DoEvents
Loop Until winform.Top > 7200
If winform.Top > 7200 Then End
End Sub

Sub FormExit_Left (winform As Form)
'the form files left off the screen, and
'ends the program
Do
winform.Left = Trim(Str(Int(winform.Left) - 300))
DoEvents
Loop Until winform.Left < -6300
If winform.Left < -6300 Then End
End Sub

Sub FormExit_right (winform As Form)
'the form files right off the screen, and
'ends the program
Do
winform.Left = Trim(Str(Int(winform.Left) + 300))
DoEvents
Loop Until winform.Left > 9600
If winform.Left > 9600 Then End
End Sub

Sub FormExit_up (winform As Form)
'the form files up off the screen, and
'ends the program
Do
winform.Top = Trim(Str(Int(winform.Top) - 300))
DoEvents
Loop Until winform.Top < -4500
If winform.Top < -4500 Then End
End Sub

Sub Formheight (m As Form)

m.Height = 5
Pause (.1)
m.Height = 400
Pause (.1)
m.Height = 700
Pause (.1)
m.Height = 1000
Pause (.1)
m.Height = 2000
Pause (.1)
m.Height = 3000
Pause (.1)
m.Height = 4000
Pause (.1)
m.Height = 5000
Pause (.1)
m.Height = 4000
Pause (.1)
m.Height = 3000
Pause (.1)
m.Height = 2000
Pause (.1)
m.Height = 1000
Pause (.1)
m.Height = 700
Pause (.1)
m.Height = 400
Pause (.1)
m.Height = 5
Pause (.1)
m.Height = 400
Pause (.1)
m.Height = 700
Pause (.1)
m.Height = 1000
Pause (.1)
m.Height = 2000

End Sub

Sub Formhw (m As Form)

m.Height = 5
m.Width = 5
Pause (.1)
m.Height = 400
m.Width = 400
Pause (.1)
m.Height = 700
m.Width = 700
Pause (.1)
m.Height = 1000
m.Width = 1000
Pause (.1)
m.Height = 2000
m.Width = 2000
Pause (.1)
m.Height = 3000
m.Width = 3000
Pause (.1)
m.Height = 4000
m.Width = 4000
Pause (.1)
m.Height = 5000
m.Width = 5000
Pause (.1)
m.Height = 4000
m.Width = 4000
Pause (.1)
m.Height = 3000
m.Width = 3000
Pause (.1)
m.Height = 2000
m.Width = 2000
Pause (.1)
m.Height = 1000
m.Width = 1000
Pause (.1)
m.Height = 700
m.Width = 700
Pause (.1)
m.Height = 400
m.Width = 400
Pause (.1)
m.Height = 5
m.Width = 5
Pause (.1)
m.Height = 400
Pause (.1)
m.Height = 700
Pause (.1)
m.Height = 1000
Pause (.1)
m.Height = 2000

End Sub

Sub Formwidth (m As Form)

m.Width = 5
Pause (.1)
m.Width = 400
Pause (.1)
m.Width = 700
Pause (.1)
m.Width = 1000
Pause (.1)
m.Width = 2000
Pause (.1)
m.Width = 3000
Pause (.1)
m.Width = 4000
Pause (.1)
m.Width = 5000
Pause (.1)
m.Width = 4000
Pause (.1)
m.Width = 3000
Pause (.1)
m.Width = 2000
Pause (.1)
m.Width = 1000
Pause (.1)
m.Width = 700
Pause (.1)
m.Width = 400
Pause (.1)
m.Width = 5
Pause (.1)
m.Width = 400
Pause (.1)
m.Width = 700
Pause (.1)
m.Width = 1000
Pause (.1)
m.Width = 2190

End Sub

Function Pause (interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Function

Sub sweet (a As Form)
a.Enabled = False
a.Left = "-30"
a.Top = "1170"
a.Height = "0"
Pause (.1)
a.Left = "30"
a.Top = "1250"
a.Height = "100"
Pause (.1)
a.Left = "130"
a.Top = "1300"
a.Height = "200"
Pause (.1)
a.Left = "230"
a.Top = "1350"
a.Height = "300"
Pause (.1)
a.Left = "330"
a.Top = "1400"
a.Height = "400"
Pause (.1)
a.Left = "430"
a.Top = "1450"
a.Height = "500"
Pause (.1)
a.Left = "530"
a.Top = "1500"
a.Height = "600"
Pause (.1)
a.Left = "630"
a.Top = "1550"
a.Height = "700"
Pause (.1)
a.Left = "830"
a.Top = "1600"
a.Height = "800"
Pause (.1)
a.Top = "1650"
a.Left = "1030"
a.Height = "900"
Pause (.1)
a.Top = "1700"
a.Left = "1230"
a.Height = "1000"
Pause (.1)
a.Top = "1750"
a.Left = "1430"
a.Height = "1100"
Pause (.1)
a.Top = "1800"
a.Left = "1630"
a.Height = "1200"
Pause (.1)
a.Top = "1850"
a.Left = "1830"
a.Height = "1300"
Pause (.1)
a.Top = "1900"
a.Left = "2030"
a.Height = "1400"
Pause (.1)
a.Top = "1910"
a.Left = "2230"
a.Height = "1500"
Pause (.1)
a.Top = "1920"
a.Left = "2430"
a.Height = "1600"
Pause (.1)
a.Top = "1930"
a.Left = "2630"
a.Height = "1700"
Pause (.1)
a.Top = "1940"
a.Left = "2830"
a.Height = "1800"
Pause (.1)
a.Top = "1950"
a.Left = "3030"
a.Height = "1900"
Pause (.1)
a.Left = "3130"
a.Height = "2000"
Pause (.1)
a.Left = "3230"
a.Height = "2100"
Pause (.1)
a.Left = "3330"
a.Height = "2200"
Pause (.1)
a.Height = "2300"
Pause (.1)
a.Height = "2360"
a.Enabled = True
End Sub

Sub textmove (l As Label)
l.Caption = "Pepsi v4 By CpRide and dc"
Call Pause(.1)
l.Caption = " Pepsi v4 By CpRide and d"
Call Pause(.1)
l.Caption = "  Pepsi v4 By CpRide and "
Call Pause(.1)
l.Caption = "   Pepsi v4 By CpRide and"
Call Pause(.1)
l.Caption = "    Pepsi v4 By CpRide an"
Call Pause(.1)
l.Caption = "     Pepsi v4 By CpRide a"
Call Pause(.1)
l.Caption = "      Pepsi v4 By CpRide "
Call Pause(.1)
l.Caption = "       Pepsi v4 By CpRide"
Call Pause(.1)
l.Caption = "        Pepsi v4 By CpRid"
Call Pause(.1)
l.Caption = "         Pepsi v4 By CpRi"
Call Pause(.1)
l.Caption = "          Pepsi v4 By CpR"
Call Pause(.1)
l.Caption = "           Pepsi v4 By Cp"
Call Pause(.1)
l.Caption = "            Pepsi v4 By C"
Call Pause(.1)
l.Caption = "             Pepsi v4 By "
Call Pause(.1)
l.Caption = "              Pepsi v4 By"
Call Pause(.1)
l.Caption = "               Pepsi v4 B"
Call Pause(.1)
l.Caption = "                Pepsi v4 "
Call Pause(.1)
l.Caption = "                 Pepsi v4"
Call Pause(.1)
l.Caption = "                  Pepsi v"
Call Pause(.1)
l.Caption = "                   Pepsi "
Call Pause(.1)
l.Caption = "                    Pepsi"
Call Pause(.1)
l.Caption = "                     Peps"
Call Pause(.1)
l.Caption = "                      Pep"
Call Pause(.1)
l.Caption = "                       Pe"
Call Pause(.1)
l.Caption = "                        P"
Call Pause(.1)
l.Caption = "                         "
End Sub

