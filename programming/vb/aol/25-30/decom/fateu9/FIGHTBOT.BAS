' FIGHTBOT.FRM
Option Explicit

Sub Check3D1_Click (Value As Integer)
If  Check3D1.Value = True Then
Text3.Enabled = True
Command3D2.Enabled = True
Else
Text3.Enabled = False
Command3D2.Enabled = False
End If
End Sub

Sub Command3D1_Click ()
If  Text1.Text = "" Then
MsgBox "Please enter the name of the first fighter!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "Please enter the name of the second fighter!", 16
Exit Sub
End If
fightbot.Hide
fight2.Show
Timer1.Enabled = True
Timer1.Interval = 1
fight2.Frame3D1.Caption = Text1.Text
fight2.Frame3D2.Caption = Text2.Text
End Sub

Sub Command3D2_Click ()
Dim l004A As Variant
Dim l004E As Variant
Dim l0052 As Variant
If  Command3D2.Caption = "Start" Then
If  Text1.Text = "" And Text2.Text > "" Then
l004A = fnA68("«‹- -– Fate Ultra Fight Bot –-- -›»")
l004E = fnA68("«‹- -– " + Text2.Text + " needs a challenger! –-- -›»")
l0052 = fnA68("«‹- -– To challenge type" + Chr(34) + Text3.Text + Chr(34) + " –-- -›»")
End If
If  Text2.Text = "" And Text1.Text > "" Then
l004A = fnA68("«‹- -– Fate Ultra Fight Bot –-- -›»")
l004E = fnA68("«‹- -– " + Text1.Text + " needs a challenger! –-- -›»")
l0052 = fnA68("«‹- -– To challenge type" + Chr(34) + Text3.Text + Chr(34) + " –-- -›»")
End If
If  Text2.Text = "" And Text1.Text = "" Then
l004A = fnA68("«‹- -– Fate Ultra Fight Bot –-- -›»")
l004E = fnA68("«‹- -– Two challengers are needed! –-- -›»")
l0052 = fnA68("«‹- -– To challenge type" + Chr(34) + Text3.Text + Chr(34) + " –-- -›»")
End If
Command3D2.Caption = "Stop"
Timer2.Enabled = True
Timer2.Interval = 1
Else
Command3D2.Caption = "Start"
Timer2.Interval = 0
Timer2.Enabled = False
End If
End Sub

Sub Form_Load ()
extsub1050 fightbot.hWnd, -1, 0, 0, 0, 0, 3
Combo1.AddItem "1"
Combo1.AddItem "3"
Combo1.AddItem "5"
Combo1.AddItem "7"
Combo1.AddItem "9"
Combo1.ListIndex = 1
Combo2.AddItem "Slow"
Combo2.AddItem "Medium"
Combo2.AddItem "Fast"
Combo2.ListIndex = 1
End Sub

Sub Timer1_Timer ()
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
l0068 = Int(Combo1.Text)
l006C = 1
l0070 = 0
l0074 = 0
l0078 = 0
l007C = 0
l0080 = fnA68("«‹- -– Fate Ultra Fight Bot Enabled –-- -›»")
l0086 = fn13A0(.6#)
DoEvents
l0080 = fnA68(" ")
l0086 = fn13A0(.6#)
DoEvents
If  Combo1.Text = "1" Then
l008A = ""
Else
l008A = "s"
End If
l0080 = fnA68("«‹- -– Todays Matchup: (" + Combo1.Text + " Round" + l008A + ") –-- -›»")
l0086 = fn13A0(.6#)
l0080 = fnA68(" ")
l0086 = fn13A0(.6#)
DoEvents
l0080 = fnA68("«‹- -– " + Text1.Text + " Vs. " + Text2.Text + " –-- -›»")
l0086 = fn13A0(.6#)
DoEvents
l0080 = fnA68(" ")
l0086 = fn13A0(.6#)
DoEvents
L246E8:
l0080 = fnA68("«‹- -–      - Round " + Trim(Str(l006C)) + " -      –-- -›»")
l0086 = fn13A0(.6#)
fight2.Label7.Caption = "Round: " + Trim(Str(l006C))
L247A6:
DoEvents
Randomize
l008E = Int(14 * Rnd)
Select Case l008E
Case 0
l0092 = " kicks"
l0070 = l0070 + 1
Case 1
l0092 = " jabs"
l0070 = l0070 + 1
Case 2
l0092 = " uppercuts"
l0070 = l0070 + 3
Case 3
l0092 = " elbos"
l0070 = l0070 + 2
Case 4
l0092 = " socks"
l0070 = l0070 + 1
Case 5
l0092 = " jumpkicks"
l0070 = l0070 + 2
Case 6
l0092 = " trips"
l0070 = l0070 + 2
Case 7
l0092 = " sweeps"
l0070 = l0070 + 2
Case 8
l0092 = " knocks"
l0070 = l0070 + 3
Case 9
l0092 = " COMBO"
Case 10
l0092 = " hooks"
l0070 = l0070 + 1
Case 11
l0092 = " headlocks"
l0070 = l0070 + 1
Case 12
l0092 = " pokes"
l0070 = l0070 + 1
Case 13
l0092 = " spin kicks"
l0070 = l0070 + 4
End Select
Randomize
l0096 = Int(10 * Rnd)
Select Case l0096
Case 0
l009A = " in the head"
l0070 = l0070 + 5
Case 1
l009A = " in the stomach"
l0070 = l0070 + 2
Case 2
l009A = " in the groin"
l0070 = l0070 + 7
Case 3
l009A = " in the face"
l0070 = l0070 + 3
Case 4
l009A = " in the legs"
l0070 = l0070 + 2
Case 5
l009A = " in the knees"
l0070 = l0070 + 2
Case 6
l009A = " in the kneck"
l0070 = l0070 + 2
Case 7
l009A = " in the back"
Case 8
l009A = " in the eyes"
Case 9
l009A = " in the thighs"
End Select
Randomize
l009E = Int(2 * Rnd)
Select Case l009E
Case 0
If  l0092 = " headlocks" Then
l00A2 = l0092
l00A6 = ""
l0070 = 1
GoTo L24ED0
End If
If  l0092 = " trips" Then
l00A2 = l0092
l00A6 = ""
l0070 = 2
GoTo L24ED0
End If
If  l0092 = " COMBO" Then
l00A2 = " does a combo on"
l00A6 = ""
l0070 = 20
GoTo L24ED0
End If
l00A2 = l0092
l00A6 = l009A
L24ED0:
l00AA = fn13A0(2)
l00AE = fnA68("«‹- -– " + Text1.Text + l00A2 + " " + Text2.Text + l00A6 + " –-- -›»")
fight2.Label2.Caption = Trim(Str(Int(fight2.Label2.Caption) + Int(l0070 * 300)))
If  Combo2.Text = "Slow" Then
l00B2 = l0070 * 30
End If
If  Combo2.Text = "Medium" Then
l00B2 = l0070 * 60
End If
If  Combo2.Text = "Fast" Then
l00B2 = l0070 * 100
End If
l007C = l007C + l00B2
fight2.Picture2.Line (x,y)(0, 0)-(l007C, 500), QBColor(4), BF
Case 1
If  l0092 = " headlocks" Then
l00A2 = l0092
l00A6 = ""
l0070 = 1
GoTo L24ED0
End If
If  l0092 = " trips" Then
l00A2 = l0092
l00A6 = ""
l0070 = 2
GoTo L2519E
End If
If  l0092 = " COMBO" Then
l00A2 = " does a combo on"
l00A6 = ""
l0070 = 20
GoTo L24ED0
End If
l00A2 = l0092
l00A6 = l009A
L2519E:
l00B6 = fnA68("«‹- -– " + Text2.Text + l00A2 + " " + Text1.Text + l00A6 + " –-- -›»")
l00AA = fn13A0(2)
fight2.Label4.Caption = Trim(Str(Int(fight2.Label4.Caption + Int(l0070 * 300))))
If  Combo2.Text = "Slow" Then
l00B2 = l0070 * 30
End If
If  Combo2.Text = "Medium" Then
l00B2 = l0070 * 60
End If
If  Combo2.Text = "Fast" Then
l00B2 = l0070 * 100
End If
l0078 = l0078 + l00B2
fight2.Picture1.Line (x,y)(0, 0)-(l0078, 500), QBColor(4), BF
End Select
If  l0078 > 1703 Then
l00BA = fnA68("«‹- -– " + Text1.Text + " loses the first round! –-- -›»")
fight2.Picture1.Cls
fight2.Picture2.Cls
l0078 = 0
l007C = 0
l00BE = Mid(fight2.Label8.Caption, 7, Len(fight2.Label8.Caption))
fight2.Label8.Caption = "Lost: " + Trim(Str(Int(l00BE) + 1))
l00BE = Mid(fight2.Label11.Caption, 6, Len(fight2.Label11.Caption))
fight2.Label11.Caption = "Won: " + Trim(Str(Int(l00BE) + 1))
l0068 = l0068 - 1
If  l0068 = 0 Then
l00C2 = Mid(fight2.Label10.Caption, 6, Len(fight2.Label10.Caption))
l00C6 = Mid(fight2.Label11.Caption, 6, Len(fight2.Label11.Caption))
If  l00C2 > l00C6 Then
l00CA = fnA68("«‹- -– " + Text1.Text + " is the Winner! –-- -›»")
l00CE = fnA68("«‹- -– Fate Ultra Fight Bot Disabled –-- -›»")
Else
l00CA = fnA68("«‹- -– " + Text2.Text + " is the Winner! –-- -›»")
l00CE = fnA68("«‹- -– Fate Ultra Fight Bot Disabled –-- -›»")
End If
GoTo L25A7E
End If
l006C = l006C + 1
GoTo L246E8
End If
If  l007C > 1703 Then
l00BA = fnA68("«‹- -– " + Text2.Text + " loses the first round! –-- -›»")
fight2.Picture1.Cls
fight2.Picture2.Cls
l007C = 0
l0078 = 0
l00BE = Mid(fight2.Label9.Caption, 7, Len(fight2.Label9.Caption))
fight2.Label9.Caption = "Lost: " + Trim(Str(Int(l00BE) + 1))
l00BE = Mid(fight2.Label10.Caption, 6, Len(fight2.Label10.Caption))
fight2.Label10.Caption = "Won: " + Trim(Str(Int(l00BE) + 1))
l0068 = l0068 - 1
If  l0068 = 0 Then
l00C2 = Mid(fight2.Label10.Caption, 6, Len(fight2.Label10.Caption))
l00C6 = Mid(fight2.Label11.Caption, 6, Len(fight2.Label11.Caption))
If  l00C2 > l00C6 Then
l00CA = fnA68("«‹- -– " + Text1.Text + " is the Winner! –-- -›»")
l00CE = fnA68("«‹- -– Fate Ultra Fight Bot Disabled –-- -›»")
Else
l00CA = fnA68("«‹- -– " + Text2.Text + " is the Winner! –-- -›»")
l00CE = fnA68("«‹- -– Fate Ultra Fight Bot Disabled –-- -›»")
End If
GoTo L25A7E
End If
l006C = l006C + 1
GoTo L246E8
End If
l0070 = 0
GoTo L247A6
L25A7E:
fightbot.Show
Unload fight2
Timer1.Interval = 0
Timer1.Enabled = False
End Sub

Sub Timer2_Timer ()
Dim l00DC As Variant
If  Timer2.Interval = 0 Then GoTo 100
Text4.Text = fn170()
l00DC = fn13D8()
100 :
End Sub
