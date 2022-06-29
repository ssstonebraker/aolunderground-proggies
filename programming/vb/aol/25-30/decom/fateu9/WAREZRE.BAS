' WAREZRE.FRM
Option Explicit

Sub sub7B98 ()
End Sub

Sub Command3D1_Click ()
Dim l0044 As Variant
Dim l004A As Variant
Dim l004E As String
Dim l0052 As Variant
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006A As Variant
If  warez.Text = "" Then
MsgBox "You must enter in a mail to request!", 16
Exit Sub
End If
If  got.Text = "" Then
MsgBox "You must enter in a trigger word!", 16
Exit Sub
End If
If  Text4.Text = "" Then
MsgBox "You must enter in what to notify!", 16
Exit Sub
End If
If  Text5.Text = "" Then
MsgBox "You must enter in when to mention!", 16
Exit Sub
End If
If  Not IsNumeric(Text5.Text) Then
MsgBox "You must enter in a numeric number!", 16
Exit Sub
End If
l0044 = extfn1018("AOL FRAME25", 0&)
l004A = extfn10F8(l0044, "Welcome")
l004E$ = String(30, 0)
l0052 = extfnFA8(l004A, l004E$, 250)
If  l0052 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  warez.Text = "" Then
MsgBox "You must enter in something to request", 16
Exit Sub
End If
If  got.Text = "" Then
MsgBox "You must enter in a trigger word.", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
DoEvents
l0062 = fnA30(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(82) + Chr(101) + Chr(113) + Chr(117) + Chr(101) + Chr(115) + Chr(116) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l0066 = fnA30(fn330() + "Requested Warez: " + warez.Text + " –-- -›»")
DoEvents
l006A = fnA30(fn330() + "Type """ + got.Text + """ if you have it" + fn368())
DoEvents
Timer1.Enabled = True
Timer2.Enabled = True
Timer1.Interval = 60000&
Timer2.Interval = 1
Command3D1.Caption = "Stop"
Else
Text6.Text = "1"
Timer2.Interval = 0
Timer2.Enabled = False
Timer1.Interval = 0
Timer1.Enabled = False
Command3D1.Caption = "Start"
End If
End Sub

Sub Command3D2_Click ()
List1.Clear
End Sub

Sub Command3D3_Click ()
warezre.WindowState = 1
End Sub

Sub Form_Load ()
Option3D1.Value = True
Text6.Text = "1"
Text5.Text = "1"
extsub1050 warezre.hWnd, -1, 0, 0, 0, 0, 3
If  warezre.Caption <> fnC98() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub HScroll1_Change ()
Dim l0092 As Variant
Text5.Text = Trim(Str(l0092))
End Sub

Sub List1_Click ()
Dim l0096 As Variant
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B6 As Variant
l0096 = extfn1018("AOL FRAME25", 0&)
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
l009E = fn5D0("Mem&bers", "Send an Instant Message")
13 :
DoEvents
l00A2 = extfn10F8(l0096, "Send Instant Message")
If  l00A2 = 0 Then GoTo 13
l00A8 = extfn10C0(l00A2, "_AOL_EDIT")
l00AC = List1.ListIndex
l00B0 = List1.List(l00AC)
l00B6 = extfn11A0(l00A8, 12, 0, l00B0)
End Sub

Sub List1_DblClick ()
List1.RemoveItem List1.ListIndex
End Sub

Sub Timer1_Timer ()
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
If  Timer1.Interval = 0 Then GoTo 40
If  Int(Text5.Text) = Int(Text6.Text) Then
DoEvents
l00BA = fnA30(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(82) + Chr(101) + Chr(113) + Chr(117) + Chr(101) + Chr(115) + Chr(116) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l00BE = fnA30(fn330() + "Requested Warez: " + warez.Text + " –-- -›»")
DoEvents
l00C2 = fnA30(fn330() + "Type """ + got.Text + """ if you have it" + fn368())
DoEvents
Text6.Text = "1"
Else
Text6.Text = Int(Text6.Text) + 1
End If
40 :
End Sub

Sub Timer2_Timer ()
Dim l00D0 As Variant
Text1.Text = fn170()
l00D0 = fn8E0()
End Sub

Sub sub9980 ()
Unload warezre
End Sub
