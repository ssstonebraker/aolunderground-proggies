' ENCRYPTC.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l002A As Variant
Dim l002E As Variant
If  Text2.Text = "" Then GoTo 10
For l0020 = 0 To List1.ListCount - 1
l002A = List1.List(l0020)
l002E = InStr(1, l002A, Text2.Text, 1)
If  l002E Then
If  Len(Text2.Text) = Len(l002A) Then
GoTo 10
End If
End If
Next l0020
List1.AddItem Text2.Text
10 :
Text2.Text = ""
End Sub

Sub Command3D2_Click ()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Sub Command3D3_Click ()
List1.Clear
End Sub

Sub Command3D4_Click ()
Dim l003E As Variant
Dim l0042 As Variant
Dim l0046 As Variant
Dim l004A As Variant
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0068 As Variant
If  Text3.Text = "" Then
MsgBox "You must have a keyword!", 16
Exit Sub
End If
If  Text4.Text = "" Then
Exit Sub
End If
l003E = 1
For l0042 = 1 To Len(Text4.Text)
l0046 = Mid(Text4.Text, l0042, 1)
l004A = Asc(l0046)
l004E = Mid(Text3.Text, l003E, 1)
l003E = l003E + 1
If  l003E > Len(Text3.Text) Then l003E = 1
l0052 = Asc(l004E)
l0056 = l004A + l0052
If  l0056 > 255 Then
l005A = l0056 - 255
l005E = Chr(l005A)
Else
l005E = Chr(l0056)
End If
l0062 = l0062 + l005E
DoEvents
Next l0042
l0068 = fnA68(l0062)
Text4.Text = ""
End Sub

Sub Command3D5_Click ()
Dim l006E As Variant
Dim l0074 As Variant
Dim l007A As Variant
Dim l0080 As Variant
Dim l0086 As Variant
Dim l0094 As Variant
If  Text3.Text = "" Then
MsgBox "You must have a keyword!", 16
Exit Sub
End If
l006E = extfn1018("AOL FRAME25", 0&)
l0074 = extfn10F8(l006E, "AOL Live!")
l007A = extfn1088(l0074)
l0080 = extfn10C0(l007A, "_AOL_EDIT")
If  l0080 = 0 Then
MsgBox "You must be in a chat room!", 16
Exit Sub
End If
l0086 = extfn1758(l006E, 2)
If  Command3D5.Caption = "Start" Then
l0094 = fnA68(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(69) + Chr(110) + Chr(99) + Chr(114) + Chr(121) + Chr(112) + Chr(116) + Chr(105) + Chr(111) + Chr(110) + Chr(32) + Chr(67) + Chr(104) + Chr(97) + Chr(116) + Chr(32) + Chr(69) + Chr(110) + Chr(97) + Chr(98) + Chr(108) + Chr(101) + Chr(100) + fn368())
Text1.Text = "Fate Host:" + Chr(9) + "«‹- -– Welcome To Encryption Chat –-- -›»"
DoEvents
Command3D5.Caption = "Stop"
Timer1.Enabled = True
Timer1.Interval = 1
Else
l0086 = extfn1758(l006E, 4)
l0094 = fnA68(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(69) + Chr(110) + Chr(99) + Chr(114) + Chr(121) + Chr(112) + Chr(116) + Chr(105) + Chr(111) + Chr(110) + Chr(32) + Chr(67) + Chr(104) + Chr(97) + Chr(116) + Chr(32) + Chr(68) + Chr(105) + Chr(115) + Chr(97) + Chr(98) + Chr(108) + Chr(101) + Chr(100) + fn368())
DoEvents
Command3D5.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
End If
End Sub

Sub Command3D6_Click ()
Dim l00A8 As Variant
help1.Show
For l00A8 = 0 To help1.List1.ListCount - 1
If  help1.List1.List(l00A8) = "Encryption Chat" Then
help1.List1.Selected(l00A8) = True
Exit Sub
End If
Next l00A8
End Sub

Sub Command3D7_Click ()
encryptchat.WindowState = 1
End Sub

Sub Form_Load ()
Dim l00B4 As Variant
If  encryptchat.Caption <> fn2F8() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
extsub1050 encryptchat.hWnd, -1, 0, 0, 0, 0, 3
l00B4 = extfn1210("encrypt.wav", 1)
End Sub

Sub Text1_GotFocus ()
Text1.SelStart = Len(Text1.Text)
Text4.SetFocus
End Sub

Sub Text2_KeyPress (KeyAscii As Integer)
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
l00BA = Chr(KeyAscii)
If  l00BA = Chr(13) Then
If  Text2.Text = "" Then GoTo 20
For l00BA = 0 To List1.ListCount - 1
l00BE = List1.List(l00BA)
l00C2 = InStr(1, l00BE, Text2.Text, 1)
If  l00C2 Then
If  Len(Text2.Text) = Len(l00BE) Then
GoTo 20
End If
End If
Next l00BA
List1.AddItem Text2.Text
20 :
Text2.Text = ""
End If
End Sub

Sub Text4_KeyPress (KeyAscii As Integer)
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00E8 As Variant
Dim l00EC As Variant
Dim l00F0 As Variant
l00C8 = Chr(KeyAscii)
If  l00C8 = Chr(13) Then
Text1.SelStart = Len(Text1.Text)
If  Text3.Text = "" Then
MsgBox "You must have a keyword!", 16
Exit Sub
End If
If  Text4.Text = "" Then
Exit Sub
End If
l00CC = 1
For l00C8 = 1 To Len(Text4.Text)
l00D0 = Mid(Text4.Text, l00C8, 1)
l00D4 = Asc(l00D0)
l00D8 = Mid(Text3.Text, l00CC, 1)
l00CC = l00CC + 1
If  l00CC > Len(Text3.Text) Then l00CC = 1
l00DC = Asc(l00D8)
l00E0 = l00D4 + l00DC
If  l00E0 > 255 Then
l00E4 = l00E0 - 255
l00E8 = Chr(l00E4)
Else
l00E8 = Chr(l00E0)
End If
l00EC = l00EC + l00E8
DoEvents
Next l00C8
l00F0 = fnA68(l00EC)
Text4.Text = ""
End If
End Sub

Sub Timer1_Timer ()
Dim l00FE As Variant
Text5.Text = fn170()
l00FE = fn2C0()
End Sub
