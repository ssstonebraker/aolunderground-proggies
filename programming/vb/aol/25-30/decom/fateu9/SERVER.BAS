' SERVER.FRM
Option Explicit

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
List1.Clear
Frame3D1.Caption = "Served: 0"
End Sub

Sub Command3D2_Click ()
List2.Clear
Frame3D3.Caption = "Ignore: 0"
End Sub

Sub Command3D3_Click ()
Dim l004C As Variant
Dim l0056 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As String
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
If  server.Caption <> fn14F0() Then Unload server
If  List3.ListCount = 0 Then
MsgBox "You must have a Mail list. Please click on 'Refresh'", 16
Exit Sub
End If
If  Command3D3.Caption = "Start" Then
Command3D3.Caption = "Stop"
Timer1.Enabled = True
Timer1.Interval = 60000&
l004C = extfn1018("AOL FRAME25", 0&)
If  Option3D1.Value = True Then l0056 = "New Mail"
If  Option3D2.Value = True Then l0056 = "Old Mail"
If  Option3D3.Value = True Then l0056 = "Incoming FlashMail"
l0068 = extfn10F8(l004C, l0056)
If  l0068 = 0 Then
MsgBox "Please open your " + l0056 + " before begining."
Exit Sub
End If
l006C = extfn1018("AOL FRAME25", 0&)
l0070 = extfn10F8(l006C, "Welcome")
l0074$ = String(30, 0)
l0078 = extfnFA8(l0070, l0074$, 250)
l007C = Mid(l0074$, 10, 10)
l0080 = InStr(l007C, "!")
If  l0080 Then
l007C = Mid(l007C, 1, l0080 - 1)
End If
l0086 = fnA68("«‹- -– Fate Ultra WareZ Server By MaGuS –-- -›»")
DoEvents
DoEvents
l008A = fnA68("«‹- -– Type " + Chr(34) + "/" + l007C + " Send X" + Chr(34) + " X Is The Index # –-- -›»")
DoEvents
DoEvents
l008E = fnA68("«‹- -– Type " + Chr(34) + "/" + l007C + " Send List" + Chr(34) + " To Get List –-- -›»")
DoEvents
DoEvents
Text1.Text = fn170()
l0086 = fnB10()
Else
DoEvents
l0086 = fnA68("«‹- -– Fate Ultra WareZ Server Canceled –-- -›»")
DoEvents
DoEvents
Command3D3.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
End If
End Sub

Sub Command3D4_Click ()
Dim l009C As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B4 As Variant
Dim l00BC As Variant
Dim l00C2 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As String
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00E8 As Variant
Dim l00EC As Variant
Dim l00F0 As Variant
If  server.Caption <> fn14F0() Then Unload server
If  Option3D1.Value = True Then l009C = "New Mail"
If  Option3D2.Value = True Then l009C = "Old Mail"
If  Option3D3.Value = True Then l009C = "Incoming FlashMail"
pwait.Show
List3.Clear
Text5.Text = ""
l00AA = extfn1018("AOL FRAME25", 0&)
l00AE = extfn10F8(l00AA, l009C)
l00B4 = extfn10C0(l00AE, "_AOL_TREE")
If  l00B4 = 0 Then
If  fn138() = 5 Then
If  Option3D1.Value = True Then l00BC = fn5D0("&Mail", "Read &New Mail")
If  Option3D2.Value = True Then l00BC = fn5D0("&Mail", "Check Mail You've &Read")
If  Option3D3.Value = True Then l00BC = fn5D0("&Mail", "Read &Incoming Mail")
MsgBox "Please Wait For Your Mailbox To Be Full.", 16
Else
If  Option3D1.Value = True Then l00BC = fn5D0("&Mail", "Read &New Mail")
If  Option3D2.Value = True Then l00BC = fn5D0("&Mail", "Check Mail You've &Read")
If  Option3D3.Value = True Then l00BC = fn5D0("&Mail", "Read &Incoming FlashMail")
End If
End If
l00AA = extfn1018("AOL FRAME25", 0&)
l00AE = extfn10F8(l00AA, l009C)
l00C2 = extfn1758(l00AE, 2)
l00B4 = extfn10C0(l00AE, "_AOL_TREE")
l00C8 = extfnFE0(l00B4, 1036, 0, 0)
For l00CC = 0 To l00C8 - 1
DoEvents
l00D0$ = String(255, 0)
l00D4 = extfn11A0(l00B4, 1034, l00CC, l00D0$)
l00D8 = InStr(l00D0$, Chr(9))
l00DC = Mid(l00D0$, l00D8 + 1, Len(l00D0$))
l00E0 = InStr(l00DC, Chr(9))
l00E4 = Mid(l00DC, l00E0 + 1, Len(l00DC))
l00E8 = InStr(l00E4, Chr(0))
l00EC = Mid(l00E4, 1, l00E8 - 1)
List3.AddItem Str(l00CC) + ")" + l00EC
Text5.Text = Text5.Text + Str(l00CC) + ")" + l00EC + Chr(13) + Chr(10)
Next l00CC
l00F0 = extfn1758(l00AE, 0)
Unload pwait
End Sub

Sub Command3D5_Click ()
If  server.Caption <> fn14F0() Then Unload server
server.WindowState = 1
End Sub

Sub Command3D6_Click ()
Dim l00F4 As Variant
l00F4 = List2.ListIndex
On Error Resume Next
List2.RemoveItem l00F4
If  Err Then
MsgBox "Please highlight the person you wish to delete!", 16
Exit Sub
Else
l00F4 = Mid(Frame3D3, 9, Len(Frame3D3))
Frame3D3.Caption = "Ignore: " + Trim(Str(l00F4 - 1))
End If
End Sub

Sub Command3D7_Click ()
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
l00F8 = InputBox("Enter in the name to add:", "Enter in name to add:")
If  l00F8 = "" Then Exit Sub
If  Len(l00F8) > 10 Then
MsgBox "Invalid screen name!", 16
Exit Sub
End If
For l00FC = 0 To List2.ListCount - 1
l0100 = InStr(1, List2.List(l00FC), l00F8, 1)
If  l0100 Then
If  Len(List2.List(l00FC)) = Len(l00F8) Then
Exit Sub
End If
End If
Next l00FC
List2.AddItem l00F8
l00F8 = Mid(Frame3D3.Caption, 9, Len(Frame3D3.Caption))
Frame3D3.Caption = "Ignore: " + Trim(Str(l00F8 + 1))
End Sub

Sub Form_Load ()
Dim l010A As Variant
Dim l010E As Variant
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011C As Variant
Dim l0122 As Variant
Dim l0126 As Variant
Dim l012A As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Variant
Dim l0146 As Variant
Dim l014A As Variant
tell.Text = ""
If  server.Caption <> fn14F0() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
Option3D1.Value = True
extsub1050 server.hWnd, -1, 0, 0, 0, 0, 3
If  fn138() = 5 Then
l010A = fn5D0("Mem&bers", "Set Preferences")
Else
l010A = fn5D0("Mem&bers", "Preferences")
End If
l010E = extfn1018("AOL FRAME25", 0&)
30 :
DoEvents
l0112 = extfn10F8(l010E, "Preferences")
If  l0112 = 0 Then GoTo 30
l0116 = extfn10F8(l0112, "Mail")
l011C = extfn1168(l0116, 2)
l0122 = fn1480(l011C)
40 :
DoEvents
l0126 = extfn1018("_AOL_MODAL", 0&)
If  l0126 = 0 Then GoTo 40
l012A = extfn10F8(l0126, "close")
l012E = extfn1168(l012A, 3)
l0132 = extfnFE0(l012E, 1025, 0, 0)
l0136 = extfn10F8(l0126, "confirm")
l013A = extfn1168(l0136, 2)
l013E = extfnFE0(l013A, 1025, 1, 0)
l0142 = extfn10F8(l0126, "OK")
l0146 = fn1480(l0142)
l014A = extfnFE0(l0112, 16, 0, 0)
End Sub

Sub Form_LostFocus ()
Dim l014E As Variant
Dim l0152 As Variant
Dim l0156 As String
Dim l0158 As Variant
Dim l015C As Variant
Dim l0160 As Variant
Dim l0164 As Variant
Dim l0168 As Variant
Dim l016C As Variant
If  Timer1.Interval = 0 Then GoTo 45
If  server.Caption <> fn14F0() Then Unload server
l014E = extfn1018("AOL FRAME25", 0&)
l0152 = extfn10F8(l014E, "Welcome")
l0156$ = String(30, 0)
l0158 = extfnFA8(l0152, l0156$, 250)
l015C = Mid(l0156$, 10, 10)
l0160 = InStr(l015C, "!")
If  l0160 Then
l015C = Mid(l015C, 1, l0160 - 1)
End If
l0164 = fnA68("«‹- -– Fate Ultra WareZ Server By MaGuS –-- -›»")
DoEvents
DoEvents
l0168 = fnA68("«‹- -– Type " + Chr(34) + "/" + l015C + " Send X" + Chr(34) + " Where X Is The Index –-- -›»")
DoEvents
DoEvents
l016C = fnA68("«‹- -– Type " + Chr(34) + "/" + l015C + " Send List" + Chr(34) + " To Get List –-- -›»")
DoEvents
DoEvents
45 :
End Sub

Sub HScroll1_Change ()
End Sub

Sub Timer1_Timer ()
tell.Text = "1"
End Sub
