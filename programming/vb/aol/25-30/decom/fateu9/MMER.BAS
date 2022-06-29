' MMER.FRM
Option Explicit
Dim m001A As Variant
Dim m004A As Variant

Sub Check3D1_Click (Value As Integer)
If  Check3D1.Value = True Then
List2.Enabled = True
Text6.Enabled = True
Text11.Enabled = True
Command3D12.Enabled = True
Command3D13.Enabled = True
Command3D14.Enabled = True
Else
List2.Enabled = False
Text6.Enabled = False
Text11.Enabled = False
Command3D12.Enabled = False
Command3D13.Enabled = False
Command3D14.Enabled = False
End If
End Sub

Sub sub9748 ()
End Sub

Sub Command3D1_Click ()
Dim l0088 As Variant
Dim l008E As Variant
Dim l0092 As String
Dim l0096 As Variant
Dim l009A As Variant
Dim l00A0 As Variant
Dim l00A6 As Variant
Dim l00D8 As Variant
l0088 = extfn1018("AOL FRAME25", 0&)
l008E = extfn10F8(l0088, "Welcome")
l0092$ = String(30, 0)
l0096 = extfnFA8(l008E, l0092$, 250)
If  l0096 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l009A = extfn10F8(l0088, "AOL Live!")
l00A0 = extfn1088(l009A)
l00A6 = extfn10C0(l00A0, "_AOL_EDIT")
If  l00A6 = 0 Then
MsgBox "You must be in a chat room to Mass Mail!", 16
Exit Sub
End If
If  text7.Text = "" Then
MsgBox "Please enter in the minutes until the mass mail!", 16
Exit Sub
End If
If  Not IsNumeric(text7.Text) Then
MsgBox "Please enter in a numeric number for the minutes until the mass mail!", 16
Exit Sub
End If
If  Check3D1.Value = True Then
If  List2.ListCount = 0 Then
MsgBox "You must have at least one room if you want to run the room manager!", 16
Exit Sub
End If
End If
If  Text4.Text = "" Then
MsgBox "Please enter in a trigger phrase!", 16
Exit Sub
End If
If  m001A = 0 Then
m001A = 1
Command3D1.Caption = "&Stop Bot"
TIMER4.Enabled = True
TIMER4.Interval = 1
Timer1.Enabled = True
Timer3.Enabled = True
Timer1.Interval = 1
Timer3.Interval = 60000&
If  Check3D1.Value = True Then
Timer5.Enabled = True
Timer5.Interval = 1
End If
Else
m001A = 0
DoEvents
Command3D1.Caption = "&Start Bot"
Timer5.Interval = 0
Timer5.Enabled = False
TIMER4.Interval = 0
TIMER4.Enabled = False
Timer1.Interval = 0
Timer1.Enabled = False
Timer3.Interval = 0
Timer3.Enabled = False
l00D8 = fnA30(fn720())
End If
End Sub

Sub Command3D10_Click ()
If  Text2.Text = "" And List1.ListCount = 0 Then
MsgBox "You have nothing to save!", 16
Exit Sub
End If
mmsl.Show
mmsl.Caption = "Save Message"
mmsl.Command3D1.Caption = "Save"
mmsl.Option3D1.Caption = "Save Message"
mmsl.Option3D2.Caption = "Save Mass Mail List"
End Sub

Sub Command3D11_Click ()
mmsl.Show
mmsl.Caption = "Load Message"
mmsl.Command3D1.Caption = "Load"
mmsl.Option3D1.Caption = "Load Message"
mmsl.Option3D2.Caption = "Load Mass Mail List"
End Sub

Sub Command3D12_Click ()
Dim l00EC As Variant
l00EC = List2.ListIndex
On Error Resume Next
List2.RemoveItem l00EC
If  Err Then
MsgBox "Highlight the room you want to delete!", 16
Else
text7.Text = List2.ListCount * Text11.Text + Text11.Text
End If
End Sub

Sub Command3D13_Click ()
Dim l00F0 As Variant
Dim l00F4 As Variant
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
Dim l0104 As Variant
If  Text6.Text = "" Then GoTo 150
If  Text11.Text = 0 Then
Text11.Text = "1"
End If
For l00F0 = 0 To List2.ListCount
l00F4 = List2.List(l00F0)
l00F8 = InStr(1, l00F4, Text6.Text, 1)
If  l00F8 Then
If  Len(List2.List(l00F0)) = Len(Text6.Text) Then
GoTo 150
End If
End If
Next l00F0
List2.AddItem Text6.Text
150 :
l00FC = Int(List2.ListCount)
On Error Resume Next
l0100 = Int(Text11.Text)
If  Err Then
Text11.Text = "1"
l0100 = Int(Text11.Text)
End If
l0104 = l00FC * l0100
text7.Text = l0104 + l0100
Text6.Text = ""
End Sub

Sub Command3D14_Click ()
List2.Clear
text7.Text = Text11.Text
End Sub

Sub Command3D2_Click ()
Dim l0108 As Variant
Dim l010C As Variant
Dim l0110 As String
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l0120 As Variant
Dim l0124 As Variant
Dim l0128 As Variant
Dim l012C As Variant
Dim l0130 As String
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As String
Dim l0142 As Variant
Dim l0146 As Variant
Dim l0150 As Variant
Dim l015A As Variant
Dim l015E As Variant
l0108 = extfn1018("AOL FRAME25", 0&)
l010C = extfn10F8(l0108, "Welcome")
l0110$ = String(30, 0)
l0112 = extfnFA8(l010C, l0110$, 250)
If  l0112 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0116 = extfn1018("AOL FRAME25", 0&)
l0108 = extfn10F8(l0116, "AOL Live!")
l0108 = extfn1088(l0108)
If  l0108 = 0 Then
MsgBox "Room could not be found!", 16
Exit Sub
End If
l011A = extfn10C0(l0108, "_AOL_listbox")
l0120 = extfnFE0(l011A, 1036, 0, 0)
For l0124 = 0 To l0120 - 1
DoEvents
l0128 = extfn1018("AOL FRAME25", 0&)
l012C = extfn10F8(l0128, "Welcome")
l0130$ = String(l012C, 0)
l0132 = extfnFA8(l012C, l0130$, 250)
l0136 = l0132 - 1
On Error Resume Next
l013A = Mid(Left(l0130$, l0136), 10)
If  Err Then
Exit Sub
End If
l013E$ = String(30, 0)
l0142 = extfn1608(l0124, l013E$)
l0146 = InStr(l013E$, Chr(0))
If  l0146 Then
l013E$ = Mid(l013E$, 1, l0146 - 1)
End If
l0108 = InStr(1, Text1.Text, l013E$, 1)
If  l0108 Then
l010C = Mid(Text1.Text, l0108 + Len(l013E$), 1)
If  l010C = "," Then GoTo 3003
End If
l0108 = InStr(1, l013E$, "Guide", 1)
If  l0108 Then GoTo 3003
l0150 = InStr(1, l013E$, l013A, 1)
If  l0150 Then GoTo 3003
List1.AddItem l013E$
Text1.Text = Text1.Text + l013E$ + ","
l015A = Mid(Frame3D1.Caption, 15, Len(Frame3D1.Caption))
l015E = Int(l015A) + 1
Frame3D1.Caption = "People to MM: " + Str(l015E)
3003 :
Next l0124
l0116 = fnA30("«‹- -– Fate Ultra - The Room Has Been Added –-- -›»")
End Sub

Sub Command3D3_Click ()
Dim l0162 As Variant
Dim l0166 As Variant
Dim l016A As String
Dim l016C As Variant
If  List1.ListCount = 0 Then
MsgBox "You must have at least one person to Mass Mail", 16
Exit Sub
End If
l0162 = extfn1018("AOL FRAME25", 0&)
l0166 = extfn10F8(l0162, "Welcome")
l016A$ = String(30, 0)
l016C = extfnFA8(l0166, l016A$, 250)
If  l016C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
Timer2.Enabled = True
Timer2.Interval = 1
mmER.Hide
End Sub

Sub Command3D4_Click ()
Dim l017A As Variant
Dim l017E As Variant
Dim l0182 As Variant
Dim l0186 As Variant
Dim l018A As Variant
l017A = List1.ListIndex
l017E = List1.List(l017A)
On Error Resume Next
List1.RemoveItem l017A
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l0182 = InStr(1, Text1.Text, l017E + ",", 1)
If  l0182 Then
l0186 = Mid(Text1.Text, 1, l0182 - 1)
l018A = Mid(Text1.Text, l0182 + Len(l017E) + 1, Len(Text1.Text))
Text1.Text = l0186 + l018A
l017A = Mid(Frame3D1.Caption, 15, Len(Frame3D1))
l0186 = l017A - 1
Frame3D1.Caption = "People To MM: " + l0186
End If
End If
End Sub

Sub Command3D5_Click ()
If  m004A = 0 Then
mmER.Width = 7250
m004A = 1
Else
m004A = 0
mmER.Width = 5000
End If
End Sub

Sub Command3D6_Click ()
Frame3D1.Caption = "People To MM: 0"
List1.Clear
Text1.Text = ""
End Sub

Sub Command3D7_Click ()
Dim l0194 As Variant
Dim l0198 As Variant
Dim l019C As Variant
Dim l01A0 As Variant
If  Text5.Text = "" Then GoTo 201
For l0194 = 0 To List1.ListCount
l0198 = InStr(1, List1.List(l0194), Text5.Text, 1)
If  l0198 Then
If  Len(List1.List(l0194)) = Len(Text5.Text) Then
Text5.Text = ""
GoTo 201
End If
End If
Next l0194
List1.AddItem Text5
Text1.Text = Text1.Text + Text5.Text + ","
Text5 = ""
l019C = Mid(Frame3D1.Caption, 15, Len(Frame3D1.Caption))
l01A0 = Int(l019C) + 1
Frame3D1.Caption = "People To MM: " + Str(l01A0)
201 :
End Sub

Sub Command3D8_Click ()
mmER.WindowState = 1
End Sub

Sub Command3D9_Click ()
mmoptions.Show
End Sub

Sub Form_Load ()
Dim l01AA As Variant
L28A46:
mmoptions.Option3D1.Value = True
mmoptions.Option3D3.Value = True
mmER.Check3D3.Value = True
text7.Text = "5"
If  mmER.Caption <> fn528() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
extsub1050 mmER.hWnd, -1, 0, 0, 0, 0, 3
l01AA = 0
End Sub

Sub Form_Unload (p1 As Integer)
Unload mmoptions
Unload mmstatus
Unload mmsl
End Sub

Sub sub4420 ()
End Sub

Sub Option3D1_Click ()
End Sub

Sub Option3D2_Click ()
End Sub

Sub Text11_Change ()
Dim l01B8 As Variant
Dim l01BC As Variant
Dim l01C0 As Variant
On Error Resume Next
l01B8 = Int(List2.ListCount)
On Error Resume Next
l01BC = Int(Text11.Text)
l01C0 = l01B8 * l01BC
text7.Text = l01C0 + l01BC
End Sub

Sub Text5_KeyPress (KeyAscii As Integer)
Dim l01C6 As Variant
Dim l01CA As Variant
Dim l01CE As Variant
Dim l01D2 As Variant
l01C6 = Chr(KeyAscii)
If  l01C6 = Chr(13) Then
If  Text5.Text = "" Then GoTo 202
For l01C6 = 0 To List1.ListCount
l01CA = InStr(1, List1.List(l01C6), Text5.Text, 1)
If  l01CA Then
If  Len(List1.List(l01C6)) = Len(Text5.Text) Then
Text5.Text = ""
GoTo 202
End If
End If
Next l01C6
List1.AddItem Text5
Text1.Text = Text1.Text + Text5.Text + ","
Text5 = ""
l01CE = Mid(Frame3D1.Caption, 15, Len(Frame3D1.Caption))
l01D2 = Int(l01CE) + 1
Frame3D1.Caption = "People To MM: " + Str(l01D2)
End If
202 :
End Sub

Sub Text6_KeyPress (KeyAscii As Integer)
Dim l01D8 As Variant
Dim l01DC As Variant
Dim l01E0 As Variant
Dim l01E4 As Variant
Dim l01E8 As Variant
Dim l01EC As Variant
l01D8 = Chr(KeyAscii)
If  l01D8 = Chr(13) Then
If  Text11.Text = 0 Then
Text11.Text = "1"
End If
If  Text6.Text = "" Then GoTo 160
For l01D8 = 0 To List2.ListCount
l01DC = List2.List(l01D8)
l01E0 = InStr(1, l01DC, Text6.Text, 1)
If  l01E0 Then
If  Len(List2.List(l01D8)) = Len(Text6.Text) Then
GoTo 160
End If
End If
Next l01D8
List2.AddItem Text6.Text
160 :
Text6.Text = ""
l01E4 = Int(List2.ListCount)
On Error Resume Next
l01E8 = Int(Text11.Text)
If  Err Then
Text11.Text = "1"
l01E8 = Int(Text11.Text)
End If
l01EC = l01E4 * l01E8
text7.Text = l01EC + Text11.Text
Text6.Text = ""
End If
End Sub

Sub Timer1_Timer ()
Dim l01FA As Variant
Text8.Text = fn170()
l01FA = fn1E0()
End Sub

Sub Timer2_Timer ()
Dim l01FE As Variant
Dim l0202 As Variant
Dim l0206 As Variant
Dim l020C As Variant
Dim l0214 As Variant
Dim l0218 As Variant
Dim l021C As Variant
Dim l0222 As Variant
Dim l0228 As Variant
Dim l022C As Variant
Dim l0230 As Variant
Dim l0234 As Variant
Dim l0238 As Variant
Dim l023C As Variant
Dim l0240 As Variant
Dim l0244 As Variant
Dim l0248 As Variant
Dim l024C As Variant
Dim l0250 As Variant
Dim l0254 As Variant
Dim l0258 As Variant
Dim l025C As Variant
Dim l0260 As Variant
Dim l0264 As Variant
Dim l0268 As Variant
Dim l026C As Variant
Dim l0270 As Variant
Dim l0274 As Variant
Dim l0278 As Variant
Dim l027C As Variant
Dim l0280 As Variant
Dim l0284 As Variant
Dim l0288 As Variant
Dim l028C As Variant
Dim l0290 As Variant
Dim l0294 As Variant
Dim l0298 As Variant
Dim l029E As Variant
Dim l02A2 As Variant
Dim l02A6 As Variant
Dim l02AA As Variant
Dim l02B4 As Variant
Dim l02B8 As Variant
Dim l02BC As Variant
Dim l02C0 As Variant
Dim l02C4 As Variant
Dim l02C8 As String
Dim l02CA As Variant
Dim l02CE As Variant
Dim l02D2 As Variant
Dim l02D6 As Variant
Dim l02DA As Variant
Dim l02DE As String
Dim l02E0 As Variant
Dim l02E4 As Variant
Dim l02E8 As Variant
Dim l02EC As Variant
Dim l02F0 As Variant
Dim l02F6 As Variant
Dim l02FA As Variant
Dim l02FE As Variant
Dim l0302 As Variant
Dim l0306 As Variant
Dim l030A As Variant
Dim l030E As Variant
Dim l0312 As Variant
Dim l0316 As Variant
Dim l031A As Variant
Dim l031E As String
Dim l0320 As Variant
Dim l0324 As Variant
Dim l0328 As Variant
Dim l032C As Variant
Dim l0330 As Variant
Dim l0334 As Variant
Dim l0338 As Variant
Dim l033C As Variant
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As Variant
Dim l0350 As Variant
Dim l0354 As Variant
Dim l0358 As Variant
Dim l035C As Variant
Dim l0360 As Variant
Dim l0364 As Variant
Dim l0368 As Variant
Dim l036C As Variant
Dim l0370 As Variant
Dim l0374 As Variant
Dim l0378 As Variant
Dim l037C As Variant
Dim l0380 As Variant
Dim l0384 As Variant
Dim l0388 As Variant
Dim l038C As Variant
Dim l0390 As Variant
Dim l0394 As Variant
Dim l0398 As Variant
Dim l039C As Variant
Dim l03A0 As Variant
Dim l03A4 As Variant
Dim l03A8 As Variant
Dim l03AC As Variant
Dim l03B0 As Variant
Dim l03B4 As Variant
Dim l03BA As Variant
Dim l03BE As Variant
Dim l03C2 As Variant
Dim l03C6 As Variant
Dim l03CA As Variant
Dim l03D0 As Variant
Dim l03D4 As String
Dim l03D6 As Variant
Dim l03DA As Variant
Dim l03DE As Variant
Dim l03E2 As String
Dim l03E4 As Variant
Dim l03E8 As Variant
Dim l03EC As Variant
Dim l03F0 As Variant
Dim l03F4 As Variant
Dim l03F8 As Variant
If  mmoptions.Option3D7.Value = True Then
mmoptions.List2.Clear
For l01FE = 0 To mmoptions.List1.ListCount - 1
If  mmoptions.List1.Selected(l01FE) = True Then
mmoptions.List2.AddItem l01FE
End If
Next l01FE
End If
If  mmoptions.Option3D7.Value = True Then
mmoptions.Text6.Text = mmoptions.List2.ListCount
End If
mmstatus.Show
mmstatus.label2.Caption = " Mails Sent: 0"
If  mmoptions.Option3D2.Value = True Then
mmstatus.label3.Caption = " Mails Left: " + mmoptions.Text6.Text
End If
If  mmoptions.Option3D7.Value = True Then
mmstatus.label3.Caption = " Mails Left: " + mmoptions.Text6.Text
End If
mmER.Hide
Text1.Text = ""
For l01FE = 0 To List1.ListCount - 1
DoEvents
If  Timer2.Interval = 0 Then GoTo L2CAD2
l0202 = List1.List(l01FE)
If  mmoptions.Check3D2.Value = True Then
Text1.Text = Text1.Text + "(" + l0202 + "),"
Else
Text1.Text = Text1.Text + l0202 + ","
End If
Next l01FE
l0206 = extfn1018("AOL FRAME25", 0&)
If  Timer2.Interval = 0 Then GoTo L2CAD2
l020C = fn448()
If  Timer2.Interval = 0 Then GoTo L2CAD2
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l0214 = fn5D0("Mem&bers", "Set Preferences")
End If
If  fn138() = 8 Then
l0214 = fn5D0("Mem&bers", "Preferences")
End If
789 :
DoEvents
l0218 = extfn10F8(l0206, "Preferences")
l021C = extfn10F8(l0218, "Mail")
l0222 = extfn1168(l021C, 2)
If  l021C = 0 Then GoTo 789
l0228 = fn1480(l0222)
30 :
DoEvents
l022C = extfn1018("_AOL_Modal", 0&)
If  l022C = 0 Then GoTo 30
l0230 = extfn10F8(l022C, "Confirm mail")
l0234 = extfn1168(l0230, 2)
l0238 = extfnFE0(l0234, 1025, 1, 0&)
l023C = extfn10F8(l022C, "Close mail")
If  Timer2.Interval = 0 Then GoTo L2CAD2
l0240 = extfn1168(l023C, 3)
l0244 = extfnFE0(l0240, 1025, 0, 0&)
l0248 = extfn1018(0&, "Mail Preferences")
l024C = extfn10F8(l0248, "OK")
l0250 = fn1480(l024C)
l0254 = extfnFE0(l0218, 16, 0, 0)
If  mmoptions.Option3D3.Value = True Then
l0258 = extfn10F8(l0206, "New Mail")
Else
l0258 = extfn10F8(l0206, "Incoming Flashmail")
End If
If  mmoptions.Option3D5.Value = True Then
l0258 = extfn10F8(l0206, "Old Mail")
End If
If  mmoptions.Option3D6.Value = True Then
l0258 = extfn10F8(l0206, "Outgoing Mail")
End If
If  l0258 > 0 Then GoTo 200
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
If  mmoptions.Option3D3.Value = True Then l025C = fn5D0("&Mail", "Read &New Mail")
If  mmoptions.Option3D5.Value = True Then l025C = fn5D0("&Mail", "Check Mail You've &Read")
If  mmoptions.Option3D4.Value = True Then l025C = fn5D0("&Mail", "Read &Incoming Mail")
If  mmoptions.Option3D6.Value = True Then l025C = fn5D0("&Mail", "Check Mail You've &Sent")
End If
If  fn138() = 8 Then
If  mmoptions.Option3D3.Value = True Then l025C = fn5D0("&Mail", "Read &New Mail")
If  mmoptions.Option3D5.Value = True Then l025C = fn5D0("&Mail", "Check Mail You've &Read")
If  mmoptions.Option3D4.Value = True Then l025C = fn5D0("&Mail", "Read &Incoming FlashMail")
If  mmoptions.Option3D6.Value = True Then l025C = fn5D0("&Mail", "Check Mail You've &Sent")
End If
200 :
Do Until 3 > 4
If  mmoptions.Option3D3.Value = True Then
l0258 = extfn10F8(l0206, "New Mail")
Else
l0258 = extfn10F8(l0206, "Incoming Flashmail")
End If
If  mmoptions.Option3D5.Value = True Then l0258 = extfn10F8(l0206, "Old Mail")
If  mmoptions.Option3D6.Value = True Then l0258 = extfn10F8(l0206, "Outgoing Mail")
l0260 = extfn10C0(l0258, "_AOL_TREE")
l0264 = extfnFE0(l0260, 1036, 0, 0)
If  mmoptions.Option3D7.Value = True Then
l0268 = extfnFE0(l0260, 1031, Int(mmoptions.List2.List(0)), 0)
Else
l0268 = extfnFE0(l0260, 1031, 0, 0)
End If
If  mmoptions.Option3D4.Value = True Then
l025C = extfn10F8(l0258, "Open")
Else
l025C = extfn10F8(l0258, "Read")
End If
If  Timer2.Interval = 0 Then GoTo L2CAD2
l026C = fn1480(l025C)
l026C = extfn10F8(l0206, "Forward")
If  l026C > 0 Then GoTo 1230
DoEvents
Loop
1230 :
If  mmoptions.Option3D1.Value = True Then
mmoptions.Text6.Text = l0264
Else
End If
l0270 = 0
For l0274 = 1 To mmoptions.Text6.Text
If  mmoptions.Option3D7.Value = True Then
l0268 = extfnFE0(l0260, 1031, Int(mmoptions.List2.List(l0274 - 1)), 0)
Else
l0268 = extfnFE0(l0260, 1031, l0274 - 1, 0)
End If
60 :
Do Until 1 > 2
l0278 = extfn10C0(l0206, "AOL CHILD")
l027C = extfn10F8(l0278, "Forward")
l0280 = extfn1168(l027C, 2)
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
l028C = fn1480(l0280)
l0290 = Timer
While Timer - l0290 < .3#
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
DoEvents
Wend
l0278 = extfn10C0(l0206, "AOL CHILD")
l027C = extfn10F8(l0278, "Forward")
l0280 = extfn1168(l027C, 2)
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
l028C = fn1480(l0280)
l0290 = Timer
While Timer - l0290 < .3#
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
DoEvents
Wend
l0278 = extfn10C0(l0206, "AOL CHILD")
l027C = extfn10F8(l0278, "Forward")
l0280 = extfn1168(l027C, 2)
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
l028C = fn1480(l0280)
l0290 = Timer
While Timer - l0290 < .3#
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
DoEvents
Wend
l0278 = extfn10C0(l0206, "AOL CHILD")
l027C = extfn10F8(l0278, "Forward")
l0280 = extfn1168(l027C, 2)
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
l028C = fn1480(l0280)
l0290 = Timer
While Timer - l0290 < .3#
l0284 = extfn10C0(l0206, "AOL CHILD")
l0288 = extfn10F8(l0284, "Address" + Chr(13) + "Book")
If  l0288 > 0 Then GoTo 70
DoEvents
Wend
Loop
70 :
If  Timer2.Interval = 0 Then GoTo L2CAD2
If  mmoptions.Option3D7.Value = True Then
l0260 = extfn1088(l027C)
l0264 = extfnFE0(l0260, 16, 0, 0)
End If
l0294 = extfn10F8(l0206, "To:")
l0298 = extfn1168(l0294, 2)
l029E = extfn11A0(l0298, 12, 0, Text1.Text)
l02A2 = extfn1018("AOL FRAME25", 0&)
l02A6 = extfn10F8(extfn1088(l0294), "Subject:")
l02AA = l02AA + 1
If  Timer2.Interval = 0 Then GoTo L2CAD2
Text3.Text = l02AA
l02B4 = l02B4 + l02B8
l02BC = Fix(l02B4)
l02C0 = extfn1168(l02A6, 2)
If  mmoptions.Check3D4.Value = True Then
l02C4 = extfn11A0(l02C0, 12, 0, mmoptions.Text2.Text)
Else
l02C8$ = String(255, 0)
l02CA = extfn11A0(l02C0, 13, 255, l02C8$)
l02CE = Mid(l02C8$, 6, Len(l02C8$))
If  mmoptions.Check3D3.Value = True Then
l02CE = mmoptions.Text1.Text + l02CE
End If
l02D2 = extfn11A0(l02C0, 12, 0, l02CE)
End If
If  Timer2.Interval = 0 Then GoTo L2CAD2
l02D6 = extfn1018("AOL FRAME25", 0&)
l02DA = extfn10F8(l02D6, "Welcome")
l02DE$ = String$(l02DA, 0)
l02E0 = extfnFA8(l02DA, l02DE$, 250)
l02E4 = l02E0 - 1
l02E8 = Mid(Left(l02DE$, l02E4), 9)
l02EC = Chr(13) + Chr(10)
l0206 = extfn1018("AOL FRAME25", 0&)
l02A6 = extfn10F8(l0206, "Subject:")
l02F0 = extfn1168(l02A6, 1)
l02D6 = extfn1168(l02F0, 3)
l02EC = Chr(13) + Chr(10)
l02F6 = extfn11A0(l02D6, 12, 0, fn790() + l02EC + "«‹- -– Mail # " + Text3.Text + " Out Of " + mmoptions.Text6.Text + " –-- -›»" + l02EC + "«‹- -– Brought To You By" + l02E8 + " –-- -›»" + l02EC + l02EC + Text2.Text)
l02FA = extfn10C0(l0206, "AOL Child")
l02FE = extfn10F8(l02FA, "Send Now")
l0302 = extfn1168(l02FE, 2)
If  Timer2.Interval = 0 Then GoTo L2CAD2
l0306 = fn1480(l0302)
80 :
l030A = extfn10F8(l0284, "Address" + Chr(13) + "Book")
l030E = extfn1018("AOL FRAME25", 0&)
l0312 = extfn10F8(l030E, "Error")
If  l0312 > 0 Then
l0316 = extfn10F8(l030E, "Error")
l031A = extfn10C0(l0316, "_AOL_VIEW")
l031E$ = String(l031A, 0)
l0320 = extfnFE0(l031A, 14, 0, 0)
l0324 = extfn11A0(l031A, 13, l0320, l031E$)
Text1 = l031E$
l0328 = InStr(Text1, Chr(10))
l032C = Mid(Text1, l0328 + 3, Len(Text1))
Text1 = l032C
For l0330 = 1 To Len(Text1)
l0334 = InStr(Text1, "-")
If  l0334 Then
l0338 = Mid(Text1, 1, l0334 - 2)
l033C = InStr(1, Text1.Text, "mailbox", 1)
If  l033C Then
l0340 = "- Full Mailbox"
Else
l0340 = "- Unknown User"
End If
For l0344 = 0 To List1.ListCount - 1
l0348 = InStr(1, List1.List(l0344), l0338, 1)
If  l0348 Then
l034C = Len(l0338)
l0350 = Len(List1.List(l0344))
If  l034C <> l0350 Then
GoTo 87
End If
mmstatus.List1.AddItem List1.List(l0344) + l0340
List1.RemoveItem l0344
If  mmoptions.Check3D2.Value = True Then
l0354 = InStr(1, Text1.Text, l0344 + "),", 1)
If  l0354 Then
l0358 = Mid(Text1.Text, 1, l0354 - 1)
l035C = Mid(Text1.Text, l0354 + Len(l0344) + 2, Len(Text1.Text))
Text1.Text = l0358 + l035C
l0354 = Mid(Frame3D1.Caption, 15, Len(Frame3D1))
l0358 = l0354 - 1
Frame3D1.Caption = "People to MM: " + l0358
End If
Else
l0354 = InStr(1, Text1.Text, l0344 + ",", 1)
If  l0354 Then
l0358 = Mid(Text1.Text, 1, l0354 - 1)
l035C = Mid(Text1.Text, l0354 + Len(l0344) + 1, Len(Text1.Text))
Text1.Text = l0358 + l035C
l0354 = Mid(Frame3D1.Caption, 15, Len(Frame3D1))
l0358 = l0354 - 1
Frame3D1.Caption = "People to MM: " + l0358
End If
End If
GoTo 35
End If
87 :
Next l0344
35 :
l0360 = InStr(Text1, Chr(10))
If  l0360 Then
l0364 = Mid(Text1, l0360 + 1, Len(Text1))
Text1 = l0364
Else
GoTo 243
End If
Else
GoTo 243
End If
Next l0330
243 :
Text1 = ""
If  List1.ListCount = 0 Then
MsgBox "An error occurred: Nobody to massmail", 16
Timer2.Interval = 0
GoTo L2CAD2
Exit Sub
End If
For l0368 = 0 To List1.ListCount - 1
If  mmoptions.Option3D2.Value = True Then
Text1.Text = Text1.Text + "(" + List1.List(l0368) + "),"
Else
Text1.Text = Text1.Text + List1.List(l0368) + ","
End If
Next l0368
Frame3D1.Caption = "People to MM: " + List1.ListCount
l036C = extfn1018("AOL FRAME25", 0&)
l0370 = extfn10F8(l036C, "Error")
l0374 = extfn10F8(l0370, "OK")
l0378 = fn1480(l0374)
l037C = extfn10F8(l036C, "To:")
l0380 = extfn1168(l037C, 2)
l0384 = extfn11A0(l0380, 12, 0, Text1)
l0388 = extfn1088(l037C)
l033C = extfn10F8(l0388, "Send")
l0344 = extfn1168(l033C, 2)
l038C = fn1480(l0344)
End If
DoEvents
If  l030A > 0 Then GoTo 80
If  Timer2.Interval = 0 Then GoTo L2CAD2
mmstatus.label2.Caption = " Mails sent: " + Trim(Str(Int(l0274)))
mmstatus.label3.Caption = "Mails left: " + Str(mmoptions.Text6.Text - l0274)
If  mmoptions.Option3D3.Value = True Then
l0390 = extfn10F8(l0206, "New Mail")
Else
l0390 = extfn10F8(l0206, "Incoming Flashmail")
End If
If  mmoptions.Option3D5.Value = True Then l0390 = extfn10F8(l0206, "Old Mail")
If  mmoptions.Option3D6.Value = True Then l0390 = extfn10F8(l0206, "Outgoing Mail")
If  mmoptions.Option3D3.Value = True Then
If  mmoptions.Check3D1.Value = True Then
l0260 = extfn10C0(l0390, "_AOL_TREE")
If  mmoptions.Option3D7.Value = True Then
l0268 = extfnFE0(l0260, 1031, Int(mmoptions.List2.List(l0274 - 1)), 0)
Else
l0268 = extfnFE0(l0260, 1031, l0274 - 1, 0)
End If
l0394 = extfn10F8(l0390, "Keep As New")
l0398 = fn1480(l0394)
If  Timer2.Interval = 0 Then GoTo L2CAD2
End If
End If
l039C = extfn10C0(l0390, "_AOL_Tree")
If  Timer2.Interval = 0 Then GoTo L2CAD2
If  mmoptions.Option3D7.Value = True Then
If  l0274 = Int(mmoptions.Text6.Text) Then GoTo 46
l0268 = extfnFE0(l039C, 1031, Int(mmoptions.List2.List(l0274)), 0)
Do Until 2 > 3
If  mmoptions.Option3D3.Value = True Then
l0390 = extfn10F8(l0206, "Read")
End If
If  mmoptions.Option3D4.Value = True Then
l0390 = extfn10F8(l0206, "Open")
End If
If  mmoptions.Option3D5.Value = True Then
l0390 = extfn10F8(l0206, "Read")
End If
If  mmoptions.Option3D6.Value = True Then
l0390 = extfn10F8(l0206, "Read")
End If
l03A0 = fn1480(l0390)
l03A4 = extfn10F8(l0206, "Forward")
If  l03A4 > 0 Then GoTo 98
DoEvents
Loop
Else
If  l0274 = Int(mmoptions.Text6.Text) Then GoTo 46
l03A8 = extfn1018("AOL FRAME25", 0&)
l03AC = extfn10F8(l03A8, "Forward")
l03B0 = extfn1168(l03AC, 1)
l03B4 = fn1480(l03B0)
l03BA = fn13A0(.5#)
End If

98 :
DoEvents
Next l0274

46 :
If  Timer2.Interval = 0 Then GoTo L2CAD2
l03C2 = extfnFE0(l03BE, 16, 0, 0)
DoEvents
If  mmoptions.Option3D3.Value = True Then
l03C6 = extfn10F8(l0206, "New Mail")
Else
l03C6 = extfn10F8(l0206, "Incoming Flashmail")
End If
If  mmoptions.Option3D5.Value = True Then l03C6 = extfn10F8(l0206, "Old Mail")
If  mmoptions.Option3D6.Value = True Then l03C6 = extfn10F8(l0206, "Outgoing Mail")
l03CA = extfnFE0(l03C6, 16, 0, 0)
DoEvents
l03D0 = fn480()
DoEvents
Unload mmstatus
mmER.Show
If  mmoptions.signoff.Value = True Then
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l0244 = extfn1018("AOL FRAME25", 0&)
l0248 = extfn10F8(l0244, "Welcome")
l03D4$ = String(30, 0)
l0250 = extfnFA8(l0248, l03D4$, 250)
If  l0250 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l01FE = extfn1018("AOL FRAME25", 0&)
l0202 = extfnFE0(l01FE, 16, 0, 0)
14 :
DoEvents
l0206 = extfn1018("_AOL_MODAL", 0&)
If  l0206 = 0 Then GoTo 14
l020C = extfn10F8(l0206, "&Yes")
l0214 = fn1480(l020C)
Do Until 2 > 3
l03D6 = extfn1018("#32770", "Download Manager")
If  l03D6 > 0 Then
l03DA = extfn10F8(l03D6, "&No")
l03DE = fn1480(l03DA)
End If
l03DA = extfn10F8(l01FE, "Goodbye")
If  l03DA > 0 Then
GoTo L2CAD2
End If
DoEvents
Loop
Else
l0244 = extfn1018("AOL FRAME25", 0&)
l0248 = extfn10F8(l0244, "Welcome")
l03E2$ = String(30, 0)
l0250 = extfnFA8(l0248, l03E2$, 250)
If  l0250 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l01FE = extfn1018("AOL FRAME25", 0&)
l0202 = extfnFE0(l01FE, 16, 0, 0)
29 :
DoEvents
l0206 = extfn10F8(l01FE, "Exit?")
If  l0206 = 0 Then GoTo 29
l020C = extfn10C0(l0206, "_AOL_icon")
l0214 = extfn1168(l020C, 2)
l03D6 = extfn1168(l0214, 2)
l03DA = extfn1168(l03D6, 2)
l03DE = extfn1168(l03DA, 2)
l03E4 = extfn1168(l03DE, 2)
l03E8 = fn1480(l03E4)
Do Until 2 > 3
l03EC = extfn1018("#32770", "Download Manager")
If  l03EC > 0 Then
l03F0 = extfn10F8(l03EC, "&No")
l03F4 = fn1480(l03F0)
End If
l03F8 = extfn10F8(l01FE, "Goodbye")
If  l03F8 > 0 Then
GoTo L2CAD2
End If
DoEvents
Loop
GoTo L2CAD2
End If
End If

L2CAD2:
Unload mmstatus
Timer2.Interval = 0
Text1.Text = ""
For l01FE = 0 To List1.ListCount - 1
DoEvents
l0202 = List1.List(l01FE)
Text1.Text = Text1.Text + l0202 + ","
Next l01FE
mmER.Show
End Sub

Sub Timer3_Timer ()
Dim l0404 As Variant
Dim l040C As Variant
Dim l0414 As Variant
Dim l0418 As Variant
Dim l041E As Variant
text7.Text = text7.Text - 1
If  text7.Text = "0" Then GoTo L2CE56
Text12.Text = Trim(Str(Int(Text12.Text) + 1))
If  Timer3.Interval = 0 Then GoTo L2CF62
If  Timer3.Interval = 0 Then GoTo L2CF62
If  text7.Text = "0" Then GoTo L2CE56
If  text7.Text = "1" Then GoTo L2CD6A
If  Timer3.Interval = 0 Then GoTo L2CF62
If  Timer3.Interval = 0 Then GoTo L2CF62
DoEvents
DoEvents
l0404 = fnA30(fn608())
DoEvents
DoEvents
l040C = fnA30(fn640() + Trim(text7.Text) + fn678())
DoEvents
DoEvents
l0414 = fnA30(fn6B0() + Text4.Text + fn6E8())
DoEvents
DoEvents
If  Timer3.Interval = 0 Then GoTo L2CF62
GoTo L2CF62

L2CD6A:
If  Timer3.Interval = 0 Then GoTo L2CF62
DoEvents
l0404 = fnA30(fn608())
DoEvents
DoEvents
l040C = fnA30(fn640() + "1" + fn678())
DoEvents
DoEvents
l0414 = fnA30(fn6B0() + Text4.Text + fn6E8())
DoEvents
DoEvents
GoTo L2CF62
If  Timer3.Interval = 0 Then GoTo L2CF62

L2CE56:
If  Timer3.Interval = 0 Then GoTo L2CF62
DoEvents
l0418 = 0
l041E = fnA30(fn758())
DoEvents
Command3D1.Caption = "&Start Bot"
If  Timer3.Interval = 0 Then GoTo L2CF62
Timer1.Interval = 0
Timer1.Enabled = False
Timer3.Interval = 0
Timer3.Enabled = False
TIMER4.Interval = 0
TIMER4.Enabled = False
Timer5.Interval = 0
Timer5.Enabled = False
Command3D3_Click
DoEvents
GoTo L2CF62
L2CF62:
DoEvents
End Sub

Sub TIMER4_Timer ()
Dim l0422 As Variant
Dim l0426 As Variant
Dim l042A As Variant
If  text7.Text = 1 Then GoTo 12
If  text7.Text >= 2 Then GoTo 13
If  TIMER4.Interval = 0 Then GoTo L2D120
12 :
l0422 = fnA30(fn608())
DoEvents
l0426 = fnA30(fn640() + "1" + fn678())
DoEvents
l042A = fnA30(fn6B0() + Text4.Text + fn6E8())
DoEvents
GoTo L2D120

13 :
DoEvents
l0422 = fnA30(fn608())
DoEvents
l0426 = fnA30(fn640() + Trim(text7.Text) + fn678())
DoEvents
l042A = fnA30(fn6B0() + Text4.Text + fn6E8())
DoEvents
GoTo L2D120

L2D120:
TIMER4.Interval = 0
End Sub

Sub Timer5_Timer ()
Dim l042E As Variant
Dim l0432 As Variant
Dim l0436 As Variant
Dim l043C As Variant
Dim l0440 As Variant
Dim l0444 As Variant
Dim l0448 As Variant
Dim l044C As Variant
Dim l0450 As Variant
Dim l0454 As Variant
Dim l0458 As Variant
Dim l045C As Variant
Dim l0460 As Variant
Dim l0464 As Variant
Dim l046A As Variant
Dim l046E As Variant
Dim l0472 As Variant
If  Timer5.Interval = 0 Then GoTo L2D674
l042E = 1
If  Int(Text12.Text) = Int(Text11.Text) Then
Text12.Text = "0"
Timer3.Interval = 0
Timer3.Enabled = False
Timer1.Interval = 0
Timer1.Enabled = False
If  Timer5.Interval = 0 Then GoTo L2D674
l0432 = fnA30("«‹- -– Fate Ultra - Changing To: " + List2.List(0) + " –-- -›»")
l0436 = List2.List(0)
l043C = fnBF0(l0436)
92 :
If  Timer5.Interval = 0 Then GoTo 140
l0440 = extfn1018("AOL FRAME25", 0&)
l0444 = fn5D0("&Go To", "Keyword...")
91 :
DoEvents
l0448 = extfn10F8(l0440, "Keyword")
If  Timer5.Interval = 0 Then GoTo L2D674
If  Timer5.Interval = 0 Then GoTo 140
If  l0448 = 0 Then GoTo 91
l044C = extfn10C0(l0448, "_AOL_EDIT")
l0450 = extfn11A0(l044C, 12, 0, "aol://2719:2-2-" + List2.List(0))
l0454 = extfn10C0(l0448, "_AOL_ICON")
l0458 = fn1480(l0454)
Do Until 2 > 3
If  Timer5.Interval = 0 Then GoTo 140
l045C = extfn1018("#32770", "America Online")
If  l045C > 0 Then
If  Timer5.Interval = 0 Then GoTo 140
l0460 = extfn10F8(l045C, "OK")
l0464 = fn1480(l0460)
l042E = l042E + 1
GoTo 92
End If
l046A = fn288()
If  l046A = l043C Then
If  Int(l042E) > 1 Then
l046E = "s"
Else
l046E = ""
End If
l0472 = fn13A0(.6#)
List2.RemoveItem 0
Timer3.Enabled = True
Timer3.Interval = 60000&
Timer1.Enabled = True
Timer1.Interval = 1
GoTo 140
End If
DoEvents
Loop
End If
L2D674:
140 :
End Sub

Sub sub9980 ()
Unload mmER
End Sub
