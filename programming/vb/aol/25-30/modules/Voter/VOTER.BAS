' VOTER.FRM
Option Explicit
Declare Function extfn2EAA Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn2EE2 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long

Sub Command3D1_Click ()
Dim l0022 As Variant
Dim l0026 As Variant
Dim l002A As String
Dim l002C As Variant
Dim l0030 As Variant
Dim l0034 As Variant
Dim l0038 As Variant
Dim l0092 As Variant
l0022 = extfn678("AOL FRAME25", 0&)
l0026 = extfn640(l0022, "Welcome")
l002A$ = String(30, 0)
l002C = extfn918(l0026, l002A$, 250)
If  l002C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1.Text = Text2.Text Then
MsgBox "Agree and Disagree must be different!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "You must enter in something to vote!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You must enter in something to vote!", 16
Exit Sub
End If
If  Text3.Text = "" Then
MsgBox "You must enter in something to vote!", 16
Exit Sub
End If
If  Text4.Text = "" Then
MsgBox "You must enter in when the results will be scrolled!", 16
Exit Sub
End If
If  Not IsNumeric(Text4.Text) Then
MsgBox "You must enter in a numeric number for when the results will be scrolled!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
l0030 = fn250("              «‹-.¸.­-·´¯`¤ ' Vote Bot' aYe NaKo FiNaL ¤´¯`·-­.¸.­›»")
DoEvents
l0034 = fn250(" «‹-.¸.-­·´¯`¤ Type ''Yes'' or ''No'' Depending on Your Answer ¤´¯`·-­.¸.­›»")
sub2C0 (1.25#)
DoEvents
l0038 = fn250("«‹-.¸.­-·´¯`¤ (Question) " & Text1.Text & " ¤´¯`·-­.¸.­›»")
DoEvents
l0038 = fn250("       «‹-.¸.­-·´¯`¤ Your Votes Will Be Counted In " & Text4.Text & " Min. ¤´¯`·-­.¸.­›»")
sub2C0 (1.25#)
DoEvents
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label3.Caption = "Yes: 0"
Label5.Caption = "No: 0"
List1.Clear
Timer1.Enabled = True
Timer2.Enabled = True
Timer1.Interval = 1
Timer2.Interval = 60000&
Command3D1.Caption = "Stop Vote"
Else
Timer1.Interval = 0
Timer1.Enabled = False
Timer2.Interval = 0
Timer2.Enabled = False
Command3D1.Caption = "Start Vote"
l0092 = fn250("«‹-.¸.­-·´¯`¤ Voting Has Been Stopped ¤´¯`·-­.¸.­›»")
End If
End Sub

Sub Command3D2_Click ()
Label3.Caption = "Yes: 0"
Yes.Clear
End Sub

Sub Command3D3_Click ()
Label5.Caption = "No: 0"
No.Clear
End Sub

Sub Command3D4_Click ()
vote.WindowState = 1
End Sub

Function fn2CEA () As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As String
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As String
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l010A As Variant
Dim l010E As Variant
l00AE = extfn678("AOL FRAME25", 0&)
l00B2 = extfn640(l00AE, "Welcome")
l00B6$ = String(30, 0)
l00B8 = extfn918(l00B2, l00B6$, 250)
l00BC = Mid(l00B6$, 10, 10)
l00C0 = InStr(l00BC, "!")
If  l00C0 Then
l00BC = Mid(l00BC, 1, l00C0 - 1)
End If

32 :
If  vote.Command3D1.Caption = "Start Vote" Then GoTo L7D68E
DoEvents
l00C4 = extfn678("AOL FRAME25", 0&)
l00C8 = extfn640(l00C4, "AOL Live!")
l00C4 = extfnCD0(l00C8)
l00C8 = extfnDB0(l00C4, "_AOL_VIEW")
If  l00C8 = 0 Then
Exit Function
End If
l00CC = extfn2EAA(l00C8, 14, 0, 0)
If  l00CC = Len(vote.Text5.Text) Then GoTo 32
l00D0$ = String(l00CC, 0)
l00D2 = extfn2EE2(l00C8, 13, l00CC, l00D0$)
l00D6 = Len(vote.Text5.Text)
l00DA = Len(l00D0$) - l00D6
On Error Resume Next
l00DE = Right(l00D0$, l00DA)
l00DE = Mid(l00DE, 3, Len(l00DE))
34 :
If  vote.Command3D1.Caption = "Start Vote" Then GoTo L7D68E
DoEvents
l00E2 = InStr(l00DE, Chr(10))
If  l00E2 Then
l00E6 = Mid(l00DE, 1, l00E2 - 2)
vote.Text5.Text = l00E6
If  vote.Text5.Text <> "" Then vote.Text7.Text = vote.Text5.Text
l00EA = InStr(vote.Text7.Text, Chr(0))
If  l00EA Then
vote.Text7.Text = Mid(vote.Text7.Text, 1, l00EA - 1)
End If
l00EE = InStr(1, vote.Text7.Text, vote.Text2.Text, 1)
If  l00EE Then
l00F2 = Mid(vote.Text7.Text, l00EE - 2, 2)
If  l00F2 <> ":" + Chr(9) Then
GoTo 36
End If
l00F2 = InStr(vote.Text7, ":")
l00F6 = Mid(vote.Text7.Text, 1, l00F2 - 1)
For l00FA = 0 To vote.List1.ListCount - 1
l00FE = InStr(1, vote.List1.List(l00FA), l00F6, 1)
If  l00FE Then GoTo 36
Next l00FA
If  l00F6 = l00BC Then GoTo 36
vote.List1.AddItem l00F6
vote.Yes.AddItem l00F6
sub2C0 (.6#)
l00AA = fn250("«‹-.¸.­-·´¯`¤ " & l00F6 & " Thank You For Your Vote ¤´¯`·-­.¸.­›»")
l010A = Mid(vote.Label3.Caption, 6, Len(vote.Label3.Caption))
l010E = l010A + 1
vote.Label3.Caption = "Yes: " + l010E
End If
l00EE = InStr(1, vote.Text7.Text, vote.Text3.Text, 1)
If  l00EE Then
l00F2 = Mid(vote.Text7.Text, l00EE - 2, 2)
If  l00F2 <> ":" + Chr(9) Then
GoTo 36
End If
l00F2 = InStr(vote.Text7, ":")
l00F6 = Mid(vote.Text7.Text, 1, l00F2 - 1)
For l00FA = 0 To vote.List1.ListCount - 1
l00FE = InStr(1, vote.List1.List(l00FA), l00F6, 1)
If  l00FE Then GoTo 36
Next l00FA
If  l00F6 = l00BC Then GoTo 36
vote.List1.AddItem l00F6
vote.No.AddItem l00F6
 sub2C0 (.6#)
l00AA = fn250("«‹-.¸.­-·´¯`¤ " & l00F6 & " Thank You For Your Vote ¤´¯`·-­.¸.­›»")
l010A = Mid(vote.Label5.Caption, 5, Len(vote.Label5.Caption))
l010E = l010A + 1
vote.Label5.Caption = "No: " + l010E
End If

36 :
l00DE = Mid(l00DE, l00E2 + 1, Len(l00DE))
GoTo 34
End If
vote.Text5.Text = l00DE
If  vote.Text5.Text <> "" Then vote.Text7.Text = vote.Text5.Text
l00EA = InStr(vote.Text7.Text, Chr(0))
If  l00EA Then
vote.Text7.Text = Mid(vote.Text7.Text, 1, l00EA - 1)
End If
l00EE = InStr(1, vote.Text7.Text, vote.Text2.Text, 1)
If  l00EE Then
l00F2 = Mid(vote.Text7.Text, l00EE - 2, 2)
If  l00F2 <> ":" + Chr(9) Then
GoTo 35
End If
l00F2 = InStr(vote.Text7, ":")
l00F6 = Mid(vote.Text7.Text, 1, l00F2 - 1)
For l00FA = 0 To vote.List1.ListCount - 1
l00FE = InStr(1, vote.List1.List(l00FA), l00F6, 1)
If  l00FE Then GoTo 35
Next l00FA
If  l00F6 = l00BC Then GoTo 35
vote.List1.AddItem l00F6
vote.Yes.AddItem l00F6
sub2C0 (.6#)
l00AA = fn250("«‹-.¸.­-·´¯`¤ " & l00F6 & " Thank You For Your Vote ¤´¯`·-­.¸.­›»")
l010A = Mid(vote.Label3.Caption, 6, Len(vote.Label3.Caption))
l010E = l010A + 1
vote.Label3.Caption = "Yes: " + l010E
End If
l00EE = InStr(1, vote.Text7.Text, vote.Text3.Text, 1)
If  l00EE Then
l00F2 = Mid(vote.Text7.Text, l00EE - 2, 2)
If  l00F2 <> ":" + Chr(9) Then
GoTo 35
End If
l00F2 = InStr(vote.Text7, ":")
l00F6 = Mid(vote.Text7.Text, 1, l00F2 - 1)
For l00FA = 0 To vote.List1.ListCount - 1
l00FE = InStr(1, vote.List1.List(l00FA), l00F6, 1)
If  l00FE Then GoTo 35
Next l00FA
If  l00F6 = l00BC Then GoTo 35
vote.List1.AddItem l00F6
vote.No.AddItem l00F6
sub2C0 (.6#)
l00AA = fn250("«‹-.¸.­-·´¯`¤ " & l00F6 & " Thank You For Your Vote ¤´¯`·-­.¸.­›»")
l010A = Mid(vote.Label5.Caption, 5, Len(vote.Label5.Caption))
l010E = l010A + 1
vote.Label5.Caption = "No: " + l010E
End If

35 :
If  vote.Command3D1.Caption = "Start Vote" Then GoTo L7D68E
vote.Text5.Text = l00D0$
DoEvents
GoTo 32
L7D68E:
vote.Text5.Text = ""
vote.Text5.Text = ""
vote.Text7.Text = ""
l00AA = fn250("«‹-.¸.­-·´¯`¤ Voting Has Been Stopped ¤´¯`·-­.¸.­›»")
End Function

Function fn2D22 () As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
Dim l0126 As Variant
Dim l012A As Variant
Dim l012E As String
Dim l0130 As Variant
l011A = extfn678("AOL FRAME25", 0&)
l011E = extfn640(l011A, "AOL Live!")
l0122 = extfnCD0(l011E)
l0126 = extfn640(l0122, "_AOL_VIEW")
l012A = extfn2EAA(l0126, 14, 0, 0)
l012E$ = String(l012A, 0)
l0130 = extfn2EE2(l0126, 13, l012A, l012E$)
fn2D22 = l012E$
End Function

Sub Form_Load ()
sub288 Me
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Sub mnudone_Click ()
Timer1.Enabled = False
Timer2.Enabled = False
Unload Me
End Sub

Sub mnumin_Click ()
Me.WindowState = 1
End Sub

Sub SSCommand1_Click ()
Dim l0138 As Variant
l0138 = fn250("              «‹-.¸.­-·´¯`¤ ' Vote Bot' aYe NaKo FiNaL ¤´¯`·-­.¸.­›»")
DoEvents
l0138 = fn250(" «‹-.¸.-­·´¯`¤ Type ''Yes'' or ''No'' Depending On Your Answer ¤´¯`·-­.¸.­›»")
sub2C0 (1.25#)
DoEvents
l0138 = fn250("«‹-.¸.­-·´¯`¤ (Question) " & Text1.Text & " ¤´¯`·-­.¸.­›»")
DoEvents
l0138 = fn250("       «‹-.¸.­-·´¯`¤ Your Votes Will Be Counted In " & Text4.Text & " Min. ¤´¯`·-­.¸.­›»")
DoEvents
sub2C0 (1.25#)
End Sub

Sub Timer1_Timer ()
Dim l013C As Variant
Text5.Text = fn2D22()
l013C = fn2CEA()
End Sub

Sub Timer2_Timer ()
Dim l0140 As Variant
Dim l014C As Variant
Dim l0150 As Variant
Text4.Text = Text4.Text - 1
If  Text4.Text = "0" Then
l0140 = fn250("«‹-.¸.­-·´¯`¤ ' Vote Bot' aYe NaKo FiNaL ßy JePP¤´¯`·-­.¸.­›»")
DoEvents
l0140 = fn250("   «‹-.¸.­-·´¯`¤ The Votes Have Been Tallied Up  ¤´¯`·-­.¸.­›»")
sub2C0 (1.25#)
DoEvents
l0140 = fn250("«‹-.¸.­-·´¯`¤ (Question) " & Text1.Text & " ¤´¯`·-­.¸.­›»")
DoEvents
l014C = Mid(Label3.Caption, 6, Len(Label3.Caption))
l0150 = Mid(Label5.Caption, 5, Len(Label5.Caption))
l0140 = fn250("«‹-.¸.­-·´¯`¤ " & l014C & " Voted ''Yes'' + " & l0150 & " Voted ''No'' ¤´¯`·-­.¸.­›»")
DoEvents
Command3D1.Caption = "Start Vote"
l0140 = fn250("«‹-.¸.­-·´¯`¤ Voting Has Been Stopped ¤´¯`·-­.¸.­›»")
Timer1.Interval = 0
Timer2.Interval = 0
GoTo 20
End If
l0140 = fn250("       «‹-.¸.­-·´¯`¤ ' Vote Bot' aYe NaKo FiNaL ßy JePP¤´¯`·-­.¸.­›»")
DoEvents
l0140 = fn250(" «‹-.¸.-­·´¯`¤ Type ''Yes'' or ''No'' Depending on Your Answer ¤´¯`·-­.¸.­›»")
sub2C0 (1.25#)
DoEvents
l0140 = fn250("«‹-.¸.­-·´¯`¤ (Question) " & Text1.Text & " ¤´¯`·-­.¸.­›»")
DoEvents
l0140 = fn250("       «‹-.¸.­-·´¯`¤ Your Votes Will Be Counted In " & Text4.Text & " Min. ¤´¯`·-­.¸.­›»")
DoEvents
sub2C0 (1.25#)
Yes.Clear
No.Clear
GoTo 30
20 :
l0140 = fn250("«‹-.¸.­-·´¯`¤ Voting Has Been Stopped ¤´¯`·-­.¸.­›»")
Timer2.Interval = 0

30 :
End Sub
