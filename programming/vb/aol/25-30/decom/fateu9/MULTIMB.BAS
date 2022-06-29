' MULTIMB.FRM
Option Explicit

Sub co_Click ()
Dim l005E As Variant
Dim l0062 As Variant
Dim l006C As Variant
Clipboard.Clear
l005E = 1
For l0062 = 0 To List2.ListCount - 1
If  List2.Selected(l0062) = True Then
If  l005E = 1 Then
l006C = l006C + List2.List(l0062)
l005E = 0
Else
l006C = l006C + Chr(13) + Chr(10) + List2.List(l0062)
End If
End If
Next l0062
Clipboard.SetText l006C
End Sub

Sub Command3D1_Click ()
End Sub

Sub Command3D3_Click ()
multimb.WindowState = 1
End Sub

Sub cu_Click ()
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Clipboard.Clear
l0074 = 1
For l0078 = 0 To List2.ListCount - 1
24 :
On Error Resume Next
If  List2.Selected(l0078) = True Then
If  l0074 = 1 Then
l007C = l007C + List2.List(l0078)
l0074 = 0
Else
l007C = l007C + Chr(13) + Chr(10) + List2.List(l0078)
End If
List2.RemoveItem l0078
If  l0078 = List2.ListCount Then GoTo 25
GoTo 24
End If
Next l0078
25 :
Clipboard.SetText l007C
Select Case List1.List(List1.ListIndex)
Case "New Mail"
If  List2.ListCount > 1 Then
Panel3D6.Caption = "New Mail Contains " + List2.ListCount + " Mails"
Else
Panel3D6.Caption = "New Mail Contains " + List2.ListCount + " Mail"
End If
Case "Old Mail"
If  List2.ListCount > 1 Then
Panel3D6.Caption = "Old Mail Contains " + List2.ListCount + " Mails"
Else
Panel3D6.Caption = "Old Mail Contains " + List2.ListCount + " Mail"
End If
Case "Outgoing Mail"
If  List2.ListCount > 1 Then
Panel3D6.Caption = "Outgoing Mail Contains " + List2.ListCount + " Mails"
Else
Panel3D6.Caption = "Outgoing Mail Contains " + List2.ListCount + " Mail"
End If
Case "Incoming FlashMail"
If  List2.ListCount > 1 Then
Panel3D6.Caption = "Incoming FlashMail Contains " + List2.ListCount + " Mails"
Else
Panel3D6.Caption = "Incoming FlashMail Contains " + List2.ListCount + " Mail"
End If
Case "Outgoing FlashMail"
If  List2.ListCount > 1 Then
Panel3D6.Caption = "Outgoing FlashMail Contains " + List2.ListCount + " Mails"
Else
Panel3D6.Caption = "Outgoing FlashMail Contains " + List2.ListCount + " Mail"
End If
Case "Personal Filing Cabinet"
If  List2.ListCount > 1 Then
Panel3D6.Caption = "Personal Filing Cabinet Contains " + List2.ListCount + " Mails"
Else
Panel3D6.Caption = "Personal Filing Cabinet Contains " + List2.ListCount + " Mail"
End If
End Select
End Sub

Sub ex_Click ()
Unload multimb
End Sub

Sub Form_Load ()
extsub1050 multimb.hWnd, -1, 0, 0, 0, 0, 3
List1.AddItem "New Mail"
List1.AddItem "Old Mail"
List1.AddItem "Outgoing Mail"
List1.AddItem "Incoming FlashMail"
List1.AddItem "Outgoing FlashMail"
List1.AddItem "Personal Filing Cabinet"
End Sub

Sub List1_Click ()
Dim l008C As Variant
Dim l0098 As Variant
Dim l009E As Variant
Dim l00A2 As String
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00B0 As Variant
Dim l00B6 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As String
Dim l00C8 As Variant
l008C = List1.List(List1.ListIndex)
Select Case l008C
Case "New Mail"
Panel3D5.Caption = "Contents of New Mail"
List2.Clear
l0098 = extfn1018("AOL FRAME25", 0&)
l009E = extfn10F8(l0098, "Welcome")
l00A2$ = String(30, 0)
l00A6 = extfnFA8(l009E, l00A2$, 250)
If  l00A6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
16 :
l008C = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l008C, "New Mail")
l00B0 = extfn10C0(l00AA, "_AOL_TREE")
If  l00AA = 0 Then
If  l00B0 = 0 Then
l00B6 = fn5D0("&Mail", "Read &New Mail")
MsgBox "Please wait for your mail to be full!", 16
GoTo 16
End If
End If
l00BC = extfnFE0(l00B0, 1036, 0, 0)
For l00C0 = 0 To l00BC - 1
l00C4$ = String(255, 0)
l00C8 = extfn11A0(l00B0, 1034, l00C0, l00C4$)
List2.AddItem l00C4$
Next l00C0
If  l00BC > 1 Then
Panel3D6.Caption = "New Mail Contains " + l00BC + " Mails"
Else
Panel3D6.Caption = "New Mail Contains " + l00BC + " Mail"
End If
Case "Old Mail"
Panel3D5.Caption = "Contents of Old Mail"
List2.Clear
l0098 = extfn1018("AOL FRAME25", 0&)
l009E = extfn10F8(l0098, "Welcome")
l00A2$ = String(30, 0)
l00A6 = extfnFA8(l009E, l00A2$, 250)
If  l00A6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
15 :
l008C = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l008C, "Old Mail")
l00B0 = extfn10C0(l00AA, "_AOL_TREE")
If  l00AA = 0 Then
If  l00B0 = 0 Then
l00B6 = fn5D0("&Mail", "Check Mail You've &Read")
MsgBox "Please wait for your mail to be full!", 16
GoTo 15
End If
End If
l00BC = extfnFE0(l00B0, 1036, 0, 0)
For l00C0 = 0 To l00BC - 1
l00C4$ = String(255, 0)
l00C8 = extfn11A0(l00B0, 1034, l00C0, l00C4$)
List2.AddItem l00C4$
Next l00C0
If  l00BC > 1 Then
Panel3D6.Caption = "Old Mail Contains " + l00BC + " Mails"
Else
Panel3D6.Caption = "Old Mail Contains " + l00BC + " Mail"
End If
Case "Outgoing Mail"
Panel3D5.Caption = "Contents of Outgoing Mail"
List2.Clear
l0098 = extfn1018("AOL FRAME25", 0&)
l009E = extfn10F8(l0098, "Welcome")
l00A2$ = String(30, 0)
l00A6 = extfnFA8(l009E, l00A2$, 250)
If  l00A6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
14 :
l008C = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l008C, "Outgoing Mail")
l00B0 = extfn10C0(l00AA, "_AOL_TREE")
If  l00AA = 0 Then
If  l00B0 = 0 Then
l00B6 = fn5D0("&Mail", "Check Mail You've &Sent")
MsgBox "Please wait for your mail to be full!", 16
GoTo 14
End If
End If
l00BC = extfnFE0(l00B0, 1036, 0, 0)
For l00C0 = 0 To l00BC - 1
l00C4$ = String(255, 0)
l00C8 = extfn11A0(l00B0, 1034, l00C0, l00C4$)
List2.AddItem l00C4$
Next l00C0
If  l00BC > 1 Then
Panel3D6.Caption = "Outgoing Mail Contains " + l00BC + " Mails"
Else
Panel3D6.Caption = "Outgoing Mail Contains " + l00BC + " Mail"
End If
Case "Incoming FlashMail"
Panel3D5.Caption = "Contents of Incoming FlashMail"
List2.Clear
l0098 = extfn1018("AOL FRAME25", 0&)
l009E = extfn10F8(l0098, "Welcome")
l00A2$ = String(30, 0)
l00A6 = extfnFA8(l009E, l00A2$, 250)
If  l00A6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
13 :
l008C = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l008C, "Incoming FlashMail")
l00B0 = extfn10C0(l00AA, "_AOL_TREE")
If  l00AA = 0 Then
If  l00B0 = 0 Then
If  fn138() = 5 Then
l00B6 = fn5D0("&Mail", "Read &Incoming Mail")
Else
l00B6 = fn5D0("&Mail", "Read &Incoming FlashMail")
End If
MsgBox "Please wait for your mail to be full!", 16
GoTo 13
End If
End If
l00BC = extfnFE0(l00B0, 1036, 0, 0)
For l00C0 = 0 To l00BC - 1
l00C4$ = String(255, 0)
l00C8 = extfn11A0(l00B0, 1034, l00C0, l00C4$)
List2.AddItem l00C4$
Next l00C0
If  l00BC > 1 Then
Panel3D6.Caption = "Incoming FlashMail Contains " + l00BC + " Mails"
Else
Panel3D6.Caption = "Incoming FlashMail Contains" + l00BC + " Mail"
End If
Case "Outgoing FlashMail"
Panel3D5.Caption = "Contents of Outgoing FlashMail"
List2.Clear
l0098 = extfn1018("AOL FRAME25", 0&)
l009E = extfn10F8(l0098, "Welcome")
l00A2$ = String(30, 0)
l00A6 = extfnFA8(l009E, l00A2$, 250)
If  l00A6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
12 :
l008C = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l008C, "Outgoing Mail")
l00B0 = extfn10C0(l00AA, "_AOL_TREE")
If  l00AA = 0 Then
If  l00B0 = 0 Then
If  fn138() = 5 Then
l00B6 = fn5D0("&Mail", "Read &Outgoing Mail")
Else
l00B6 = fn5D0("&Mail", "Read &Outgoing FlashMail")
End If
MsgBox "Please wait for your mail to be full!", 16
GoTo 12
End If
End If
l00BC = extfnFE0(l00B0, 1036, 0, 0)
For l00C0 = 0 To l00BC - 1
l00C4$ = String(255, 0)
l00C8 = extfn11A0(l00B0, 1034, l00C0, l00C4$)
List2.AddItem l00C4$
Next l00C0
If  l00BC > 1 Then
Panel3D6.Caption = "Outgoing FlashMail Contains " + l00BC + " Mails"
Else
Panel3D6.Caption = "Outgoing FlashMail Contains " + l00BC + " Mail"
End If
Case "Personal Filing Cabinet"
Panel3D5.Caption = "Contents of Personal Filing Cabinet"
List2.Clear
11 :
l008C = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l008C, "Personal Filing Cabinet")
l00B0 = extfn10C0(l00AA, "_AOL_TREE")
If  l00AA = 0 Then
If  l00B0 = 0 Then
l00B6 = fn5D0("&Mail", "P&ersonal Filing Cabinet")
MsgBox "Please wait for your mail to be full!", 16
GoTo 11
End If
End If
l00BC = extfnFE0(l00B0, 1036, 0, 0)
For l00C0 = 0 To l00BC - 1
l00C4$ = String(255, 0)
l00C8 = extfn11A0(l00B0, 1034, l00C0, l00C4$)
List2.AddItem l00C4$
Next l00C0
If  l00BC > 1 Then
Panel3D6.Caption = "Personal Filing Cabinet Contains " + l00BC + " Mails"
Else
Panel3D6.Caption = "Personal Filing Cabinet Contains " + l00BC + " Mail"
End If
End Select
End Sub

Sub Option3D2_Click ()
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As String
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As String
Dim l00F4 As Variant
List1.Clear
l00D0 = extfn1018("AOL FRAME25", 0&)
l00D4 = extfn10F8(l00D0, "Welcome")
l00D8$ = String(30, 0)
l00DA = extfnFA8(l00D4, l00D8$, 250)
If  l00DA <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00DE = extfn1018("AOL FRAME25", 0&)
l00E2 = extfn10F8(l00DE, "Old Mail")
l00E6 = extfn10C0(l00E2, "_AOL_TREE")
If  l00E2 = 0 Then
If  l00E6 = 0 Then
MsgBox "Old Mail Could Not Be Found", 16
Exit Sub
End If
End If
l00EA = extfnFE0(l00E6, 1036, 0, 0)
For l00EE = 0 To l00EA - 1
l00F2$ = String(255, 0)
l00F4 = extfn11A0(l00E6, 1034, l00EE, l00F2$)
List1.AddItem l00F2$
Next l00EE
End Sub

Sub Option3D3_Click ()
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0106 As Variant
Dim l010A As Variant
Dim l010E As String
Dim l0110 As Variant
List1.Clear
l00FA = extfn1018("AOL FRAME25", 0&)
l00FE = extfn10F8(l00FA, "Personal Filing Cabinet")
l0102 = extfn10C0(l00FE, "_AOL_TREE")
If  l00FE = 0 Then
If  l0102 = 0 Then
MsgBox "Personal Filing Cabinet Could Not Be Found", 16
Exit Sub
End If
End If
l0106 = extfnFE0(l0102, 1036, 0, 0)
For l010A = 0 To l0106 - 1
l010E$ = String(255, 0)
l0110 = extfn11A0(l0102, 1034, l010A, l010E$)
List1.AddItem l010E$
Next l010A
End Sub

Sub Option3D4_Click ()
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As String
Dim l0120 As Variant
Dim l0124 As Variant
Dim l0128 As Variant
Dim l012C As Variant
Dim l0130 As Variant
Dim l0134 As Variant
Dim l0138 As String
Dim l013A As Variant
List1.Clear
l0116 = extfn1018("AOL FRAME25", 0&)
l011A = extfn10F8(l0116, "Welcome")
l011E$ = String(30, 0)
l0120 = extfnFA8(l011A, l011E$, 250)
If  l0120 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0124 = extfn1018("AOL FRAME25", 0&)
l0128 = extfn10F8(l0124, "Outgoing Mail")
l012C = extfn10C0(l0128, "_AOL_TREE")
If  l0128 = 0 Then
If  l012C = 0 Then
MsgBox "Outgoing Mail Could Not Be Found", 16
Exit Sub
End If
End If
l0130 = extfnFE0(l012C, 1036, 0, 0)
For l0134 = 0 To l0130 - 1
l0138$ = String(255, 0)
l013A = extfn11A0(l012C, 1034, l0134, l0138$)
List1.AddItem l0138$
Next l0134
End Sub

Sub sa_Click ()
Dim l013E As Variant
For l013E = 0 To List2.ListCount - 1
List2.Selected(l013E) = True
Next l013E
End Sub

Sub sav_Click ()
Dim l0142 As Variant
Dim l0146 As Variant
Dim l014A As Variant
Dim l0152 As Variant
l0142 = 0
l0146 = InputBox("Enter in the name of the file to save to:", "Save Mailbox", "c:\mailbox.txt")
If  l0146 = "" Then
Exit Sub
End If
For l014A = 0 To List2.ListCount - 1
If  List2.Selected(l014A) = True Then
l0142 = l0142 + 1
End If
Next l014A
If  l0142 = 0 Then
MsgBox "There is nothing to put in the file! Please select mails to save or select all to save all!", 16
Exit Sub
End If
pwait.Show
For l0152 = 0 To List2.ListCount - 1
If  List2.Selected(l0152) = True Then
Text1.Text = Text1.Text + List2.List(l0152) + Chr(13) + Chr(10)
End If
Next l0152
On Error Resume Next
Open l0146 For Output As #1
If  Err Then
MsgBox "An error occured: Invalid Directory"
Exit Sub
End If
Print #1, Text1.Text
Close #1
Text1.Text = ""
Unload pwait
MsgBox "Mailbox has been saved!", 16
End Sub

Sub usa_Click ()
Dim l015C As Variant
For l015C = 0 To List2.ListCount - 1
List2.Selected(l015C) = False
Next l015C
End Sub

Sub sub9980 ()
Unload multimb
End Sub
