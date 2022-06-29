' MMOPTION.FRM
Option Explicit

Sub Check3D3_Click (Value As Integer)
If  Check3D3.Value = True Then
Text1.Enabled = True
Else
Text1.Enabled = False
End If
End Sub

Sub Check3D4_Click (Value As Integer)
If  Check3D4.Value = True Then
Text2.Enabled = True
Else
Text2.Enabled = False
End If
End Sub

Sub Command3D1_Click ()
Dim l0036 As Variant
Dim l0040 As Variant
l0036 = 0
If  Option3D7.Value = True Then
For l0040 = 0 To List1.ListCount - 1
If  List1.Selected(l0040) = True Then l0036 = l0036 + 1
Next l0040
If  l0036 = 0 Then
MsgBox "Please select mails that you want to send!", 16
Exit Sub
End If
End If
mmoptions.Hide
End Sub

Sub Command3D2_Click ()
Dim l004E As Variant
For l004E = 0 To List1.ListCount - 1
List1.Selected(l004E) = True
Next l004E
End Sub

Sub Command3D3_Click ()
Dim l0052 As Variant
For l0052 = 0 To List1.ListCount - 1
List1.Selected(l0052) = False
Next l0052
End Sub

Sub Command3D4_Click ()
Dim l0056 As Variant
Dim l005A As Variant
Dim l0064 As Variant
For l0056 = 0 To List1.ListCount - 1
l005A = List1.List(l0056)
l0064 = InStr(1, l005A, Text3.Text, 1)
If  l0064 Then
List1.Selected(l0056) = True
End If
Next l0056
End Sub

Sub Form_Load ()
extsub1050 mmoptions.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Option3D1_Click (Value As Integer)
Text6.Enabled = False
mmoptions.Height = 2445
List1.Clear
List2.Clear
End Sub

Sub Option3D2_Click (Value As Integer)
Text6.Enabled = True
mmoptions.Height = 2445
List1.Clear
List2.Clear
End Sub

Sub Option3D3_Click (Value As Integer)
Dim l0082 As Variant
Dim l0088 As Variant
Dim l008E As Variant
Dim l0094 As Variant
Dim l009A As Variant
Dim l009E As Variant
Dim l00A2 As String
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Check3D1.Enabled = True
Check3D1.Value = True
If  Option3D7.Value = True Then
List1.Clear
l0082 = extfn1018("AOL FRAME25", 0&)
l0088 = extfn10F8(l0082, "New Mail")
If  l0088 > 0 Then GoTo 366
l008E = fn5D0("&Mail", "Read &New Mail")
MsgBox "Please Wait For Your Mail To Fill Up!", 16
366 :
l0082 = extfn1018("AOL FRAME25", 0&)
l0088 = extfn10F8(l0082, "New Mail")
l0094 = extfn10C0(l0088, "_AOL_TREE")
l009A = extfnFE0(l0094, 1036, 0, 0)
For l009E = 0 To l009A - 1
l00A2$ = String(255, 0)
l00A6 = extfn11A0(l0094, 1034, l009E, l00A2$)
l00AA = InStr(l00A2$, Chr(0))
If  l00AA Then
l00A2$ = Mid(l00A2$, 1, l00AA - 1)
End If
l00AE = InStr(l00A2$, Chr(9))
If  l00AE Then
l00A2$ = Mid(l00A2$, l00AE + 1, Len(l00A2$))
End If
l00AE = InStr(l00A2$, Chr(9))
If  l00AE Then
l00A2$ = Mid(l00A2$, l00AE + 1, Len(l00A2$))
End If
List1.AddItem Str(l009E) + ") " + l00A2$
Next l009E
End If
End Sub

Sub Option3D4_Click (Value As Integer)
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As String
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Check3D1.Enabled = False
Check3D1.Value = False
If  Option3D7.Value = True Then
List1.Clear
l00B4 = extfn1018("AOL FRAME25", 0&)
l00B8 = extfn10F8(l00B4, "Outgoing FlashMail")
If  l00B8 > 0 Then GoTo 365
If  fn138() = 5 Then
l00BE = fn5D0("&Mail", "Read &Incoming Mail")
Else
l00BE = fn5D0("&Mail", "Read &Incoming FlahsMail")
End If
MsgBox "Please Wait For Your Mail To Fill Up!", 16
365 :
l00B4 = extfn1018("AOL FRAME25", 0&)
l00B8 = extfn10F8(l00B4, "Outgoing FlashMail")
l00C2 = extfn10C0(l00B8, "_AOL_TREE")
l00C6 = extfnFE0(l00C2, 1036, 0, 0)
For l00CA = 0 To l00C6 - 1
l00CE$ = String(255, 0)
l00D0 = extfn11A0(l00C2, 1034, l00CA, l00CE$)
l00D4 = InStr(l00CE$, Chr(0))
If  l00D4 Then
l00CE$ = Mid(l00CE$, 1, l00D4 - 1)
End If
l00D8 = InStr(l00CE$, Chr(9))
If  l00D8 Then
l00CE$ = Mid(l00CE$, l00D8 + 1, Len(l00CE$))
End If
l00D8 = InStr(l00CE$, Chr(9))
If  l00D8 Then
l00CE$ = Mid(l00CE$, l00D8 + 1, Len(l00CE$))
End If
List1.AddItem Str(l00CA) + ") " + l00CE$
Next l00CA
End If
End Sub

Sub Option3D5_Click (Value As Integer)
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As String
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
Check3D1.Enabled = False
Check3D1.Value = False
If  Option3D7.Value = True Then
List1.Clear
l00DE = extfn1018("AOL FRAME25", 0&)
l00E2 = extfn10F8(l00DE, "Old Mail")
If  l00E2 > 0 Then GoTo 367
l00E6 = fn5D0("&Mail", "Check Mail You've &Read")
MsgBox "Please Wait For Your Mail To Fill Up!", 16
367 :
l00DE = extfn1018("AOL FRAME25", 0&)
l00E2 = extfn10F8(l00DE, "Old Mail")
l00EA = extfn10C0(l00E2, "_AOL_TREE")
l00EE = extfnFE0(l00EA, 1036, 0, 0)
For l00F2 = 0 To l00EE - 1
l00F6$ = String(255, 0)
l00F8 = extfn11A0(l00EA, 1034, l00F2, l00F6$)
l00FC = InStr(l00F6$, Chr(0))
If  l00FC Then
l00F6$ = Mid(l00F6$, 1, l00FC - 1)
End If
l0100 = InStr(l00F6$, Chr(9))
If  l0100 Then
l00F6$ = Mid(l00F6$, l0100 + 1, Len(l00F6$))
End If
l0100 = InStr(l00F6$, Chr(9))
If  l0100 Then
l00F6$ = Mid(l00F6$, l0100 + 1, Len(l00F6$))
End If
List1.AddItem Str(l00F2) + ") " + l00F6$
Next l00F2
End If
End Sub

Sub Option3D6_Click (Value As Integer)
Dim l0106 As Variant
Dim l010A As Variant
Dim l010E As Variant
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As String
Dim l0120 As Variant
Dim l0124 As Variant
Dim l0128 As Variant
Check3D1.Enabled = False
Check3D1.Value = False
If  Option3D7.Value = True Then
List1.Clear
l0106 = extfn1018("AOL FRAME25", 0&)
l010A = extfn10F8(l0106, "Outgoing Mail")
If  l010A > 0 Then GoTo 368
l010E = fn5D0("&Mail", "Check Mail You've &Sent")
MsgBox "Please Wait For Your Mail To Fill Up!", 16
368 :
l0106 = extfn1018("AOL FRAME25", 0&)
l010A = extfn10F8(l0106, "Outgoing Mail")
l0112 = extfn10C0(l010A, "_AOL_TREE")
l0116 = extfnFE0(l0112, 1036, 0, 0)
For l011A = 0 To l0116 - 1
l011E$ = String(255, 0)
l0120 = extfn11A0(l0112, 1034, l011A, l011E$)
l0124 = InStr(l011E$, Chr(0))
If  l0124 Then
l011E$ = Mid(l011E$, 1, l0124 - 1)
End If
l0128 = InStr(l011E$, Chr(9))
If  l0128 Then
l011E$ = Mid(l011E$, l0128 + 1, Len(l011E$))
End If
l0128 = InStr(l011E$, Chr(9))
If  l0128 Then
l011E$ = Mid(l011E$, l0128 + 1, Len(l011E$))
End If
List1.AddItem Str(l011A) + ") " + l011E$
Next l011A
End If
End Sub

Sub Option3D7_Click (Value As Integer)
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As String
Dim l013A As Variant
Dim l014A As Variant
Dim l0160 As Variant
Dim l0164 As Variant
Dim l0168 As Variant
Dim l016C As Variant
Dim l0170 As Variant
Dim l0174 As Variant
Dim l0178 As String
Dim l017A As Variant
Dim l017E As Variant
Dim l0182 As Variant
l012E = extfn1018("AOL FRAME25", 0&)
l0132 = extfn10F8(l012E, "Welcome")
l0136$ = String(30, 0)
l013A = extfnFA8(l0132, l0136$, 250)
If  l013A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Option3D1.Value = True
Exit Sub
End If
Text6.Enabled = False
mmoptions.Height = 5365
If  Option3D3.Value = True Then l014A = "New Mail"
If  Option3D4.Value = True Then l014A = "Incoming FlashMail"
If  Option3D5.Value = True Then l014A = "Old Mail"
If  Option3D6.Value = True Then l014A = "Outgoing Mail"
l0160 = extfn1018("AOL FRAME25", 0&)
l0164 = extfn10F8(l0160, l014A)
If  l0164 > 0 Then GoTo 345
If  fn138() = 5 Then
If  mmoptions.Option3D3.Value = True Then l0168 = fn5D0("&Mail", "Read &New Mail")
If  mmoptions.Option3D4.Value = True Then l0168 = fn5D0("&Mail", "Read &Incoming Mail")
If  mmoptions.Option3D5.Value = True Then l0168 = fn5D0("&Mail", "Check Mail You've &Read")
If  mmoptions.Option3D6.Value = True Then l0168 = fn5D0("&Mail", "Check Mail You've &Sent")
Else
If  mmoptions.Option3D3.Value = True Then l0168 = fn5D0("&Mail", "Read &New Mail")
If  mmoptions.Option3D4.Value = True Then l0168 = fn5D0("&Mail", "Read &Incoming FlashMail")
If  mmoptions.Option3D5.Value = True Then l0168 = fn5D0("&Mail", "Check Mail You've &Read")
If  mmoptions.Option3D6.Value = True Then l0168 = fn5D0("&Mail", "Check Mail You've &Sent")
End If
MsgBox "Please Wait For Your Mail To Fill Up!", 16
345 :
l0160 = extfn1018("AOL FRAME25", 0&)
l0164 = extfn10F8(l0160, l014A)
l016C = extfn10C0(l0164, "_AOL_TREE")
l0170 = extfnFE0(l016C, 1036, 0, 0)
For l0174 = 0 To l0170 - 1
l0178$ = String(255, 0)
l017A = extfn11A0(l016C, 1034, l0174, l0178$)
l017E = InStr(l0178$, Chr(0))
If  l017E Then
l0178$ = Mid(l0178$, 1, l017E - 1)
End If
l0182 = InStr(l0178$, Chr(9))
If  l0182 Then
l0178$ = Mid(l0178$, l0182 + 1, Len(l0178$))
End If
l0182 = InStr(l0178$, Chr(9))
If  l0182 Then
l0178$ = Mid(l0178$, l0182 + 1, Len(l0178$))
End If
List1.AddItem Str(l0174) + ") " + l0178$
Next l0174
End Sub

Sub Text6_Change ()
If  Not IsNumeric(Text6.Text) Then
MsgBox "The mails that you want to mass mail must be in numeric form, from 1 to 550!", 16
Text6.Text = "100"
End If
If  Int(Text6.Text) > 550 Then
MsgBox "You can mass mail more than 550 mails!", 16
Text6.Text = "100"
Exit Sub
End If
If  Int(Text6.Text) <= 0 Then
MsgBox "You can mass mail less than 1 mail!", 16
Text6.Text = "100"
Exit Sub
End If
End Sub
