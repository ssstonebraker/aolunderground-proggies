' MAILBOX.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l002E As Variant
Dim l0036 As Variant
Dim l003C As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004C As String
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As String
Dim l005E As Variant
Dim l0062 As Variant
Dim l006C As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l002E = extfn1018("AOL FRAME25", 0&)
l001C = extfn10F8(l002E, "AOL Live!")
l001C = extfn1088(l001C)
If  l001C = 0 Then
MsgBox "Room could not be found!", 16
Exit Sub
End If
l0036 = extfn10C0(l001C, "_AOL_listbox")
l003C = extfnFE0(l0036, 1036, 0, 0)
For l0040 = 0 To l003C - 1
DoEvents
l0044 = extfn1018("AOL FRAME25", 0&)
l0048 = extfn10F8(l0044, "Welcome")
l004C$ = String(30, 0)
l004E = extfnFA8(l0048, l004C$, 250)
l0052 = l004E - 1
On Error Resume Next
l0056 = Mid(Left(l004C$, l0052), 10)
If  Err Then
Exit Sub
End If
l005A$ = String(30, 0)
l005E = extfn1608(l0040, l005A$)
l0062 = InStr(l005A$, Chr(0))
If  l0062 Then
l005A$ = Mid(l005A$, 1, l0062 - 1)
End If
l001C = InStr(1, Text3.Text, l005A$, 1)
If  l001C Then
l0022 = Mid(Text3.Text, l001C + Len(l005A$), 1)
If  l0022 = "," Then GoTo 3003
End If
l001C = InStr(1, l005A$, "Guide", 1)
If  l001C Then GoTo 3003
l006C = InStr(1, l005A$, l0056, 1)
If  l006C Then GoTo 3003
List1.AddItem l005A$
Text3.Text = Text3.Text + l005A$ + ","
3003 :
Next l0040
End Sub

Sub Command3D2_Click ()
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As String
Dim l008C As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A6 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
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
Dim l012C As Variant
Dim l0130 As Variant
Dim l0134 As Variant
Dim l0138 As Variant
Dim l013C As Variant
Dim l0140 As Variant
Dim l0144 As Variant
Dim l0148 As Variant
Dim l014C As Variant
Dim l0150 As Variant
Dim l0154 As Variant
Dim l0158 As Variant
Dim l015C As Variant
Dim l0160 As Variant
Dim l0164 As Variant
Dim l0168 As Variant
Dim l016C As Variant
Dim l0170 As Variant
Dim l0174 As Variant
Dim l0178 As Variant
Dim l017C As Variant
Dim l0180 As Variant
Dim l0184 As Variant
If  List1.ListCount = 0 Then
MsgBox "Nobody to send!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "Please enter in a subject!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "Please enter in a message!", 16
Exit Sub
End If
l0082 = extfn1018("AOL FRAME25", 0&)
l0086 = extfn10F8(l0082, "Welcome")
l008A$ = String(30, 0)
l008C = extfnFA8(l0086, l008A$, 250)
If  l008C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  fn138() = 5 Then
l0094 = fn5D0("Mem&bers", "Set Preferences")
Else
l0094 = fn5D0("Mem&bers", "Preferences")
End If
100 :
DoEvents
l0098 = extfn1018("AOL FRAME25", 0&)
l009C = extfn10F8(l0098, "Preferences")
If  l009C = 0 Then GoTo 100
l00A0 = extfn10F8(l009C, "Mail")
l00A6 = extfn1168(l00A0, 2)
l00AC = fn1480(l00A6)
200 :
DoEvents
l00B0 = extfn1018("_AOL_MODAL", 0&)
If  l00B0 = 0 Then GoTo 200
l00B4 = extfn10F8(l00B0, "close")
l00B8 = extfn1168(l00B4, 3)
l00BC = extfnFE0(l00B8, 1025, 1, 0)
l00C0 = extfn10F8(l00B0, "confirm")
l00C4 = extfn1168(l00C0, 2)
l00C8 = extfnFE0(l00C4, 1025, 1, 0)
l00CC = extfn10F8(l00B0, "OK")
l00D0 = fn1480(l00CC)
l00D4 = extfnFE0(l009C, 16, 0, 0)
l00D8 = fn5D0("&Mail", "&Compose Mail")
l00DC = extfn1018("AOL FRAME25", 0&)
10 :
DoEvents
l00D8 = extfn10F8(l00DC, "Compose Mail")
If  l00D8 = 0 Then GoTo 10
l00E0 = extfn10F8(l00D8, "To:")
l00E4 = extfn1168(l00E0, 2)
l00EA = extfn11A0(l00E4, 12, 0, Text3.Text)
l00EE = extfn1168(l00E0, 1)
l00F2 = extfn11A0(l00EE, 12, 0, Text2.Text)
l00F6 = extfn10F8(l00D8, "Subject:")
l00FA = extfn1168(l00F6, 2)
l00FE = extfn11A0(l00FA, 12, 0, Text1.Text)
l0102 = extfn10F8(l00D8, "Send")
l0106 = extfn1168(l0102, 2)
l010A = fn1480(l0106)
20 :
DoEvents
l010E = fn138()
If  l010E = 5 Then
l0112 = extfn1018("#32770", "America Online")
Else
l0112 = extfn1018("_AOL_MODAL", 0&)
End If
l0116 = extfn10F8(l00DC, "Error")
If  l0116 > 0 Then
l011A = extfn10C0(l0116, "_AOL_VIEW")
l011E$ = String(255, 0)
l0120 = extfnFE0(l011A, 14, 0, 0)
l0124 = extfn11A0(l011A, 13, l0120, l011E$)
Text3 = l011E$
l0128 = InStr(Text3, Chr(10))
l012C = Mid(Text3, l0128 + 3, Len(Text3))
Text3 = l012C
For l0130 = 1 To Len(Text3)
l0134 = InStr(Text3, "-")
If  l0134 Then
l0138 = Mid(Text3, 1, l0134 - 2)
For l013C = 0 To List1.ListCount
l0140 = InStr(1, List1.List(l013C), l0138, 1)
If  l0140 Then
l0144 = Len(l0138)
l0148 = Len(List1.List(l013C))
If  l0144 <> l0148 Then
GoTo 87
End If
List1.RemoveItem l013C
l0094 = InStr(1, Text3.Text, l013C + ",", 1)
If  l0094 Then
l011A = Mid(Text3.Text, 1, l0094 - 1)
l014C = Mid(Text3.Text, l0094 + Len(l013C) + 1, Len(Text3.Text))
Text3.Text = l011A + l014C
End If
GoTo 35
End If
87 :
Next l013C
35 :
l0150 = InStr(Text3, Chr(10))
If  l0150 Then
l0154 = Mid(Text3, l0150 + 1, Len(Text3))
Text3 = l0154
Else
GoTo 243
End If
Else
GoTo 243
End If
Next l0130
243 :
Text3 = ""
If  List1.ListCount = 0 Then
MsgBox "An error occurred: Nobody to mail", 16
Exit Sub
End If
For l009C = 0 To List1.ListCount
Text3 = Text3 + "," + List1.List(l009C)
Next l009C
l0158 = extfn1018("AOL FRAME25", 0&)
l015C = extfn10F8(l0158, "Error")
l0160 = extfn10F8(l015C, "OK")
l0164 = fn1480(l0160)
l0168 = extfn10F8(l0158, "Compose Mail")
l016C = extfn10F8(l0168, "To:")
l0170 = extfn1168(l016C, 2)
l0174 = extfn11A0(l0170, 12, 0, Text3)
l0178 = extfn10F8(l0168, "Send")
l013C = extfn1168(l0178, 2)
l017C = fn1480(l013C)
End If
If  l0112 = 0 Then GoTo 20
l0180 = extfn10F8(l0112, "OK")
l0184 = fn1480(l0180)
End Sub

Sub Command3D3_Click ()
List1.Clear
Text3.Text = ""
End Sub

Sub Command3D4_Click ()
Dim l0188 As Variant
Dim l018C As Variant
Dim l0190 As Variant
l0188 = InputBox("Enter in name to add:", "Add Name", "")
If  l0188 = "" Then
Exit Sub
End If
If  Len(l0188) > 10 Then
MsgBox "Invalid Screen Name!", 16
Exit Sub
End If
For l018C = 0 To List1.ListCount
l0190 = InStr(1, List1.List(l018C), l0188, 1)
If  l0190 Then
If  Len(List1.List(l018C)) = Len(l0188) Then
l0188 = ""
GoTo 201
End If
End If
Next l018C
List1.AddItem l0188
Text3.Text = Text3.Text + l0188 + ","
l0188 = ""
201 :
End Sub

Sub Command3D5_Click ()
Dim l0194 As Variant
Dim l0198 As Variant
Dim l019C As Variant
Dim l01A0 As Variant
Dim l01A4 As Variant
l0194 = List1.ListIndex
l0198 = List1.List(l0194)
On Error Resume Next
List1.RemoveItem l0194
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l019C = InStr(1, Text3.Text, l0198 + ",", 1)
If  l019C Then
l01A0 = Mid(Text3.Text, 1, l019C - 1)
l01A4 = Mid(Text3.Text, l019C + Len(l0198) + 1, Len(Text3.Text))
Text3.Text = l01A0 + l01A4
End If
End If
End Sub

Sub Form_Load ()
extsub1050 MAILBOX.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub sub9980 ()
Unload MAILBOX
End Sub
