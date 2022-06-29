' MAILLIST.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0026 As Variant
Dim l003E As Variant
Dim l0044 As Variant
Dim l0048 As String
Dim l004C As Variant
Dim l0054 As Variant
Dim l0058 As Variant
Dim l005E As Variant
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As String
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
If  List1.ListCount = 0 Then
MsgBox "There is nobody to send the list to!", 16
Exit Sub
End If
If  Option3D1.Value = True Then l0026 = "New Mail"
If  Option3D2.Value = True Then l0026 = "Old Mail"
If  Option3D3.Value = True Then l0026 = "Outgoing Mail"
If  Option3D4.Value = True Then l0026 = "Incoming FlashMail"
l003E = extfn1018("AOL FRAME25", 0&)
l0044 = extfn10F8(l003E, "Welcome")
l0048$ = String(30, 0)
l004C = extfnFA8(l0044, l0048$, 250)
If  l004C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
pwait.Show
pwait.Label1.Caption = "Please Wait..."
l0054 = extfn1018("AOL FRAME25", 0&)
l0058 = extfn10F8(l0054, l0026)
l005E = extfn10C0(l0058, "_AOL_TREE")
If  l0058 = 0 Then
If  l005E = 0 Then
If  fn138() = 5 Then
If  Option3D1.Value = True Then l0054 = fn5D0("&Mail", "Read &New Mail")
If  Option3D2.Value = True Then l0054 = fn5D0("&Mail", "Check Mail You've &Read")
If  Option3D3.Value = True Then l0054 = fn5D0("&Mail", "Check Mail You've &Sent")
If  Option3D4.Value = True Then l0054 = fn5D0("&Mail", "Read &Incoming Mail")
Else
If  Option3D1.Value = True Then l0054 = fn5D0("&Mail", "Read &New Mail")
If  Option3D2.Value = True Then l0054 = fn5D0("&Mail", "Check Mail You've &Read")
If  Option3D3.Value = True Then l0054 = fn5D0("&Mail", "Check Mail You've &Sent")
If  Option3D4.Value = True Then l0054 = fn5D0("&Mail", "Read &Incoming FlashMail")
End If
pwait.Text2.Text = Text4.Text
pwait.Text1.Text = Text1.Text
pwait.Command3D1.Visible = True
pwait.Panel3D2.Visible = True
pwait.Label1.Caption = "Click Continue When Mailbox Is Ready"
Exit Sub
End If
End If
l0058 = extfn10F8(l0054, l0026)
l005E = extfn10C0(l0058, "_AOL_TREE")
l0074 = extfnFE0(l005E, 1036, 0, 0)
pwait.Label1.Caption = "Collecting Mailbox..."
DoEvents
For l0078 = 0 To l0074 - 1
DoEvents
l007C$ = String(255, 0)
l0080 = extfn11A0(l005E, 1034, l0078, l007C$)
l0084 = InStr(l007C$, Chr(9))
l0088 = Mid(l007C$, l0084 + 1, Len(l007C$))
l008C = InStr(l0088, Chr(9))
l007C$ = Mid(l0088, l008C + 1, Len(l0088))
l0090 = Chr(13) + Chr(10)
Text1 = Text1 + l0090 + Str(l0078) + ".) " + l007C$
Next l0078
l0094 = fn5D0("&Mail", "&Compose Mail")
40 :
DoEvents
l0098 = extfn10F8(l0054, "Compose Mail")
If  l0098 = 0 Then GoTo 40
pwait.Label1.Caption = "Setting information..."
l009C = extfn10F8(l0098, "To:")
l00A2 = extfn1168(l009C, 2)
l00A6 = extfn11A0(l00A2, 12, 0, Text4.Text)
l00AA = extfn10F8(l0098, "Subject:")
l00AE = extfn1168(l00AA, 2)
l00B2 = extfn11A0(l00AE, 12, 0, "«‹- -– Fate Ultra 9.0 Mail Box List - " + l0026 + " –-- -›»")
l00B6 = extfn1168(l009C, 1)
l003E = extfn11A0(l00B6, 12, 0, Text1.Text)
DoEvents
l0044 = extfn10F8(l0098, "Send")
l00BA = extfn1168(l0044, 2)
l004C = fn1480(l00BA)
Unload pwait
End Sub

Sub Command3D2_Click ()
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As String
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00E8 As String
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As String
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l010C As Variant
Dim l0110 As Variant
l00C0 = extfn1018("AOL FRAME25", 0&)
l00C4 = extfn10F8(l00C0, "Welcome")
l00C8$ = String(30, 0)
l00CA = extfnFA8(l00C4, l00C8$, 250)
If  l00CA <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00CE = extfn1018("AOL FRAME25", 0&)
l00C0 = extfn10F8(l00CE, "AOL Live!")
l00C0 = extfn1088(l00C0)
If  l00C0 = 0 Then
MsgBox "Room could not be found!", 16
Exit Sub
End If
l00D4 = extfn10C0(l00C0, "_AOL_listbox")
l00D8 = extfnFE0(l00D4, 1036, 0, 0)
For l00DC = 0 To l00D8 - 1
DoEvents
l00E0 = extfn1018("AOL FRAME25", 0&)
l00E4 = extfn10F8(l00E0, "Welcome")
l00E8$ = String(30, 0)
l00EA = extfnFA8(l00E4, l00E8$, 250)
l00EE = l00EA - 1
On Error Resume Next
l00F2 = Mid(Left(l00E8$, l00EE), 10)
If  Err Then
Exit Sub
End If
l00F6$ = String(30, 0)
l00FA = extfn1608(l00DC, l00F6$)
l00FE = InStr(l00F6$, Chr(0))
If  l00FE Then
l00F6$ = Mid(l00F6$, 1, l00FE - 1)
End If
l00C0 = InStr(1, Text4.Text, l00F6$, 1)
If  l00C0 Then
l00C4 = Mid(Text4.Text, l00C0 + Len(l00F6$), 1)
If  l00C4 = "," Then GoTo 3003
End If
l00C0 = InStr(1, l00F6$, "Guide", 1)
If  l00C0 Then GoTo 3003
l0102 = InStr(1, l00F6$, l00F2, 1)
If  l0102 Then GoTo 3003
List1.AddItem l00F6$
Text4.Text = Text4.Text + l00F6$ + ","
l010C = Mid(Frame3D1.Caption, 10, Len(Frame3D1.Caption))
l0110 = Int(l010C) + 1
Frame3D1.Caption = "Send To: " + Str(l0110)
3003 :
Next l00DC
End Sub

Sub Command3D3_Click ()
Dim l0114 As Variant
Dim l0118 As Variant
Dim l011C As Variant
Dim l0120 As Variant
Dim l0124 As Variant
l0114 = List1.ListIndex
l0118 = List1.List(l0114)
On Error Resume Next
List1.RemoveItem l0114
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l011C = InStr(1, Text4.Text, l0118 + ",", 1)
If  l011C Then
l0120 = Mid(Text4.Text, 1, l011C - 1)
l0124 = Mid(Text4.Text, l011C + Len(l0118) + 1, Len(Text4.Text))
Text4.Text = l0120 + l0124
l0114 = Mid(Frame3D1.Caption, 10, Len(Frame3D1))
l0120 = l0114 - 1
Frame3D1.Caption = "Send To: " + l0120
End If
End If
End Sub

Sub Command3D4_Click ()
Frame3D1.Caption = "Send To: 0"
List1.Clear
Text4.Text = ""
End Sub

Sub Command3D5_Click ()
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
If  Text3.Text = "" Then GoTo 201
For l012E = 0 To List1.ListCount
l0132 = InStr(1, List1.List(l012E), Text3.Text, 1)
If  l0132 Then
If  Len(List1.List(l012E)) = Len(Text3.Text) Then
Text3.Text = ""
GoTo 201
End If
End If
Next l012E
List1.AddItem Text3
Text4.Text = Text4.Text + Text3.Text + ","
Text3 = ""
l0136 = Mid(Frame3D1.Caption, 10, Len(Frame3D1.Caption))
l013A = Int(l0136) + 1
Frame3D1.Caption = "Send To: " + Str(l013A)
201 :
End Sub

Sub Form_Load ()
extsub1050 maillist.hWnd, -1, 0, 0, 0, 0, 3
Option3D1.Value = True
End Sub

Sub Text3_KeyPress (KeyAscii As Integer)
Dim l0144 As Variant
Dim l0148 As Variant
Dim l014C As Variant
Dim l0150 As Variant
l0144 = Chr(KeyAscii)
If  l0144 = Chr(13) Then
If  Text3.Text = "" Then GoTo 202
For l0144 = 0 To List1.ListCount
l0148 = InStr(1, List1.List(l0144), Text3.Text, 1)
If  l0148 Then
If  Len(List1.List(l0144)) = Len(Text3.Text) Then
Text3.Text = ""
GoTo 202
End If
End If
Next l0144
List1.AddItem Text3
Text4.Text = Text4.Text + Text3.Text + ","
Text3 = ""
l014C = Mid(Frame3D1.Caption, 10, Len(Frame3D1.Caption))
l0150 = Int(l014C) + 1
Frame3D1.Caption = "Send To: " + Str(l0150)
202 :
End If
End Sub
