' IM.FRM
Option Explicit

Sub sub7B98 ()
End Sub

Sub Command3D1_Click ()
Unload imform
End Sub

Sub Command3D2_Click ()
Dim l0020 As Variant
Dim l0026 As Variant
Dim l002A As String
Dim l002E As Variant
Dim l0032 As Variant
Dim l003A As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As String
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As String
Dim l0062 As Variant
Dim l0066 As Variant
Dim l0070 As Variant
Dim l0080 As Variant
Dim l0084 As Variant
l0020 = extfn1018("AOL FRAME25", 0&)
l0026 = extfn10F8(l0020, "Welcome")
l002A$ = String(30, 0)
l002E = extfnFA8(l0026, l002A$, 250)
If  l002E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0032 = extfn1018("AOL FRAME25", 0&)
l0020 = extfn10F8(l0032, "AOL Live!")
l0020 = extfn1088(l0020)
If  l0020 = 0 Then
MsgBox "Room List not found!", 16
Exit Sub
End If
l003A = extfn10C0(l0020, "_AOL_listbox")
l0040 = extfnFE0(l003A, 1036, 0, 0)
For l0044 = 0 To l0040 - 1
DoEvents
l0048 = extfn1018("AOL FRAME25", 0&)
l004C = extfn10F8(l0048, "Welcome")
l0050$ = String(30, 0)
l0052 = extfnFA8(l004C, l0050$, 250)
l0056 = l0052 - 1
On Error Resume Next
l005A = Mid(Left(l0050$, l0056), 10)
If  Err Then
Exit Sub
End If
l005E$ = String(30, 0)
l0062 = extfn1608(l0044, l005E$)
l0066 = InStr(l005E$, Chr(0))
If  l0066 Then
l005E$ = Mid(l005E$, 1, l0066 - 1)
End If
l0020 = InStr(1, Text3.Text, l005E$, 1)
If  l0020 Then
l0026 = Mid(Text3.Text, l0020 + Len(l005E$), 1)
If  l0026 = "," Then GoTo 3003
End If
l0020 = InStr(1, l005E$, "Guide", 1)
If  l0020 Then GoTo 3003
l0070 = InStr(1, l005E$, l005A, 1)
If  l0070 Then GoTo 3003
List1.AddItem l005E$
Text3.Text = Text3.Text + l005E$ + ","
l0080 = Mid(frame3d1.Caption, 15, Len(frame3d1.Caption))
l0084 = Int(l0080) + 1
frame3d1.Caption = "People To IM: " + Str(l0084)

3003 :
Next l0044
End Sub

Sub Command3D3_Click ()
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As String
Dim l0092 As Variant
l0088 = extfn1018("AOL FRAME25", 0&)
l008C = extfn10F8(l0088, "Welcome")
l0090$ = String(30, 0)
l0092 = extfnFA8(l008C, l0090$, 250)
If  l0092 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "An error occurred: No Message", 16
Exit Sub
End If
If  List1.ListCount = 0 Then
MsgBox "There is nobody to instant message to!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
Timer1.Enabled = True
Timer1.Interval = 1
Command3D3.Caption = "Stop"
Else
Timer1.Interval = 0
Timer1.Enabled = False
Command3D3.Caption = "Start"
End If
End Sub

Sub Command3D4_Click ()
frame3d1.Caption = "People To IM: 0"
List1.Clear
Text3.Text = ""
End Sub

Sub Command3D5_Click ()
imform.WindowState = 1
End Sub

Sub Command3D6_Click ()
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
l00A8 = List1.ListIndex
l00AC = List1.List(l00A8)
On Error Resume Next
List1.RemoveItem l00A8
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l00B0 = InStr(1, Text3.Text, l00AC + ",", 1)
If  l00B0 Then
l00B4 = Mid(Text3.Text, 1, l00B0 - 1)
l00B8 = Mid(Text3.Text, l00B0 + Len(l00AC) + 1, Len(Text3.Text))
Text3.Text = l00B4 + l00B8
l00A8 = Mid(frame3d1.Caption, 15, Len(frame3d1))
l00B4 = l00A8 - 1
frame3d1.Caption = "People To IM: " + l00B4
End If
End If
End Sub

Sub Command3D7_Click ()
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
If  Text2.Text = "" Then Exit Sub
For l00C2 = 0 To List1.ListCount
l00C6 = InStr(1, List1.List(l00C2), Text2.Text, 1)
If  l00C6 Then
If  Len(List1.List(l00C2)) = Len(Text2.Text) Then
Text2.Text = ""
GoTo 200
End If
End If
Next l00C2
List1.AddItem Text2
Text3.Text = Text3.Text + Text2.Text + ","
Text2 = ""
l00CA = Mid(frame3d1.Caption, 15, Len(frame3d1.Caption))
l00CE = Int(l00CA) + 1
frame3d1.Caption = "People To IM: " + Str(l00CE)
200 :
End Sub

Sub Form_Load ()
If  imform.Caption <> fn410() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
extsub1050 imform.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Text2_KeyPress (KeyAscii As Integer)
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
l00D6 = Chr(KeyAscii)
If  Text2.Text = "" Then GoTo 20
If  l00D6 = Chr(13) Then
For l00DA = 0 To List1.ListCount
l00DE = InStr(1, List1.List(l00DA), Text2.Text, 1)
If  l00DE Then
If  Len(List1.List(l00DA)) = Len(Text2.Text) Then
Text2.Text = ""
GoTo 20
End If
End If
Next l00DA
List1.AddItem Text2
Text3.Text = Text3.Text + Text2.Text + ","
Text2 = ""
l00E2 = Mid(frame3d1.Caption, 15, Len(frame3d1.Caption))
l00E6 = Int(l00E2) + 1
frame3d1.Caption = "People To IM: " + Str(l00E6)
End If
20 :
End Sub

Sub Timer1_Timer ()
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F4 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0108 As Variant
Dim l010C As String
Dim l010E As Variant
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As String
Dim l011C As Variant
Dim l0120 As Variant
Dim l0126 As Variant
Dim l012A As Variant
Dim l0132 As Variant
Dim l0138 As Variant
Dim l013C As Variant
Dim l0140 As Variant
Dim l0144 As Variant
Dim l0148 As Variant
Dim l014C As Variant
Dim l0150 As Variant
Dim l0154 As Variant
l00EA = extfn1018("AOL FRAME25", 0&)
For l00EE = 0 To List1.ListCount - 1
34 :
If  Timer1.Interval = 0 Then GoTo L549F4
l00F4 = fn5D0("Mem&bers", "Send an Instant Message")
l00FA = fn138()
If  l00FA = 8 Then
36 :
DoEvents
l00FE = extfn10F8(l00EA, "Send Instant Message")
If  l00FE = 0 Then GoTo 36
Else
l00FE = extfn10F8(l00EA, "Send Instant Message")
If  l00FE = 0 Then GoTo 34
End If
l0102 = extfn10C0(l00FE, "_AOL_EDIT")
l0108 = extfn11A0(l0102, 12, 0, List1.List(l00EE))
l010C$ = String(Len(List1.List(l00EE)), 0)
l010E = extfn11A0(l0102, 13, Len(List1.List(l00EE)) + 1, l010C$)
If  l010C$ <> List1.List(l00EE) Then GoTo 36

56 :
DoEvents
If  Timer1.Interval = 0 Then GoTo L549F4
l0112 = extfn10F8(l00EA, "To:")
l0116 = extfn1088(l0112)
l011A$ = String(255, 0)
l011C = extfnFA8(l0116, l011A$, 250)
If  l011A$ Like "*Send Instant Message*" Then
Else
GoTo 56
End If
l0120 = extfn10F8(l00FE, "Send")
l0126 = extfn1168(l0120, 3)
If  l00FA = 5 Then
Else
l0126 = extfn10C0(l00FE, "RICHCNTL")
End If
l012A = Chr(13) + Chr(10)
l0132 = extfn11A0(l0126, 12, 0, Text1 + l012A + l012A + "              " + fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(86) + Chr(101) + Chr(114) + Chr(115) + Chr(105) + Chr(111) + Chr(110) + Chr(32) + Chr(57) + Chr(46) + Chr(48) + fn368())
If  l00FA = 5 Then
l0138 = fn1480(l0120)
Else
l013C = extfn1168(l0126, 2)
l0140 = fn1480(l013C)
End If
Do Until 2 > 3
If  Timer1.Interval = 0 Then GoTo L549F4
If  l00FA = 5 Then
l0144 = extfn10F8(l00EA, "Available?")
Else
l0112 = extfn10F8(l00EA, "To:")
l0116 = extfn1088(l0112)
l011A$ = String(255, 0)
l011C = extfnFA8(l0116, l011A$, 250)
If  l011A$ Like "*Send Instant Message*" Then
l0144 = extfn10F8(l0116, "To:")
Else
l0144 = 0
End If
End If
If  l0144 = 0 Then GoTo 13
l0148 = extfn1018("#32770", "America Online")
If  l0148 > 0 Then
l014C = extfn10F8(l0148, "OK")
l0150 = fn1480(l014C)
l0154 = extfnFE0(extfn1088(l0144), 16, 0, 0)
GoTo 13
End If
DoEvents
Loop
13 :
If  Timer1.Interval = 0 Then GoTo L549F4
Next l00EE
Timer1.Interval = 0
GoTo L549F4
L549F4:
Command3D3.Caption = "Start"
End Sub

Sub sub9980 ()
Unload imform
End Sub
