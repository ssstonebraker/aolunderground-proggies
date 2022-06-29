' MAILFIX.FRM
Option Explicit

Sub Command3D2_Click ()
Dim l0022 As Variant
Dim l0028 As Variant
Dim l002E As Variant
Dim l0032 As Variant
Dim l0036 As Variant
Dim l0042 As Variant
Dim l0048 As Variant
Dim l004E As Variant
Dim l0064 As Variant
Dim l0068 As String
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
If  Option3D1.Value = True Then
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10C0(l0022, "_AOL_TREE")
l002E = extfnFE0(l0028, 1036, 0, 0)
l0032 = extfnFE0(l0028, 1031, 0, 0)
For l0036 = 1 To Int(Text4.Text)
l0042 = extfn1088(l0028)
l0048 = extfn10F8(l0042, "Delete")
DoEvents
DoEvents
l004E = fn1480(l0048)
DoEvents
DoEvents
Next l0036
End If
If  Option3D2.Value = True Then
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10C0(l0022, "_AOL_TREE")
l002E = extfnFE0(l0028, 1036, 0, 0)
l0032 = extfnFE0(l0028, 1031, 0, 0)
For l0036 = 0 To l002E - 1
l0042 = extfn1088(l0028)
l0048 = extfn10F8(l0042, "Delete")
l004E = fn1480(l0048)
Next l0036
End If
If  Option3D3.Value = True Then
If  Text1.Text = "" Then
MsgBox "Please enter in the person to delete mails from!", 16
Exit Sub
End If
l0064 = 0
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10F8(l0022, "New Mail")
l002E = extfn10C0(l0028, "_AOL_TREE")
l0036 = extfnFE0(l002E, 1036, 0, 0)
For l0032 = 0 To l0036 - 1
10 :
l0068$ = String(255, 0)
l0048 = extfn11A0(l002E, 1034, l0032, l0068$)
l004E = InStr(l0068$, Chr(9))
l006C = Mid(l0068$, l004E + 1, Len(l0068$))
l0070 = InStr(l006C, Chr(9))
On Error Resume Next
l0074 = Mid(l006C, 1, l0070 - 1)
If  Err Then GoTo 20
l0078 = InStr(1, l0074, Text1.Text, 1)
If  l0078 Then
If  Len(l0074) = Len(Text1.Text) Then
l007C = extfnFE0(l002E, 1031, l0032, 0)
l0080 = extfn10F8(l0028, "Delete")
l0084 = fn1480(l0080)
l0064 = l0064 + 1
GoTo 10
End If
End If
DoEvents
Next l0032
20 :
If  l0064 = 0 Then
MsgBox "0 mails deleted!", 16
Else
MsgBox l0064 + " mails deleted!", 16
End If
End If
If  Option3D4.Value = True Then
If  Text2.Text = "" Then
MsgBox "Please enter in the person to delete mails from!", 16
Exit Sub
End If
l0064 = 0
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10F8(l0022, "New Mail")
l002E = extfn10C0(l0028, "_AOL_TREE")
l0036 = extfnFE0(l002E, 1036, 0, 0)
For l0032 = 0 To l0036 - 1
l0068$ = String(255, 0)
l0048 = extfn11A0(l002E, 1034, l0032, l0068$)
l004E = InStr(l0068$, Chr(9))
l006C = Mid(l0068$, l004E + 1, Len(l0068$))
l0070 = InStr(l006C, Chr(9))
On Error Resume Next
l0074 = Mid(l006C, 1, l0070 - 1)
If  Err Then GoTo 200
l0078 = InStr(1, l0074, Text2.Text, 1)
If  l0078 Then
If  Len(l0074) = Len(Text2.Text) Then
l007C = extfnFE0(l002E, 1031, l0032, 0)
l0080 = extfn10F8(l0028, "Ignore")
l0084 = fn1480(l0080)
l0064 = l0064 + 1
End If
End If
DoEvents
Next l0032
200 :
If  l0064 = 0 Then
MsgBox "0 mails Ignored!", 16
Else
MsgBox l0064 + " mails ignored!", 16
End If
End If
If  Option3D5.Value = True Then
If  Text3.Text = "" Then
MsgBox "Please enter in the string to delete mails from!", 16
Exit Sub
End If
l0064 = 0
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10F8(l0022, "New Mail")
l002E = extfn10C0(l0028, "_AOL_TREE")
l0036 = extfnFE0(l002E, 1036, 0, 0)
For l0032 = 0 To l0036 - 1
100 :
l0068$ = String(255, 0)
l0048 = extfn11A0(l002E, 1034, l0032, l0068$)
l004E = InStr(l0068$, Chr(9))
l006C = Mid(l0068$, l004E + 1, Len(l0068$))
l0070 = InStr(l006C, Chr(9))
On Error Resume Next
l0074 = Mid(l006C, l0070 + 1, Len(l006C))
If  Err Then GoTo 2000
l0078 = InStr(1, l0074, Text3.Text, 1)
If  l0078 Then
l007C = extfnFE0(l002E, 1031, l0032, 0)
l0080 = extfn10F8(l0028, "Delete")
l0084 = fn1480(l0080)
l0064 = l0064 + 1
GoTo 100
End If
DoEvents
Next l0032
2000 :
If  l0064 = 0 Then
MsgBox "0 mails deleted!", 16
Else
MsgBox l0064 + " mails deleted!", 16
End If
End If
If  Option3D6.Value = True Then
If  Text5.Text = "" Then
MsgBox "Please enter in the string to delete mails from!", 16
Exit Sub
End If
l0064 = 0
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10F8(l0022, "New Mail")
l002E = extfn10C0(l0028, "_AOL_TREE")
l0036 = extfnFE0(l002E, 1036, 0, 0)
For l0032 = 0 To l0036 - 1
l0068$ = String(255, 0)
l0048 = extfn11A0(l002E, 1034, l0032, l0068$)
l004E = InStr(l0068$, Chr(9))
l006C = Mid(l0068$, l004E + 1, Len(l0068$))
l0070 = InStr(l006C, Chr(9))
On Error Resume Next
l0074 = Mid(l006C, l0070 + 1, Len(l006C))
If  Err Then GoTo 20000
l0078 = InStr(1, l0074, Text5.Text, 1)
If  l0078 Then
l007C = extfnFE0(l002E, 1031, l0032, 0)
l0080 = extfn10F8(l0028, "Ignore")
l0084 = fn1480(l0080)
l0064 = l0064 + 1
End If
DoEvents
Next l0032
20000 :
If  l0064 = 0 Then
MsgBox "0 mails ignored!", 16
Else
MsgBox l0064 + " mails ignored!", 16
End If
End If
End Sub

Sub Form_Load ()
extsub1050 mailfix.hWnd, -1, 0, 0, 0, 0, 3
Option3D2.Value = True
End Sub

Sub HScroll1_Change ()
End Sub
