' MMSL.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l002A As Variant
Dim l002E As Variant
Dim l0038 As Variant
Dim l003C As Variant
Dim l0040 As Variant
Dim l0054 As String
Dim l0056 As Variant
Dim l005A As String
Dim l005C As String
Dim l005E As Variant
Dim l0062 As String
Dim l0064 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
l0020 = InStr(text1.Text, "*")
If  l0020 Then
File1.Pattern = text1.Text
Exit Sub
End If
l002A = InStr(text1.Text, ":")
If  l002A Then
l002E = text1.Text
GoTo 10
End If
l0038 = Dir1.Path
l003C = Mid(l0038, Len(l0038), 1)
If  l003C = "\" Then
l0040 = ""
Else
l0040 = "\"
End If
l002E = l0038 + l0040 + text1.Text
10 :
If  Command3D1.Caption = "Save" Then
If  Option3D1.Value = True Then
On Error Resume Next
Open l002E For Output As #1
If  Err Then
MsgBox l002E + " - Invalid Path!", 16
Exit Sub
Close #1
End If
l0054$ = mmER.Text2.Text
Print #1, l0054$
Close #1
Else
On Error Resume Next
Open l002E For Append As #2
If  Err Then
MsgBox l002E + " - Invalid Path!", 16
Exit Sub
Close #2
End If
For l0056 = 0 To mmER.List1.ListCount - 1
l005A$ = mmER.List1.List(l0056)
Print #2, l005A$
Next l0056
Close #2
End If
Else
If  Option3D1.Value = True Then
On Error Resume Next
Open l002E For Binary As #3
If  Err Then
MsgBox l002E + " - Invalid Path!", 16
Exit Sub
Close #3
End If
On Error Resume Next
l005C$ = String(LOF(3), 0)
If  Err Then
MsgBox "This file is to big to open!", 16
Close #3
Exit Sub
End If
Get #3, 1, l005C$
mmER.Text2.Text = l005C$
Close #3
Else
On Error Resume Next
Open l002E For Binary As #4
If  Err Then
MsgBox l002E + " - Invalid Path!", 16
Exit Sub
Close #4
End If
On Error Resume Next
l005E = LOF(4)
l0062$ = String(LOF(4), 0)
If  Err Then
MsgBox "This file is to big to open!", 16
Close #4
Exit Sub
End If
Get #4, 1, l0062$
50 :
l0064 = InStr(l0062$, Chr(13))
If  l0064 Then
l0068 = Mid(l0062$, 1, l0064 - 1)
If  l0068 = "" Then GoTo 80
mmER.List1.AddItem l0068
mmER.Text1.Text = mmER.Text1.Text + l0068 + ","
l006C = Mid(mmER.Frame3D1.Caption, 15, Len(mmER.Frame3D1.Caption))
l0070 = l006C + 1
mmER.Frame3D1.Caption = "People To MM: " + l0070
80 :
l0062$ = Mid(l0062$, l0064 + 2, Len(l0062$))
GoTo 50
Else
If  l0062$ = "" Then GoTo 70
mmER.List1.AddItem l0062$
mmER.Text1.Text = mmER.Text1.Text + l0062$ + ","
l006C = Mid(mmER.Frame3D1.Caption, 15, Len(mmER.Frame3D1.Caption))
l0070 = l006C + 1
mmER.Frame3D1.Caption = "People To MM: " + l0070
70 :
Close #4
End If
End If
End If
mmsl.Hide
End Sub

Sub Command3D2_Click ()
Unload mmsl
End Sub

Sub Dir1_Change ()
File1.Path = Dir1.Path
End Sub

Sub Drive1_Change ()
On Error Resume Next
Dir1.Path = Drive1.Drive
If  Err Then
MsgBox "Drive is not available", 16
Exit Sub
End If
End Sub

Sub File1_Click ()
Dim l007E As Variant
l007E = File1.List(File1.ListIndex)
text1.Text = l007E
End Sub

Sub File1_DblClick ()
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As String
Dim l009C As Variant
Dim l00A0 As String
Dim l00A2 As String
Dim l00A4 As Variant
Dim l00A8 As String
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
l0082 = InStr(text1.Text, "*")
If  l0082 Then
File1.Pattern = text1.Text
Exit Sub
End If
l0086 = InStr(text1.Text, ":")
If  l0086 Then
l008A = text1.Text
GoTo 1
End If
l008E = Dir1.Path
l0092 = Mid(l008E, Len(l008E), 1)
If  l0092 = "\" Then
l0096 = ""
Else
l0096 = "\"
End If
l008A = l008E + l0096 + text1.Text
1 :
If  Command3D1.Caption = "Save" Then
If  Option3D1.Value = True Then
On Error Resume Next
Open l008A For Output As #1
If  Err Then
MsgBox l008A + " - Invalid Path!", 16
Exit Sub
Close #1
End If
l009A$ = mmER.Text2.Text
Print #1, l009A$
Close #1
Else
On Error Resume Next
Open l008A For Append As #2
If  Err Then
MsgBox l008A + " - Invalid Path!", 16
Exit Sub
Close #2
End If
For l009C = 0 To mmER.List1.ListCount - 1
l00A0$ = mmER.List1.List(l009C)
Print #2, l00A0$
Next l009C
Close #2
End If
Else
If  Option3D1.Value = True Then
On Error Resume Next
Open l008A For Binary As #3
If  Err Then
MsgBox l008A + " - Invalid Path!", 16
Exit Sub
Close #3
End If
On Error Resume Next
l00A2$ = String(LOF(3), 0)
If  Err Then
MsgBox "This file is to big to open!", 16
Close #3
Exit Sub
End If
Get #3, 1, l00A2$
mmER.Text2.Text = l00A2$
Close #3
Else
On Error Resume Next
Open l008A For Binary As #4
If  Err Then
MsgBox l008A + " - Invalid Path!", 16
Exit Sub
Close #4
End If
On Error Resume Next
l00A4 = LOF(4)
l00A8$ = String(LOF(4), 0)
If  Err Then
MsgBox "This file is to big to open!", 16
Close #4
Exit Sub
End If
Get #4, 1, l00A8$
2 :
l00AA = InStr(l00A8$, Chr(13))
If  l00AA Then
l00AE = Mid(l00A8$, 1, l00AA - 1)
If  l00AE = "" Then GoTo 3
mmER.List1.AddItem l00AE
l00B2 = Mid(mmER.Frame3D1.Caption, 15, Len(mmER.Frame3D1.Caption))
l00B6 = l00B2 + 1
mmER.Frame3D1.Caption = "People To MM: " + l00B6
3 :
l00A8$ = Mid(l00A8$, l00AA + 2, Len(l00A8$))
GoTo 2
Else
If  l00A8$ = "" Then GoTo 4
mmER.List1.AddItem l00A8$
l00B2 = Mid(mmER.Frame3D1.Caption, 15, Len(mmER.Frame3D1.Caption))
l00B6 = l00B2 + 1
mmER.Frame3D1.Caption = "People To MM: " + l00B6
4 :
Close #4
End If
End If
End If
mmsl.Hide
End Sub

Sub Form_Load ()
extsub1050 mmsl.hWnd, -1, 0, 0, 0, 0, 3
Option3D1.Value = True
text1.AddItem "*.*"
text1.AddItem "*.bat"
End Sub

Sub text1_KeyPress (KeyAscii As Integer)
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As String
Dim l00D6 As Variant
Dim l00DA As String
Dim l00DC As String
Dim l00DE As Variant
Dim l00E2 As String
Dim l00E4 As Variant
Dim l00E8 As Variant
Dim l00EC As Variant
Dim l00F0 As Variant
l00BC = Chr(KeyAscii)
If  l00BC = Chr(13) Then
l00BC = InStr(text1.Text, "*")
If  l00BC Then
File1.Pattern = text1.Text
Exit Sub
End If
l00C0 = InStr(text1.Text, ":")
If  l00C0 Then
l00C4 = text1.Text
GoTo 5
End If
l00C8 = Dir1.Path
l00CC = Mid(l00C8, Len(l00C8), 1)
If  l00CC = "\" Then
l00D0 = ""
Else
l00D0 = "\"
End If
l00C4 = l00C8 + l00D0 + text1.Text
5 :
If  Command3D1.Caption = "Save" Then
If  Option3D1.Value = True Then
On Error Resume Next
Open l00C4 For Output As #1
If  Err Then
MsgBox l00C4 + " - Invalid Path!", 16
Exit Sub
Close #1
End If
l00D4$ = mmER.Text2.Text
Print #1, l00D4$
Close #1
Else
On Error Resume Next
Open l00C4 For Append As #2
If  Err Then
MsgBox l00C4 + " - Invalid Path!", 16
Exit Sub
Close #2
End If
For l00D6 = 0 To mmER.List1.ListCount - 1
l00DA$ = mmER.List1.List(l00D6)
Print #2, l00DA$
Next l00D6
Close #2
End If
Else
If  Option3D1.Value = True Then
On Error Resume Next
Open l00C4 For Binary As #3
If  Err Then
MsgBox l00C4 + " - Invalid Path!", 16
Exit Sub
Close #3
End If
On Error Resume Next
l00DC$ = String(LOF(3), 0)
If  Err Then
MsgBox "This file is to big to open!", 16
Close #3
Exit Sub
End If
Get #3, 1, l00DC$
mmER.Text2.Text = l00DC$
Close #3
Else
On Error Resume Next
Open l00C4 For Binary As #4
If  Err Then
MsgBox l00C4 + " - Invalid Path!", 16
Exit Sub
Close #4
End If
On Error Resume Next
l00DE = LOF(4)
l00E2$ = String(LOF(4), 0)
If  Err Then
MsgBox "This file is to big to open!", 16
Close #4
Exit Sub
End If
Get #4, 1, l00E2$
6 :
l00E4 = InStr(l00E2$, Chr(13))
If  l00E4 Then
l00E8 = Mid(l00E2$, 1, l00E4 - 1)
If  l00E8 = "" Then GoTo 7
mmER.List1.AddItem l00E8
l00EC = Mid(mmER.Frame3D1.Caption, 15, Len(mmER.Frame3D1.Caption))
l00F0 = l00EC + 1
mmER.Frame3D1.Caption = "People To MM: " + l00F0
7 :
l00E2$ = Mid(l00E2$, l00E4 + 2, Len(l00E2$))
GoTo 6
Else
If  l00E2$ = "" Then GoTo 8
mmER.List1.AddItem l00E2$
l00EC = Mid(mmER.Frame3D1.Caption, 15, Len(mmER.Frame3D1.Caption))
l00F0 = l00EC + 1
mmER.Frame3D1.Caption = "People To MM: " + l00F0
8 :
Close #4
End If
End If
End If
mmsl.Hide
End If
End Sub
