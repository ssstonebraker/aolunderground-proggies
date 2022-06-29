' VALID.FRM
Option Explicit

Sub Command3D2_Click ()
Dim l0020 As Variant
Dim l002A As Variant
If  Text3.Text = "" Then GoTo 34
For l0020 = 0 To List1.ListCount - 1
l002A = InStr(1, List1.List(l0020), Text3.Text, 1)
If  l002A Then
If  Len(List1.List(l0020)) = Len(Text3.Text) Then
GoTo 34
End If
End If
Next l0020
List1.AddItem Text3.Text
34 :
Text3 = ""
End Sub

Sub Command3D3_Click ()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Sub Command3D4_Click ()
Dim l0030 As Variant
Dim l0036 As Variant
Dim l003A As String
Dim l003E As Variant
Dim l0048 As Variant
Dim l004C As Variant
Dim l0052 As Variant
Dim l0058 As Variant
Dim l0062 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0094 As Variant
Dim l009A As Variant
Dim l009E As String
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D2 As Variant
l0030 = extfn1018("AOL FRAME25", 0&)
l0036 = extfn10F8(l0030, "Welcome")
l003A$ = String(30, 0)
l003E = extfnFA8(l0036, l003A$, 250)
If  l003E <= 7 Then
MsgBox "An error occurred: Not signed on!", 16
Exit Sub
End If
If  List1.ListCount = 0 Then
MsgBox "You must need at least one person to check for validity!", 16
Exit Sub
End If
List2.Clear
For l0048 = 0 To List1.ListCount - 1
l004C = List1.List(l0048)
List2.AddItem l004C
Next l0048
l0030 = extfn1018("AOL FRAME25", 0&)
l0036 = fn5D0("&Mail", "&Compose Mail")
100 :
DoEvents
l0052 = extfn10F8(l0030, "Compose Mail")
If  l0052 = 0 Then GoTo 100
l003E = extfn10F8(l0052, "To:")
l0058 = extfn1168(l003E, 2)
text2.Text = ""
text2.Text = "Da Mailz" + ","
For l0062 = 0 To List2.ListCount - 1
text2.Text = text2.Text + List2.List(l0062) + ","
Next l0062
l0068 = extfn11A0(l0058, 12, 0, text2.Text)
l006C = extfn10F8(l0052, "Subject:")
l0070 = extfn1168(l006C, 2)
l0074 = extfn11A0(l0070, 12, 0, "Fate Ultra 9.0 - Account Validity")
l0078 = extfn1168(l003E, 1)
l007C = extfn11A0(l0078, 12, 0, "Fate Ultra 9.0 - Account Valididy")
l0080 = extfn10F8(l0052, "Send")
l0084 = extfn1168(l0080, 2)
l008A = fn1480(l0084)
200 :
DoEvents
l008E = extfn10F8(l0030, "Error")
If  l008E = 0 Then GoTo 200
l0094 = extfn10C0(l008E, "_AOL_VIEW")
l009A = extfnFE0(l0094, 14, 0, 0)
l009E$ = String(255, 0)
l00A0 = extfn11A0(l0094, 13, l009A, l009E$)
text2.Text = l009E$
300 :
l00A4 = InStr(text2.Text, Chr(10))
If  l00A4 Then
l00A8 = Mid(text2.Text, 1, l00A4 - 2)
text4.Text = l00A8
text2.Text = Mid(text2.Text, l00A4 + 1, Len(text2.Text))
l00B2 = InStr(text4.Text, "-")
If  l00B2 Then
l00B6 = Mid(text4.Text, 1, l00B2 - 2)
For l00BA = 0 To List2.ListCount - 1
l00BE = InStr(1, l00B6, List2.List(l00BA), 1)
If  l00BE Then
If  Len(l00B6) = Len(List2.List(l00BA)) Then
l00C2 = InStr(1, text4.Text, "known", 1)
If  l00C2 Then
List2.List(l00BA) = List2.List(l00BA) + " - Terminated"
Else
List2.List(l00BA) = List2.List(l00BA) + " - Alive"
End If
End If
End If
Next l00BA
End If
GoTo 300
End If
l00B2 = InStr(text2.Text, "-")
If  l00B2 Then
l00B6 = Mid(text2.Text, 1, l00B2 - 2)
For l00BA = 0 To List2.ListCount - 1
l00BE = InStr(1, l00B6, List2.List(l00BA), 1)
If  l00BE Then
If  Len(l00B6) = Len(List2.List(l00BA)) Then
l00C2 = InStr(1, text2.Text, "known", 1)
If  l00C2 Then
List2.List(l00BA) = List2.List(l00BA) + " - Terminated"
Else
List2.List(l00BA) = List2.List(l00BA) + " - Alive"
End If
End If
End If
Next l00BA
End If
For l00C6 = 0 To List2.ListCount - 1
l0048 = List2.List(l00C6)
l004C = InStr(l0048, "-")
If  l004C = 0 Then
List2.List(l00C6) = List2.List(l00C6) + " - Alive"
End If
Next l00C6
l00CA = extfn10F8(l008E, "OK")
l00CE = fn1480(l00CA)
l00D2 = extfnFE0(l0052, 16, 0, 0)
text2.Text = ""
End Sub

Sub Command3D5_Click ()
List1.Clear
End Sub

Sub Command3D6_Click ()
List2.Clear
End Sub

Sub Command3D7_Click ()
valid.WindowState = 1
End Sub

Sub Form_Load ()
extsub1050 valid.hWnd, -1, 0, 0, 0, 0, 3
If  valid.Caption <> fnC28() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub Text3_KeyPress (KeyAscii As Integer)
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
l00DE = Chr(KeyAscii)
If  l00DE = Chr(13) Then
For l00E2 = 0 To List1.ListCount - 1
l00E6 = InStr(1, List1.List(l00E2), Text3.Text, 1)
If  l00E6 Then
If  Len(List1.List(l00E2)) = Len(Text3.Text) Then
Text3 = ""
GoTo 23
End If
End If
Next l00E2
List1.AddItem Text3.Text
Text3 = ""
End If
23 :
End Sub
