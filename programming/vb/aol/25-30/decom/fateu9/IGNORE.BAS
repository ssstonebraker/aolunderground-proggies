' IGNORE.FRM
Option Explicit

Sub sub9748 ()
End Sub

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  List1.ListCount = 0 Then
MsgBox "You Must Enter In Someone To Ignore"
Exit Sub
End If
If  Timer1.Interval = 0 Then
Timer1.Enabled = True
Timer1.Interval = 1
Command3D1.Caption = "Stop"
Else
Command3D1.Caption = "Okay"
Timer1.Interval = 0
Timer1.Enabled = False
End If
End Sub

Sub Command3D2_Click ()
List1.Clear
End Sub

Sub Command3D3_Click ()
ignore.WindowState = 1
End Sub

Sub Command3D4_Click ()
Dim l004A As Variant
Dim l004E As Variant
Dim l0052 As Variant
If  Text1.Text = "" Then
Text1.Text = ""
GoTo 200
End If
For l004A = 0 To List1.ListCount - 1
l004E = List1.List(l004A)
l0052 = InStr(1, l004E, Text1.Text, 1)
If  l0052 Then
If  Len(List1.List(l004A)) = Len(Text1.Text) Then
Text1.Text = ""
GoTo 200
End If
End If
DoEvents
Next l004A
List1.AddItem Text1
Text1 = ""
200 :
End Sub

Sub Command3D5_Click ()
Dim l0056 As Variant
l0056 = List1.ListIndex
On Error Resume Next
List1.RemoveItem l0056
If  Err Then
MsgBox "Highlight the person you want to take off!", 16
Exit Sub
End If
End Sub

Sub Form_Load ()
extsub1050 ignore.hWnd, -1, 0, 0, 0, 0, 3
List1.Height = 1050
End Sub

Sub text1_KeyPress (KeyAscii As Integer)
Dim l005C As Variant
Dim l0060 As Variant
Dim l0064 As Variant
Dim l0068 As Variant
l005C = Chr(KeyAscii)
If  l005C = Chr(13) Then
If  Text1.Text = "" Then
Text1.Text = ""
GoTo 20
End If
For l0060 = 0 To List1.ListCount - 1
l0064 = List1.List(l0060)
l0068 = InStr(1, l0064, Text1.Text, 1)
If  l0068 Then
If  Len(List1.List(l0060)) = Len(Text1.Text) Then
Text1.Text = ""
GoTo 20
End If
End If
DoEvents
Next l0060
List1.AddItem Text1
Text1 = ""
End If
20 :
End Sub

Sub Timer1_Timer ()
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As String
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l0090 As Variant
10 :
l006C = extfn1018("AOL FRAME25", 0&)
l0070 = extfn10F8(l006C, ">Instant Message From: ")
If  Timer1.Interval = 0 Then GoTo L637DE
DoEvents
If  l0070 = 0 Then GoTo 10
l0074$ = String(255, 0)
l0076 = extfnFA8(l0070, l0074$, 250)
l007A = Mid(l0074$, 24, 10)
l007E = InStr(l007A, Chr(0))
If  l007E Then
l007A = Mid(l007A, 1, l007E - 1)
End If
For l0082 = 0 To List1.ListCount
l0086 = List1.List(l0082)
l008A = InStr(1, l0086, l007A, 1)
If  l008A Then
If  Len(l0086) = Len(l007A) Then
If  Timer1.Interval = 0 Then GoTo L637DE
l0090 = extfnFE0(l0070, 16, 0, 0)
GoTo 10
End If
End If
If  Timer1.Interval = 0 Then GoTo L637DE
DoEvents
Next l0082
If  Timer1.Interval = 0 Then GoTo L637DE
DoEvents
GoTo 10
L637DE:
End Sub

Sub sub9980 (p0094 As Integer)
Select Case p0094
Case 0
Unload ignore
Case 1
ignore.WindowState = 1
End Select
End Sub

Sub sub9620 ()
End Sub
