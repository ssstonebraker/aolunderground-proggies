' ADDROOM.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l002E As Variant
Dim l0032 As Variant
Dim l0036 As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As String
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As String
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006A As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l002E = MsgBox("Do you wish to add @aol.com after the names?", 32 + 4, "Add Room")
If  l002E = 6 Then
l0032 = "yes"
Else
l0032 = "no"
End If
l002E = extfn1018("AOL FRAME25", 0&)
l0036 = extfn10F8(l002E, "AOL Live!")
l0036 = extfn1088(l0036)
l0036 = extfn10C0(l0036, "_AOL_LISTBOX")
l0040 = extfnFE0(l0036, 1036, 0, 0)
For l0044 = 0 To l0040 - 1
l0048$ = String(30, 0)
l004C = extfn1608(l0044, l0048$)
l0050 = extfn10F8(l002E, "Welcome")
l0054$ = String(30, 0)
l0056 = extfnFA8(l0050, l0054$, 250)
l005A = Mid(l0054$, 10, 10)
l005E = InStr(l005A, "!")
If  l005E Then
l005A = Mid(l005A, 1, l005E - 1)
End If
l0062 = InStr(l0048$, l005A)
If  l0062 Then GoTo 80
l0066 = InStr(l0048$, Chr(0))
If  l0066 Then
l006A = Mid(Trim(l0048$), 1, l0066 - 1)
Else
l006A = l0048$
End If
If  l0032 = "yes" Then
Text1.Text = Text1.Text + "," + l006A + "@aol.com"
Else
Text1.Text = Text1.Text + "," + l006A
End If
80 :
Next l0044
End Sub

Sub Command3D2_Click ()
Clipboard.Clear
Clipboard.SetText Text1.Text
End Sub

Sub Command3D3_Click ()
Clipboard.SetText Text1.Text
Text1.Text = ""
End Sub

Sub Command3D4_Click ()
Dim l0078 As Variant
l0078 = MsgBox("Are you sure you want to delete this?", 4 + 32)
If  l0078 = 6 Then
Text1.Text = ""
Else
End If
End Sub

Sub Form_Load ()
extsub1050 addroom.hWnd, -1, 0, 0, 0, 0, 3
End Sub
