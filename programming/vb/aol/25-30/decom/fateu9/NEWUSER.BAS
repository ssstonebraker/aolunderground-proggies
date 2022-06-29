' NEWUSER.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0030 As Variant
Dim l0034 As Variant
Dim l0038 As Variant
Dim l003C As String
Dim l003E As Variant
Dim l0042 As String
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004E As Variant
Dim l0058 As Variant
If  Text1.Text = "" Then
MsgBox "Please enter in your America Online directory!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "Please enter in a screen name to replace with!", 16
Exit Sub
End If
Command3D1.Enabled = False
pwait.Show
l0030 = Len(Text2)
Select Case l0030
Case 3
l0034 = Text2 + "       "
Case 4
l0034 = Text2 + "      "
Case 5
l0034 = Text2 + "     "
Case 6
l0034 = Text2 + "    "
Case 7
l0034 = Text2 + "   "
Case 8
l0034 = Text2 + "  "
Case 9
l0034 = Text2 + " "
Case 10
l0034 = Text2
End Select
l0038 = 1
Do Until 2 > 3
DoEvents
l003C$ = ""
Open Text1 + "\idb\main.idx" For Binary As #1
l003C$ = String(32000, 0)
Get #1, l0038, l003C$
Close #1
Open Text1 + "\idb\main.idx" For Binary As #2
l003E = InStr(1, l003C$, l0034, 1)
If  l003E Then
Mid(l003C$, l003E) = "New User  "
l0042$ = "New User  "
Put #2, l0038 + l003E - 1, l0042$
40 :
DoEvents
l0044 = InStr(1, l003C$, l0034, 1)
If  l0044 Then
Mid(l003C$, l0044) = "New User  "
Put #2, l0038 + l0044 - 1, l0042$
GoTo 40
End If
End If
l0038 = l0038 + 32000
l0048 = LOF(2)
Close #2
If  l0038 > l0048 Then GoTo 30
Loop
30 :
l004E = extfn1018("AOL FRAME25", 0&)
l0038 = extfn10F8(l004E, "Welcome")
If  l0038 > 0 Then
l0048 = extfnFE0(l0038, 16, 0, 0)
l0058 = fn5D0("&Go To", "Set Up && Sign On")
End If
l0048 = extfn10F8(l004E, "Goodbye")
If  l0048 > 0 Then
l0048 = extfnFE0(l0048, 16, 0, 0)
l0058 = fn5D0("&Go To", "Set Up && Sign On")
End If
Command3D1.Enabled = True
Unload pwait
End Sub

Sub Command3D2_Click ()
Unload newuser
End Sub

Sub Form_Load ()
Dim l0060 As Variant
Dim l0064 As String
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
extsub1050 newuser.hWnd, -1, 0, 0, 0, 0, 3
Open "fate.ini" For Binary As #1
l0060 = LOF(1)
l0064$ = String(l0060, 0)
Get #1, 1, l0064$
Close #1
l0066 = InStr(1, l0064$, "WAOLDir=", 1)
l006A = Mid(l0064$, l0066, Len(l0064$))
l006E = InStr(l006A, "=")
l0072 = Mid(l006A, l006E + 1, Len(l006A))
l0076 = InStr(l0072, Chr(13))
l007A = Mid(l0072, 1, l0076 - 1)
Text1 = l007A
End Sub

Sub sub9980 ()
Unload newuser
End Sub
