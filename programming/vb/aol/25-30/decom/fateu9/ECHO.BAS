' ECHO.FRM
Option Explicit
Dim m001A As Variant

Sub Command3D1_Click ()
Dim l0032 As Variant
Dim l0038 As Variant
Dim l003C As String
Dim l0040 As Variant
If  echo1.Text = "" Then
MsgBox "Please enter in someone to echo!", 16
Exit Sub
End If
l0032 = extfn1018("AOL FRAME25", 0&)
l0038 = extfn10F8(l0032, "Welcome")
l003C$ = String(30, 0)
l0040 = extfnFA8(l0038, l003C$, 250)
If  l0040 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  m001A = 0 Then
Command3D1.Caption = "Stop"
text1.Text = fn170()
m001A = 1
Timer1.Enabled = True
Timer1.Interval = 1
Else
m001A = 0
Command3D1.Caption = "Echo"
Timer1.Enabled = False
Timer1.Interval = 0
End If
End Sub

Sub Command3D2_Click ()
ECHO.WindowState = 1
End Sub

Sub Form_Load ()
extsub1050 ECHO.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l005E As Variant
text1.Text = fn170()
l005E = fn1368()
End Sub

Sub sub9980 ()
Unload ECHO
End Sub
