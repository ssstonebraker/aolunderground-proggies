' GETMM.FRM
Option Explicit

Sub sub99F8 ()
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
If  Command3D1.Caption = "Start Bot" Then
Command3D1.Caption = "Stop Bot"
Timer1.Enabled = True
Timer1.Interval = 1
Else
Timer1.Interval = 0
Timer1.Enabled = False
Command3D1.Caption = "Start Bot"
End If
End Sub

Sub Command3D2_Click ()
List1.Clear
Frame3D1.Caption = "Added to: 0"
End Sub

Sub Command3D3_Click ()
getmm.WindowState = 1
End Sub

Sub Form_Load ()
extsub1050 getmm.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l0054 As Variant
If  Timer1.Interval = 0 Then GoTo 10
Text1.Text = fn170()
l0054 = fn1410()
10 :
End Sub
