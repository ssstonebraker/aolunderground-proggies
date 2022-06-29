' MACROKIL.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l0030 As Variant
Dim l0034 As Variant
Dim l0038 As Variant
Dim l003C As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0030 = fnA30("———————————————    ———————————————    ———————————————    ———————————————    ")
DoEvents
l0034 = fnA30("———————————————    ———————————————    ———————————————    ———————————————    ")
DoEvents
l0038 = fnA30("———————————————    ———————————————    ———————————————    ———————————————    ")
DoEvents
l003C = fnA30("———————————————    ———————————————    ———————————————    ———————————————    ")
End Sub

Sub Form_Load ()
extsub1050 macrokill.hWnd, -1, 0, 0, 0, 0, 3
macrokill.Width = 1095
End Sub
