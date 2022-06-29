' MINI.FRM
Option Explicit

Sub Command3D1_Click ()
main.Show
Unload mini
End Sub

Sub Command3D2_Click ()
Dim l0024 As Variant
l0024 = extfn1210("cmbultra.wav", 1)
End
End Sub

Sub Form_Load ()
mini.Top = 0
mini.Left = 8590 / 2
mini.Width = 1360
extsub1050 mini.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
extsub1050 mini.hWnd, -1, 0, 0, 0, 0, 3
End Sub
