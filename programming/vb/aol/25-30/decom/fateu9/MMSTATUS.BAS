' MMSTATUS.FRM
Option Explicit

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
mmER.Timer2.Interval = 0
mmER.Show
Unload mmstatus
End Sub

Sub Command3D2_Click ()
mmstatus.WindowState = 1
End Sub

Sub Form_Load ()
extsub1050 mmstatus.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub sub9980 ()
Dim l0026 As Variant
mmER.Timer2.Interval = 0
mmER.Show
l0026 = fn480()
Unload mmstatus
End Sub
