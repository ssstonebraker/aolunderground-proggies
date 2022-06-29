' KILLSTAT.FRM
Option Explicit

Sub Command3D1_Click ()
kills.Timer1.Interval = 0
Unload killstatus
kills.Show
End Sub

Sub Form_Load ()
extsub1050 killstatus.hWnd, -1, 0, 0, 0, 0, 3
End Sub
