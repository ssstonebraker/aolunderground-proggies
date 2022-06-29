' WAITFORR.FRM
Option Explicit
Const c001E =-1 ' &HFFFF%
Const c0020 = 3 ' &H3%

Sub Form_Load ()
Dim l0024 As Variant
l0024 = extfn4120(WaitForReset.hWnd, c001E, 0, 0, 0, 0, c0020)
WaitForReset.Left = (Screen.Width - WaitForReset.Width) / 2
WaitForReset.Top = (Screen.Height - WaitForReset.Height) / 2
End Sub
