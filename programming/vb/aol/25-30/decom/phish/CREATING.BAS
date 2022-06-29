' CREATING.FRM
Option Explicit
Const c0026 =-1 ' &HFFFF%
Const c0028 = 3 ' &H3%

Sub Command3D1_Click ()
gv0072 = 1
Command3D1.Caption = "Wait..."
End Sub

Sub Form_Load ()
Dim l002C As Variant
l002C = extfn4120(Creating.hWnd, c0026, 0, 0, 0, 0, c0028)
Creating.Left = (Screen.Width - Creating.Width) / 2
Creating.Top = 0
End Sub
