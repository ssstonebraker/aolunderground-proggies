' COUNTMM.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l0038 As Variant
Dim l003E As Variant
Dim l0046 As Variant
If  Option3D1.Value = True Then l0020 = "New Mail"
If  Option3D2.Value = True Then l0020 = "Old Mail"
If  Option3D3.Value = True Then l0020 = "Outgoing Mail"
If  Option3D4.Value = True Then l0020 = "Incoming FlashMail"
l0038 = extfn1018("AOL FRAME25", 0&)
l003E = extfn10F8(l0038, l0020)
l003E = extfn10C0(l003E, "_AOL_TREE")
l0046 = extfnFE0(l003E, 1036, 0, 0)
If  l0046 = 1 Then GoTo 230
If  l0046 = 0 Then GoTo 230
MsgBox "You have " + l0046 + " Mails"
GoTo 250
230 :
MsgBox "You have " + l0046 + " Mail"
250 :
End Sub

Sub Form_Load ()
extsub1050 countmm.hWnd, -1, 0, 0, 0, 0, 3
Option3D1.Value = True
End Sub
