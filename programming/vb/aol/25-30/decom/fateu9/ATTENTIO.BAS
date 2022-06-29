' ATTENTIO.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0022 As Variant
Dim l0028 As Variant
Dim l002C As String
Dim l0030 As Variant
Dim l0034 As Variant
Dim l0038 As Variant
Dim l003C As Variant
Dim l0042 As Variant
Dim l0048 As Variant
Dim l004E As Variant
Dim l0054 As Variant
Dim l005A As Variant
Dim l0060 As Variant
Dim l0064 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
If  Text1.Text = "" Then
MsgBox "Please enter something to give attention to!", 16
Exit Sub
End If
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10F8(l0022, "Welcome")
l002C$ = String(30, 0)
l0030 = extfnFA8(l0028, l002C$, 250)
If  l0030 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0034 = Text1.Text
l0038 = extfn1018("AOL FRAME25", 0&)
l003C = extfn10F8(l0038, "AOL Live!")
l0042 = extfn1088(l003C)
l0048 = extfn10C0(l0042, "_AOL_EDIT")
l004E = extfn1130(l0038)
l0054 = extfn11A0(l0048, 12, 0, "«‹- -–-—–––--–––––-- -¤ Attention ¤- -–-—–––--–––––-- -›»")
l005A = extfn1168(l0048, 2)
l0060 = fn1480(l005A)
DoEvents
l0064 = extfn11A0(l0048, 12, 0, l0034)
l0068 = fn1480(l005A)
DoEvents
l004E = extfn1130(l0038)
l006C = extfn11A0(l0048, 12, 0, "«‹- -–-—–––--–––––-- -¤ Attention ¤- -–-—–––--–––––-- -›»")
l0070 = fn1480(l005A)
End Sub

Sub Form_Load ()
extsub1050 attention.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub text1_KeyPress (KeyAscii As Integer)
Dim l007A As Variant
l007A = Chr(KeyAscii)
If  l007A = Chr(13) Then Command3D1_Click
End Sub
