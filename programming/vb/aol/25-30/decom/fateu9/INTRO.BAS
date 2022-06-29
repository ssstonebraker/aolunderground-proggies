' INTRO.FRM
Option Explicit

Sub Form_Load ()
Dim l001E As Variant
Dim l0022 As String
Dim l0024 As Variant
Dim l0028 As Variant
Dim l002C As Variant
Dim l0030 As Variant
Dim l0034 As Variant
Dim l003A As Variant
extsub1050 intro.hWnd, -1, 0, 0, 0, 0, 3
On Error Resume Next
SetAttr "fate.ini", 0
Open "fate.ini" For Binary As #2
l001E = LOF(2)
l0022$ = String(l001E, 0)
Get #2, 1, l0022$
Close #2
l0024 = InStr(1, l0022$, "Intro=", 1)
l0028 = Mid(l0022$, l0024 + 6, Len(l0022$))
l002C = InStr(l0028, Chr(13))
l0030 = Mid(l0028, 1, l002C - 1)
l0034 = InStr(1, l0030, "yes", 1)
If  l0034 Then
l003A = extfn1210("final.wav", 1)
Else
Unload intro
main.Show
End If
If  Err Then
sub3A0
End If
End Sub

Sub Picture1_Click ()
Unload intro
main.Show
End Sub
