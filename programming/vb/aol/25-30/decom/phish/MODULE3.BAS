' Module3
Option Explicit

Sub sub4198 (p0008 As Variant)
Dim l000C As Variant
Dim l0010 As Variant
p0008 = p0008 * 1000
l000C = extfn2DA0()
Do While extfn2DA0() < l000C + p0008
l0010 = DoEvents()
Loop
End Sub
