' FREE.FRM
Option Explicit
Dim m001A As Variant

Sub Command3D1_Click ()
Dim l002C As Variant
Dim l0032 As Variant
Dim l0036 As String
Dim l003A As Variant
l002C = extfn1018("AOL FRAME25", 0&)
l0032 = extfn10F8(l002C, "Welcome")
l0036$ = String(30, 0)
l003A = extfnFA8(l0032, l0036$, 250)
If  l003A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
Timer1.Interval = 1
Command3D1.Caption = "Stop"
Else
Command3D1.Caption = "Start"
Timer1.Interval = 0
End If
End Sub

Sub Command3D2_Click ()
Dim l004E As Variant
help1.Show
For l004E = 0 To help1.List1.ListCount - 1
If  help1.List1.List(l004E) = "Free" Then
help1.List1.Selected(l004E) = True
End If
Next l004E
End Sub

Sub Form_Load ()
extsub1050 free.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l0056 As Variant
Dim l005C As Variant
Dim l0062 As Variant
Dim l0068 As Variant
Dim l006E As Variant
Dim l0074 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008C As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As Variant
Dim l00A4 As Variant
l0056 = extfn1018("AOL FRAME25", 0&)
l005C = extfn10C0(l0056, "MDICLIENT")
m001A = extfn1838(l005C)
l0062 = extfn1018("AOL FRAME25", 0&)
l0068 = extfn1790(l0062)
l006E = extfn17C8(l0068, 0)
l0074 = extfn1280(l006E)
l007A = fn5D0("&Go To", "Keyword...")

10 :
DoEvents
If  Timer1.Interval = 0 Then GoTo 50
l007E = extfn10F8(l0056, "Keyword")
If  l007E = 0 Then GoTo 10
l0082 = extfn10C0(l007E, "_AOL_ICON")
l0086 = extfn10C0(l007E, "_AOL_EDIT")
l008C = extfn11A0(l0086, 12, 0, "Billing")
l0092 = fn1480(l0082)
20 :
If  Timer1.Interval = 0 Then GoTo 50
DoEvents
l0096 = extfn1018("#32770", "America Online")
If  l0096 = 0 Then GoTo 20
l009A = extfn10F8(l0096, "&Yes")
l009E = fn1480(l009A)
30 :
DoEvents
If  Timer1.Interval = 0 Then GoTo 50
l00A4 = extfn1758(m001A, 1)
GoTo 30
50 :
End Sub

Sub sub9980 ()
Unload free
End Sub
