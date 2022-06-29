' IF980C.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l002A As Variant
Dim l0034 As Variant
l002A = 0
If  Timer1.Interval = 0 Then
For l0034 = 0 To List1.ListCount - 1
If  List1.Selected(l0034) = True Then
l002A = l002A + 1
End If
Next l0034
If  l002A = 0 Then
MsgBox "You don't have any mail selected!", 16
Exit Sub
End If
Timer1.Interval = 1
Command3D1.Caption = "Stop"
Else
Unload pwait
Timer1.Interval = 0
Command3D1.Caption = "Start"
End If
End Sub

Sub Command3D2_Click ()
Dim l004C As Variant
help1.Show
For l004C = 0 To help1.List1.ListCount - 1
If  help1.List1.List(l004C) = "Mail Fix" Then
help1.List1.Selected(l004C) = True
Exit Sub
End If
Next l004C
End Sub

Sub Form_Load ()
Dim l0056 As Variant
Dim l005C As Variant
Dim l0062 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As String
Dim l0074 As Variant
extsub1050 if980c.hWnd, -1, 0, 0, 0, 0, 3
l0056 = extfn1018("AOL FRAME25", 0&)
l005C = extfn10F8(l0056, "New Mail")
l0062 = extfn10C0(l005C, "_AOL_TREE")
If  l0062 = 0 Then
MsgBox "Please Have Your Mailbox Open!", 16
Unload if980c
End If
l0068 = extfnFE0(l0062, 1036, 0, 0)
For l006C = 0 To l0068 - 1
l0070$ = String(255, 0)
l0074 = extfn11A0(l0062, 1034, l006C, l0070$)
List1.AddItem l0070$
Next l006C
End Sub

Sub HScroll1_Change ()
End Sub

Sub text1_KeyPress (p0078 As Integer)
Dim l007A As Variant
Dim l007E As Variant
l007A = Chr(p0078)
If  l007A = Chr(13) Then
List1.AddItem l007E
l007E = ""
End If
End Sub

Sub Timer1_Timer ()
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As String
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
If  Timer1.Interval = 0 Then GoTo 40
l0082 = extfn1018("AOL FRAME25", 0&)
l0086 = extfn10F8(l0082, "New Mail")
l008A = extfn10C0(l0086, "_AOL_TREE")
l008E = extfnFE0(l008A, 1036, 0, 0)
pwait.Show
For l0092 = 0 To l008E - 1
If  Timer1.Interval = 0 Then GoTo 40
10 :
l0096$ = String(255, 0)
l0098 = extfn11A0(l008A, 1034, l0092, l0096$)
l009C = List1.List(List1.ListIndex)
l00A0 = InStr(1, l0096$, l009C, 1)
If  l00A0 Then
l00A4 = extfnFE0(l008A, 1031, l0092, 0)
l00A8 = extfn10F8(l0086, "Delete")
l00AE = fn1480(l00A8)
l00B2 = l00B2 + 1
GoTo 10
End If
DoEvents
If  Timer1.Interval = 0 Then GoTo 40
Next l0092
40 :
If  l00B2 = "" Then l00B2 = 0
MsgBox l00B2 + " Mailbombs Deleted!"
pwait.Hide
Timer1.Interval = 0
Command3D1.Caption = "Start"
End Sub

Sub sub9980 ()
Unload if980c
End Sub
