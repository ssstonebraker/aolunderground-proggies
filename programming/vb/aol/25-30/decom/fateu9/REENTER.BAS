' REENTER.FRM
Option Explicit

Sub Check3D1_Click (Value As Integer)
If  Check3D1.Value = True Then
Text2.Enabled = True
Text2.Text = "10"
Else
Text2.Enabled = False
Text2.Text = "0"
End If
End Sub

Sub Command3D1_Click ()
Dim l002A As Variant
Dim l0030 As Variant
Dim l0034 As String
Dim l0038 As Variant
l002A = extfn1018("AOL FRAME25", 0&)
l0030 = extfn10F8(l002A, "Welcome")
l0034$ = String(30, 0)
l0038 = extfnFA8(l0030, l0034$, 250)
If  l0038 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "Please enter in the name of the room to freeze!", 16
Exit Sub
End If
If  Check3D1.Value = True Then
If  Text2.Text = "" Then
MsgBox "Please enter in the control speed!", 16
Exit Sub
End If
If  Not IsNumeric(Text2.Text) Then
MsgBox "Please enter in a numeric number for the control speed!", 16
Exit Sub
End If
End If
If  Timer1.Interval = 0 Then
Timer1.Enabled = True
Timer1.Interval = 1
Command3D1.Caption = "Stop"
Else
Timer1.Interval = 0
Timer1.Enabled = False
Command3D1.Caption = "Freeze"
End If
End Sub

Sub Form_Load ()
Check3D1.Value = True
Text2.Enabled = True
Text2.Text = "10"
List1.AddItem "Private"
List1.AddItem "Member"
List1.AddItem "Lobby"
List1.Selected(0) = True
extsub1050 reenter.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub List1_Click ()
Dim l0058 As Variant
l0058 = List1.ListIndex
If  l0058 = 0 Then type1.Caption = "1"
If  l0058 = 1 Then type1.Caption = "2"
If  l0058 = 2 Then type1.Caption = "3"
End Sub

Sub Timer1_Timer ()
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006C As Variant
Dim l0072 As Variant
Dim l0078 As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008C As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00AA As Variant
Dim l00AE As String
Dim l00B0 As Variant
l0062 = extfn1018("AOL FRAME25", 0&)
l0066 = extfn10F8(l0062, "AOL Live!")
l006C = extfn1088(l0066)
l0072 = extfn10C0(l006C, "_AOL_EDIT")
l0078 = fn138()
l007E = fn5D0("&Go To", "Edit Go To Menu")
If  Command3D1.Caption = "Freeze" Then GoTo 20

10 :
DoEvents
l0082 = extfn1018("_AOL_MODAL", "Favorite Places")
If  l0082 = 0 Then GoTo 10
l0086 = extfn10C0(l0082, "_AOL_EDIT")
l008C = extfn11A0(l0086, 12, 0, Text1)
l0092 = extfn1168(l0086, 2)
If  type1.Caption = "1" Then l0096 = extfn11A0(l0092, 12, 0, "aol://2719:2-2-" + Text1)
If  type1.Caption = "2" Then l0096 = extfn11A0(l0092, 12, 0, "aol://2719:61-2-" + Text1)
If  type1.Caption = "3" Then l0096 = extfn11A0(l0092, 12, 0, "aol://2719:21-2-" + Text1)
l009A = extfn10F8(l0082, "Save Changes")
If  Command3D1.Caption = "Freeze" Then GoTo 20
l00A0 = fn1480(l009A)
100 :
DoEvents
l0082 = extfn1018("_AOL_MODAL", "Favorite Places")
If  l0082 > 0 Then GoTo 100
Do Until 2 > 3
If  Check3D1.Value = True Then
l00A4 = l00A4 + 1
If  l00A4 = Int(Text2.Text) Then
l00AA = fn13A0(2.5#)
l00A4 = 0
End If
End If
If  Command3D1.Caption = "Freeze" Then GoTo 20
l00AE$ = Text1.Text
l00B0 = fn5D0("&Go To", l00AE$)
DoEvents
DoEvents
Loop
20 :
End Sub

