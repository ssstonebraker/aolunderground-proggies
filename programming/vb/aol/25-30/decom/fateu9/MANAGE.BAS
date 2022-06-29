' MANAGE.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1 = "" Then
MsgBox "You must enter in a Message", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
Timer1.Enabled = True
Timer1.Interval = 1
Command3D1.Caption = "On"
Else
Timer1.Interval = 0
Timer1.Enabled = False
Command3D1.Caption = "Off"
End If
End Sub

Sub Command3D2_Click ()
manage.WindowState = 1
End Sub

Sub Command3D3_Click ()
List1.Clear
List2.Clear
Text2.Text = ""
End Sub

Sub Form_Load ()
clo.Value = True
Option3D1.Value = True
extsub1050 manage.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub List1_Click ()
Dim l005E As Variant
l005E = List1.ListIndex
Text2.Text = List2.List(l005E)
End Sub

Sub List1_DblClick ()
Dim l0064 As Variant
Dim l006C As Variant
Dim l0072 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0086 As Variant
l0064 = extfn1018("AOL FRAME25", 0&)
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
l006C = fn5D0("Mem&bers", "Send an Instant Message")
13 :
DoEvents
l0072 = extfn10F8(l0064, "Send Instant Message")
If  l0072 = 0 Then GoTo 13
l0078 = extfn10C0(l0072, "_AOL_EDIT")
l007C = List1.ListIndex
l0080 = List1.List(l007C)
l0086 = extfn11A0(l0078, 12, 0, l0080)
End Sub

Sub Option3D1_Click (Value As Integer)
Text1.Text = "SuP, I am NOT at my computer right now. I'll contact you later when I am at my computer."
End Sub

Sub Option3D2_Click (Value As Integer)
Text1.Text = "I am in the bathroom right now. I'll contact you later when I am back."
End Sub

Sub Option3D3_Click (Value As Integer)
Text1.Text = "I am sleeping right now and will contact you later in the mornin."
End Sub

Sub Option3D4_Click (Value As Integer)
Text1.Text = "Sorry, I left my computer on and I am downloading something right now. I'll contact you later"
End Sub

Sub Option3D5_Click (Value As Integer)
Text1.Text = "Enter your custom message here."
End Sub

Sub Option3D6_Click (Value As Integer)
Text1.Text = "Hey, contact me later. I am outside right now and I'll get right back to you when I am available."
End Sub

Sub Timer1_Timer ()
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As String
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As String
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E4 As Variant
Dim l00E8 As Variant
Dim l00EC As Variant
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
10 :
If  Timer1.Interval = 0 Then GoTo 20
DoEvents
l0096 = extfn1018("AOL FRAME25", 0&)
l009A = extfn10F8(l0096, ">Instant Message From")
If  l009A = 0 Then GoTo 10
l009E$ = String(l009A, 0)
l00A2 = extfnFA8(l009A, l009E$, 250)
l00A6 = Mid(l009E$, 24, 10)
l00AA = InStr(l00A6, Chr(0))
If  l00AA Then
l00A6 = Mid(l00A6, 1, l00AA - 1)
End If
For l00AE = 0 To List1.ListCount - 1
l00B2 = List1.List(l00AE)
l00B6 = InStr(1, l00A6, l00B2, 1)
If  l00B6 Then
l00BA = Trim(l00B2)
l00BE = Trim(l00A6)
If  Len(l00BA) = Len(l00BE) Then
GoTo 10
DoEvents
End If
End If
If  Timer1.Interval = 0 Then GoTo 20
DoEvents
Next l00AE
DoEvents
List1.AddItem l00A6
If  fn138() = 5 Then
l00C2 = extfn10C0(l009A, "_AOL_VIEW")
Else
l00C2 = extfn10C0(l009A, "RICHCNTL")
End If
l00C6$ = String(255, 0)
l00C8 = extfn11A0(l00C2, 13, 250, l00C6$)
l00CC = InStr(l00C6$, Chr(0))
If  l00CC Then
l00C6$ = Mid(l00C6$, 1, l00CC - 1)
End If
l00CC = InStr(l00C6$, Chr(13))
If  l00CC Then
l00C6$ = Mid(l00C6$, 1, l00CC - 1)
End If
l00D0 = InStr(l00C6$, ":" + Chr(9))
If  l00D0 Then
l00C6$ = Mid(l00C6$, l00D0 + 2, Len(l00C6$))
End If
List2.AddItem l00C6$
On Error Resume Next
List1.Selected(0) = True
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
DoEvents
l00D4 = extfn10F8(l009A, "Respond")
l00DA = fn1480(l00D4)
l00DE = extfn10F8(l009A, "Send")
l00DE = extfn1168(l00DE, 3)
l00E4 = extfn11A0(l00DE, 12, 0, Text1)
l00E8 = extfn10F8(l009A, "Send")
l00EC = fn1480(l00E8)
If  clo.Value = True Then
l00F2 = extfnFE0(l009A, 16, 0, 0)
End If
DoEvents
GoTo 10
Else
DoEvents
l00D4 = extfn10C0(l009A, "_AOL_ICON")
l00D4 = extfn1168(l00D4, 1)
l00D4 = extfn1168(l00D4, 1)
l00D4 = extfn1168(l00D4, 3)
l00DA = fn1480(l00D4)
l00F6 = extfn10C0(l009A, "_AOL_ICON")
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00F6 = extfn1168(l00F6, 2)
l00FA = extfn11A0(l00F6, 12, 0, Text1.Text)
l00FE = extfn1168(l00FA, 2)
l0102 = fn1480(l00FE)
End If
If  clo.Value = True Then
l00F2 = extfnFE0(l009A, 16, 0, 0)
End If
DoEvents
GoTo 10

20 :
End Sub

Sub sub9980 (p0106 As Integer)
Select Case p0106
Case 0
Unload manage
Case 1
manage.WindowState = 1
End Select
End Sub
