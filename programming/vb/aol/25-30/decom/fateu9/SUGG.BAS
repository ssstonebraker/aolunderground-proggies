' SUGG.FRM
Option Explicit

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
Dim l0022 As Variant
Dim l0028 As Variant
Dim l002C As String
Dim l0030 As Variant
Dim l0034 As Variant
Dim l0038 As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004E As Variant
Dim l0054 As Variant
Dim l0058 As Variant
Dim l005C As Variant
Dim l0060 As Variant
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AC As Variant
If  sugger.Text = "" Then
MsgBox "You must enter in something to send!", 16
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
If  sugger.Text = "" Then
MsgBox "You must enter in something", 16
Exit Sub
End If
l0034 = extfn1018("AOL FRAME25", 0&)
l0038 = extfn1018("AOL FRAME25", 0&)
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l0040 = fn5D0("Mem&bers", "Set Preferences")
Else
l0040 = fn5D0("Mem&bers", "Preferences")
End If
101 :
DoEvents
l0044 = extfn10F8(l0038, "Preferences")
If  l0044 = 0 Then GoTo 101
l0048 = extfn10F8(l0044, "Mail")
l004E = extfn1168(l0048, 2)
l0054 = fn1480(l004E)
20 :
DoEvents
l0058 = extfn1018("_AOL_MODAL", 0&)
If  l0058 = 0 Then GoTo 20
l005C = extfn10F8(l0058, "close")
l0060 = extfn1168(l005C, 3)
l0066 = extfnFE0(l0060, 1025, 0, 0)
l006A = extfn10F8(l0058, "confirm")
l006E = extfn1168(l006A, 2)
l0072 = extfnFE0(l006E, 1025, 1, 0)
l0076 = extfn10F8(l0058, "OK")
l007A = fn1480(l0076)
l007E = extfnFE0(l0044, 16, 0, 0)
l0038 = fn5D0("&Mail", "&Compose Mail")

10 :
DoEvents
l0040 = extfn10F8(l0034, "Compose Mail")
If  l0040 = 0 Then GoTo 10
l0044 = extfn10F8(l0040, "To:")
l0048 = extfn1168(l0044, 2)
l004E = extfn11A0(l0048, 12, 0, "Fate-Ultra@JUNO.COM")
l0054 = extfn1168(l0044, 1)
If  fn560() = "1" Then
l0086 = "Master.aol = Yes"
Else
l0086 = "Master.aol = No"
End If
If  fn138() = 5 Then l008A = "America Online Version 2.5"
If  fn138() = 8 Then l008A = "America Online Version 3.0"
l0058 = extfn11A0(l0054, 12, 0, sugger + Chr(13) + Chr(10) + l008A + Chr(13) + Chr(10) + l0086 + Chr(13) + Chr(10) + "Fate Ultra Version 9.0 By: MaGuS")
l005C = extfn10F8(l0040, "Subject:")
l0060 = extfn1168(l005C, 2)
If  OPTION1.Value = True Then
l0066 = extfn11A0(l0060, 12, 0, "Suggestions")
End If
If  OPTION2.Value = True Then
l0066 = extfn11A0(l0060, 12, 0, "Comments")
End If
If  OPTION3.Value = True Then
l0066 = extfn11A0(l0060, 12, 0, "Bugs")
End If
l006A = extfn10F8(l0040, "Send")
l006E = extfn1168(l006A, 2)
l0072 = fn1480(l006E)

45 :
l00A0 = fn138()
If  l00A0 = 5 Then
l0076 = extfn1018("#32770", "America Online")
Else
l0076 = extfn1018("_AOL_MODAL", 0&)
End If
l007A = extfn10F8(l0034, "Error")
If  l007A > 0 Then
l007E = extfn10F8(l007A, "OK")
l00A4 = fn1480(l007E)
l00A8 = extfn11A0(l0048, 12, 0, "Fate-ultra@juno.com,Fate-d00d@juno.com")
l00AC = fn1480(l006E)
GoTo 45
End If
DoEvents
If  l0076 = 0 Then GoTo 45
l007A = extfn10F8(l0076, "OK")
l007E = fn1480(l007A)
End Sub

Sub Form_Load ()
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
extsub1050 sugg.hWnd, -1, 0, 0, 0, 0, 3
OPTION1.Value = True
End Sub

Sub Form_Resize ()
sugg.Width = 4545
sugg.Height = 4175
End Sub

Sub sub9980 ()
Unload sugg
End Sub
