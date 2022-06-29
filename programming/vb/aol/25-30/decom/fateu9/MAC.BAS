' MAC.FRM
Option Explicit

Sub Check3D1_Click ()
End Sub

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
Dim l002E As Variant
Dim l0034 As Variant
Dim l0038 As String
Dim l003C As Variant
Dim l0040 As Variant
Dim l0046 As Variant
Dim l004A As Variant
Dim l004E As Variant
l002E = extfn1018("AOL FRAME25", 0&)
l0034 = extfn10F8(l002E, "Welcome")
l0038$ = String(30, 0)
l003C = extfnFA8(l0034, l0038$, 250)
If  l003C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0040 = InputBox("Enter in the room to immitate:", "- --Fate Ultra-- -", "WaReZ")
If  l0040 = "" Then GoTo 40
l0046 = fnA30(" ")
DoEvents
l004A = fnA30("*** You are in """ + l0040 + """. ***")
DoEvents
l004E = fnA30(" ")
DoEvents
40 :
End Sub

Sub Command3D13_Click ()
Dim l0058 As Variant
Dim l005E As Variant
If  Command3D13.Caption = "Gamma Ray" Then
Command3D13.Caption = "Stop"
10 :
If  Command3D13.Caption = "Gamma Ray" Then GoTo L2E87E
l0058 = fnA30("_.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._")
l0058 = fnA30("_.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._")
l0058 = fnA30("_.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._")
l0058 = fnA30("_.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._.-=-._")
l005E = fn13A0(2.8#)
If  Command3D13.Caption = "Gamma Ray" Then GoTo L2E87E
GoTo 10
Else
Command3D13.Caption = "Gamma Ray"
End If
L2E87E:
End Sub

Sub Command3D14_Click ()
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
If  Command3D14.Caption = "Shock Wave" Then
Command3D14.Caption = "Stop"
20 :
If  Command3D14.Caption = "Shock Wave" Then GoTo L2EAE6
l0068 = fnA30("^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^")
l006C = fnA30("^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^")
l0070 = fnA30("^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^")
l0074 = fnA30("^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^v^")
l0078 = fn13A0(2.8#)
If  Command3D14.Caption = "Shock Wave" Then GoTo L2EAE6
GoTo 20
Else
Command3D14.Caption = "Shock Wave"
End If
L2EAE6:
End Sub

Sub Command3D2_Click ()
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As String
Dim l0086 As Variant
l007C = extfn1018("AOL FRAME25", 0&)
l0080 = extfn10F8(l007C, "Welcome")
l0084$ = String(30, 0)
l0086 = extfnFA8(l0080, l0084$, 250)
If  l0086 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Command3D2.Caption = "Macro Killer" Then
Command3D2.Caption = "Unload"
macrokill.Show
Else
Command3D2.Caption = "Macro Killer"
Unload macrokill
End If
End Sub

Sub Command3D3_Click ()
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As String
Dim l009E As Variant
l0094 = extfn1018("AOL FRAME25", 0&)
l0098 = extfn10F8(l0094, "Welcome")
l009C$ = String(30, 0)
l009E = extfnFA8(l0098, l009C$, 250)
If  l009E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
Lamesense.Show
End Sub

Sub Command3D4_Click ()
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As String
Dim l00B0 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
l00A6 = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l00A6, "Welcome")
l00AE$ = String(30, 0)
l00B0 = extfnFA8(l00AA, l00AE$, 250)
If  l00B0 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00B6 = fnA68("     .    __ _______________ _")
l00B6 = fnA68("       _/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯l")
l00B6 = fnA68("      / ___  .   - -====--- - -_____l")
l00B6 = fnA68("     / /     /   __/¯¯¯¯¯¯¯. ¯¯¯¯¯ ¯")
l00BA = fn13A0(2.5#)
l00B6 = fnA68("   ./ /     /  //")
l00B6 = fnA68("  / /___/  /.")
l00B6 = fnA68(" l______l.")
l00B6 = fnA68("  ¯¯¯¯   ¯")
End Sub

Sub Command3D5_Click ()
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As String
Dim l00C8 As Variant
Dim l00D2 As Variant
Dim l00D6 As Variant
l00BE = extfn1018("AOL FRAME25", 0&)
l00C2 = extfn10F8(l00BE, "Welcome")
l00C6$ = String(30, 0)
l00C8 = extfnFA8(l00C2, l00C6$, 250)
If  l00C8 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Command3D5.Caption = "Wave" Then
Command3D5.Caption = "Unload"
Command3D5.ForeColor = QBColor(1)
35 :
If  Command3D5.Caption = "Wave" Then GoTo L2F55E
l00D2 = fnA68(".-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.")
l00D2 = fnA68(".-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.")
l00D2 = fnA68(".-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.")
l00D2 = fnA68(".-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.   .-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-..-=¯=-.")
l00D6 = fn13A0(2.8#)
If  Command3D5.Caption = "Wave" Then GoTo L2F55E
GoTo 35
Else
Command3D5.Caption = "Wave"
End If
L2F55E:
Command3D5.ForeColor = QBColor(0)
End Sub

Sub Command3D6_Click ()
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As String
Dim l00E4 As Variant
Dim l00E8 As Variant
l00DA = extfn1018("AOL FRAME25", 0&)
l00DE = extfn10F8(l00DA, "Welcome")
l00E2$ = String(30, 0)
l00E4 = extfnFA8(l00DE, l00E2$, 250)
If  l00E4 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00E8 = fnA30("                                                                                               ")
l00E8 = fnA30("                                                                                               ")
l00E8 = fnA30("                                                                                               ")
l00E8 = fnA30("                                                                                               ")
End Sub

Sub Command3D7_Click ()
Dim l00EC As Variant
Dim l00F0 As Variant
Dim l00F4 As String
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
l00EC = extfn1018("AOL FRAME25", 0&)
l00F0 = extfn10F8(l00EC, "Welcome")
l00F4$ = String(30, 0)
l00F6 = extfnFA8(l00F0, l00F4$, 250)
If  l00F6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00FA = fnA68("__                   _ _   _ _      _ ____      _____   ______")
l00FA = fnA68("\¯¯ \   _____   /¯¯/ /¯¯¯\     /l¯¯¯¯¯\   /l  ¯¯ ¯ l/l_ _      ll")
l00FA = fnA68("l \    \/¯  ¯¯¯ \/    / /   /\   \\  l l          l l l   l¯¯¯//__ /    /")
l00FA = fnA68(" \  \        /\        ///     ¯     \ l l    /\    \ l l   l¯¯¯/ ¯/     /_")
l00FE = fn13A0(2.8#)
l00FA = fnA68("   \  \___/  \___/ /__/\¯¯ /\__\ l__l  \__\l l   ¯¯¯ l/l______ll")
l00FA = fnA68("     \/___/\ /__/   \__\/¯¯ \/__//_ /  \/__/l/¯¯ ¯¯¯//__ ___/")
l00FA = fnA68("      ¯¯¯     ¯¯ ¯  ¯¯        ¯¯   ¯¯   ¯¯   ¯¯¯¯¯¯  ¯¯¯ ¯¯")
End Sub

Sub Command3D8_Click ()
MsgBox "Send me some suggestions on making this!", 16
End Sub

Sub Command3D9_Click ()
customkill.Show
End Sub

Sub Form_Load ()
extsub1050 mac.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub sub9980 ()
Unload mac
End Sub
