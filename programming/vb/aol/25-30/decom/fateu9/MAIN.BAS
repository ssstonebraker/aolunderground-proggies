' MAIN.FRM
Option Explicit
Dim m001A As Variant
Dim m002A As Variant
Dim m003A As Variant
Dim m004A As Variant

Sub addroo_Click ()
Dim l006C As Variant
Dim l0072 As Variant
Dim l0076 As String
Dim l007A As Variant
l006C = extfn1018("AOL FRAME25", 0&)
l0072 = extfn10F8(l006C, "Welcome")
l0076$ = String(30, 0)
l007A = extfnFA8(l0072, l0076$, 250)
If  l007A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
addroom.Show
End Sub

Sub sub1D60 ()
End Sub

Sub afkbot_Click ()
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As String
Dim l008C As Variant
l0082 = extfn1018("AOL FRAME25", 0&)
l0086 = extfn10F8(l0082, "Welcome")
l008A$ = String(30, 0)
l008C = extfnFA8(l0086, l008A$, 250)
If  l008C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
afkb.Show
End Sub

Sub atten_Click ()
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
attention.Show
End Sub

Sub chatb_Click ()
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As String
Dim l00B0 As Variant
l00A6 = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l00A6, "Welcome")
l00AE$ = String(30, 0)
l00B0 = extfnFA8(l00AA, l00AE$, 250)
If  l00B0 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
kills.Show
End Sub

Sub sub1E40 ()
End Sub

Sub sub1E78 ()
End
End Sub

Sub sub1EB0 ()
End Sub

Sub sub1EE8 ()
End Sub

Sub sub1F20 ()
End Sub

Sub sub1F58 ()
End Sub

Sub sub1F90 ()
End Sub

Sub Command3D1_Click ()
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As String
Dim l00C2 As Variant
Dim l00C8 As Variant
l00B8 = extfn1018("AOL FRAME25", 0&)
l00BC = extfn10F8(l00B8, "Welcome")
l00C0$ = String(30, 0)
l00C2 = extfnFA8(l00BC, l00C0$, 250)
If  l00C2 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00C8 = fn480()
End Sub

Sub Command3D10_Click ()
Dim l00CC As Variant
Dim l00D2 As Variant
Dim l00DE As Variant
l00CC = extfn1018("AOL FRAME25", 0&)
l00D2 = extfn16E8(l00CC)
If  l00D2 = 0 Then
Command3D10.Caption = "Hide AOL"
l00DE = extfn1758(l00CC, 2)
l00DE = extfn1758(l00CC, 3)
Else
Command3D10.Caption = "Show AOL"
l00DE = extfn1758(l00CC, 0)
End If
End Sub

Sub Command3D12_Click ()
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As String
Dim l00EC As Variant
Dim l00F0 As Variant
Dim l00F4 As Variant
Dim l00FA As Variant
Dim l0100 As Variant
Dim l0106 As Variant
Dim l010A As Variant
Dim l010E As Variant
Dim l0114 As Variant
Dim l0118 As Variant
Dim l011C As Variant
Dim l0120 As Variant
l00E2 = extfn1018("AOL FRAME25", 0&)
l00E6 = extfn10F8(l00E2, "Welcome")
l00EA$ = String(30, 0)
l00EC = extfnFA8(l00E6, l00EA$, 250)
If  l00EC <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l00F0 = extfn1018("AOL FRAME25", 0&)
l00F4 = extfn10F8(l00F0, "AOL Live!")
l00FA = extfn1088(l00F4)
l0100 = extfn10C0(l00FA, "_AOL_EDIT")
m002A = fnA68(" __ _                  __ _")
l0106 = fnA68("l¯¯¯l\                 l¯¯¯l\")
l010A = fnA68("l     l ll         __   l     l  l")
l010E = fnA68("l     l_l _    /l¯¯l\\ l     l_l _")
l0114 = fn13A0(2.5#)
l0118 = fnA68("l______ l\ l/ ¯¯ \l l______ l\")
l011C = fnA68(" \______\ll \ __ /  \______\ll")
l0120 = fnA68("  ¯¯¯ ¯ ¯     ¯ ¯     ¯¯¯ ¯ ¯")
End Sub

Sub Command3D13_Click ()
Dim l0126 As Variant
Dim l012A As Variant
Dim l012E As String
Dim l0130 As Variant
Dim l0134 As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Variant
Dim l0148 As Variant
Dim l014C As Variant
Dim l0150 As Variant
Dim l0156 As Variant
Dim l015A As Variant
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l0126 = extfn1018("AOL FRAME25", 0&)
l012A = extfn10F8(l0126, "Welcome")
l012E$ = String(30, 0)
l0130 = extfnFA8(l012A, l012E$, 250)
If  l0130 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0134 = extfn1018("AOL FRAME25", 0&)
l013A = extfnFE0(l0134, 16, 0, 0)
12 :
DoEvents
l013E = extfn1018("_AOL_MODAL", 0&)
If  l013E = 0 Then GoTo 12
l0142 = extfn10F8(l013E, "&Yes")
m002A = fn1480(l0142)
Do Until 2 > 3
l0148 = extfn1018("#32770", "Download Manager")
If  l0148 > 0 Then
l014C = extfn10F8(l0148, "&No")
l0150 = fn1480(l014C)
End If
l014C = extfn10F8(l0134, "Goodbye")
If  l014C > 0 Then
Exit Sub
End If
DoEvents
Loop
Else
l0126 = extfn1018("AOL FRAME25", 0&)
l012A = extfn10F8(l0126, "Welcome")
l012E$ = String(30, 0)
l0130 = extfnFA8(l012A, l012E$, 250)
If  l0130 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0134 = extfn1018("AOL FRAME25", 0&)
l013A = extfnFE0(l0134, 16, 0, 0)
29 :
DoEvents
l013E = extfn10F8(l0134, "Exit?")
If  l013E = 0 Then GoTo 29
l0142 = extfn10C0(l013E, "_AOL_icon")
m002A = extfn1168(l0142, 2)
l0148 = extfn1168(m002A, 2)
l014C = extfn1168(l0148, 2)
l0150 = extfn1168(l014C, 2)
l0156 = extfn1168(l0150, 2)
l015A = fn1480(l0156)
End If
End Sub

Sub Command3D14_Click ()
End Sub

Sub Command3D15_Click ()
guest.Show
End Sub

Sub Command3D16_Click ()
newuser.Show
End Sub

Sub Command3D19_Click ()
Dim l0166 As Variant
Dim l016A As Variant
Dim l016E As String
Dim l0170 As Variant
Dim l0176 As Variant
Dim l017A As Variant
l0166 = extfn1018("AOL FRAME25", 0&)
l016A = extfn10F8(l0166, "Welcome")
l016E$ = String(30, 0)
l0170 = extfnFA8(l016A, l016E$, 250)
If  l0170 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0176 = fnA30("____  _   __      ___ _  _____ _   ____ _ _ _______ _")
l0176 = fnA30("l¯¯¯¯¯¯¯¯¯l\    /¯¯¯¯ \l¯¯¯¯¯¯¯¯¯¯¯¯¯ l\¯¯¯¯¯¯¯¯¯¯l\")
l0176 = fnA30("l        ____l  \ /     _    \____        ____ l  \      l¯¯¯¯¯ \ \")
l0176 = fnA30("l       l\       \  l     /_\     \     l       l \      \  l       ¯¯¯ l\__\l")
l017A = Timer: While Timer - l017A < 2.5#: DoEvents: Wend
l0176 = fnA30("l        ¯¯¯l\ _\l     __       \__l       l   \___\l      l¯¯¯¯\ \")
l0176 = fnA30("l       l\¯¯ \ \      /\ ¯\        \  .l       l   l¯¯.l       ¯¯¯¯¯¯l\")
l0176 = fnA30("l       l  \__ \l__/    \   \____\ .l____l   l    l__________l  \")
l0176 = fnA30("\¯¯¯¯\  l ¯ \      \   /¯¯ \       \ \       \  l     \                 \ l")
l017A = Timer: While Timer - l017A < 2.5#: DoEvents: Wend
l0176 = fnA30("  \____\l     \ ___\/        \___ \  \____\l       \_________\l")
l0176 = fnA30("    ¯¯   ¯       ¯  ¯            ¯¯¯ ¯  ¯¯¯¯ ¯¯       ¯¯¯¯¯¯¯¯ ¯")
l0176 = fnA30("  «‹- -– Fate Ultra 9.0 By MaGuS and FuNGii –-- -›»")
End Sub

Sub Command3D2_Click ()
Dim l017E As Variant
Dim l0182 As Variant
Dim l0186 As String
Dim l0188 As Variant
Dim l018E As Variant
l017E = extfn1018("AOL FRAME25", 0&)
l0182 = extfn10F8(l017E, "Welcome")
l0186$ = String(30, 0)
l0188 = extfnFA8(l0182, l0186$, 250)
If  l0188 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l018E = fn448()
End Sub

Sub sub21C0 ()
End Sub

Sub Command3D26_Click ()
Dim l0192 As Variant
Dim l0196 As Variant
l0192 = extfn1018("_AOL_MODAL", 0&)
If  l0192 > 0 Then
l0196 = extfnFE0(l0192, 16, 0, 0)
End If
l0192 = extfn1018("#32770", 0&)
If  l0192 > 0 Then
l0196 = extfnFE0(l0192, 16, 0, 0)
End If
End Sub

Sub sub2230 ()
End Sub

Sub Command3D3_Click ()
Dim l019A As Variant
Dim l019E As Variant
Dim l01A2 As String
Dim l01A4 As Variant
Dim l01A8 As Variant
Dim l01AE As Variant
Dim l01B4 As Variant
Dim l01BE As Variant
l019A = extfn1018("AOL FRAME25", 0&)
l019E = extfn10F8(l019A, "Welcome")
l01A2$ = String(30, 0)
l01A4 = extfnFA8(l019E, l01A2$, 250)
If  l01A4 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l01A8 = extfn1018("AOL FRAME25", 0&)
l01AE = extfn1720(l01A8)
If  l01AE = 0 Then
extsub1678 (l01A8)
l01B4 = extfn16B0(l01A8, 1)
Command3D3.Caption = "Un Upchat"
Else
l01BE = extfn1018("_AOL_MODAL", 0&)
extsub1678 (l01BE)
l01BE = extfn16B0(l01BE, 1)
Command3D3.Caption = "Upchat"
End If
End Sub

Sub Command3D4_Click ()
main.Hide
mini.Show
End Sub

Sub sub22D8 ()
End Sub

Sub Command3D5_Click ()
Dim l01CA As Variant
Dim l01CE As Variant
Dim l01D2 As String
Dim l01D4 As Variant
Dim l01D8 As Variant
Dim l01DC As Variant
Dim l01E0 As Variant
l01CA = extfn1018("AOL FRAME25", 0&)
l01CE = extfn10F8(l01CA, "Welcome")
l01D2$ = String(30, 0)
l01D4 = extfnFA8(l01CE, l01D2$, 250)
If  l01D4 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l01D8 = extfn1018("AOL FRAME25", 0&)
l01DC = extfn10F8(l01D8, "Welcome, ")
l01E0 = extfn16E8(l01DC)
If  l01E0 = 1 Then
Command3D5.Caption = "Show Welc."
l01E0 = extfn1758(l01DC, 0)
Else
Command3D5.Caption = "Kill Welc."
l01E0 = extfn1758(l01DC, 2)
End If
End Sub

Sub Command3D6_Click ()
Dim l01EC As Variant
l01EC = extfn1210("cmbultra.wav", 1)
End
End Sub

Sub Command3D7_Click ()
Dim l01F0 As Variant
Dim l01F4 As Variant
Dim l01F8 As String
Dim l01FA As Variant
Dim l0200 As Variant
Dim l0204 As Variant
Dim l0208 As Variant
l01F0 = extfn1018("AOL FRAME25", 0&)
l01F4 = extfn10F8(l01F0, "Welcome")
l01F8$ = String(30, 0)
l01FA = extfnFA8(l01F4, l01F8$, 250)
If  l01FA <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  fn138() = 5 Then
l0200 = fn5D0("&Go To", "Exit Free Area")
Else
l0200 = fn5D0("&Go To", "Exit Unlimited Use area")
End If
l0204 = extfn1018("#32770", "America Online")
If  l0204 > 0 Then
l0200 = extfn10F8(l0204, "&No")
l0208 = fn1480(l0200)
End If
End Sub

Sub Command3D8_Click ()
options.Show
End Sub

Sub Command3D9_Click ()
Dim l0210 As Variant
Dim l0214 As Variant
Dim l0218 As String
Dim l021A As Variant
l0210 = extfn1018("AOL FRAME25", 0&)
l0214 = extfn10F8(l0210, "Welcome")
l0218$ = String(30, 0)
l021A = extfnFA8(l0214, l0218$, 250)
If  l021A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
On Error Resume Next
sugg.Show
End Sub

Sub sub2428 ()
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022C As Variant
Dim l0230 As Variant
l0222 = extfn1018("AOL FRAME25", 0&)
l0226 = extfn10C0(l0222, "_AOL_ListBox")
l022C = extfn1130(l0226)
l0230 = extfnFE0(l0226, &H201, 0, 0)
m002A = extfnFE0(l0226, &H202, 0, 0)
End Sub

Sub sub2460 ()
Dim l0234 As Variant
Dim l0238 As Variant
Dim l023C As String
Dim l023E As Variant
Dim l0242 As Variant
l0234 = extfn1018("AOL FRAME25", 0&)
l0238 = extfn10F8(l0234, "Welcome")
l023C$ = String$(30, 0)
l023E = extfnFA8(l0238, l023C$, 250)
If  l023E <= 7 Then
MsgBox "Not Signed On!", 16, "Fate"
Exit Sub
End If
m002A = l023E - 1
l0242 = Mid(Left(l023C$, m002A), 10)
End Sub

Sub sub2498 ()
End Sub

Sub countm_Click ()
Dim l0246 As Variant
Dim l024A As Variant
Dim l024E As String
Dim l0250 As Variant
l0246 = extfn1018("AOL FRAME25", 0&)
l024A = extfn10F8(l0246, "Welcome")
l024E$ = String(30, 0)
l0250 = extfnFA8(l024A, l024E$, 250)
If  l0250 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
countmm.Show
End Sub

Sub credi_Click ()
credits.Show
End Sub

Sub sub2540 ()
End Sub

Sub sub2578 ()
End Sub

Sub sub25B0 ()
End Sub

Sub sub25E8 ()
End Sub

Sub dupek_Click ()
Dim l0264 As Variant
Dim l0268 As Variant
Dim l026C As String
Dim l026E As Variant
l0264 = extfn1018("AOL FRAME25", 0&)
l0268 = extfn10F8(l0264, "Welcome")
l026C$ = String(30, 0)
l026E = extfnFA8(l0268, l026C$, 250)
If  l026E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
DUPEKILL.Show
End Sub

Sub echob_Click ()
Dim l0276 As Variant
Dim l027A As Variant
Dim l027E As String
Dim l0280 As Variant
l0276 = extfn1018("AOL FRAME25", 0&)
l027A = extfn10F8(l0276, "Welcome")
l027E$ = String(30, 0)
l0280 = extfnFA8(l027A, l027E$, 250)
If  l0280 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
ECHO.Show
End Sub

Sub sub2690 ()
Dim l0288 As Variant
Dim l028C As Variant
Dim l0290 As String
Dim l0292 As Variant
l0288 = extfn1018("AOL FRAME25", 0&)
l028C = extfn10F8(l0288, "Welcome")
l0290$ = String(30, 0)
l0292 = extfnFA8(l028C, l0290$, 250)
If  l0292 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
el354.Show
End Sub

Sub elitem_Click ()
Dim l029A As Variant
Dim l029E As Variant
Dim l02A2 As String
Dim l02A4 As Variant
l029A = extfn1018("AOL FRAME25", 0&)
l029E = extfn10F8(l029A, "Welcome")
l02A2$ = String(30, 0)
l02A4 = extfnFA8(l029E, l02A2$, 250)
If  l02A4 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
el354.Show
End Sub

Sub sub2700 ()
encrypt.Show
End Sub

Sub encryptch_Click ()
Dim l02AC As Variant
Dim l02B0 As Variant
Dim l02B4 As String
Dim l02B6 As Variant
l02AC = extfn1018("AOL FRAME25", 0&)
l02B0 = extfn10F8(l02AC, "Welcome")
l02B4$ = String(30, 0)
l02B6 = extfnFA8(l02B0, l02B4$, 250)
If  l02B6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
encryptchat.Show
End Sub

Sub encrypter_Click ()
encrypt.Show
End Sub

Sub facc_Click ()
fac.Show
End Sub

Sub sub27E0 ()
End Sub

Sub Fakef_Click ()
Dim l02C2 As Variant
Dim l02C6 As Variant
Dim l02CA As String
Dim l02CC As Variant
l02C2 = extfn1018("AOL FRAME25", 0&)
l02C6 = extfn10F8(l02C2, "Welcome")
l02CA$ = String(30, 0)
l02CC = extfnFA8(l02C6, l02CA$, 250)
If  l02CC <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
fake.Show
End Sub

Sub sub2850 ()
Dim l02D4 As Variant
Dim l02D8 As Variant
Dim l02DC As String
Dim l02DE As Variant
l02D4 = extfn1018("AOL FRAME25", 0&)
l02D8 = extfn10F8(l02D4, "Welcome")
l02DC$ = String(30, 0)
l02DE = extfnFA8(l02D8, l02DC$, 250)
If  l02DE <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
fightbot.Show
End Sub

Sub findroom_Click ()
Dim l02E6 As Variant
Dim l02EA As Variant
Dim l02EE As String
Dim l02F0 As Variant
l02E6 = extfn1018("AOL FRAME25", 0&)
l02EA = extfn10F8(l02E6, "Welcome")
l02EE$ = String(30, 0)
l02F0 = extfnFA8(l02EA, l02EE$, 250)
If  l02F0 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
roomrun.Show
End Sub

Sub Form_Load ()
Dim l02F8 As Variant
Dim l02FC As Variant
Dim l0302 As Variant
Dim l0308 As Variant
Dim l030C As Variant
Dim l0314 As Variant
Dim l0318 As Variant
Dim l031C As String
Dim l031E As Variant
Dim l0322 As Variant
Dim l0326 As Variant
Dim l032A As Variant
Dim l032E As String
Dim l0330 As Variant
Dim l0334 As Variant
Dim l0338 As Variant
Dim l033C As Variant
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As Variant
Dim l0350 As Variant
Dim l0354 As Variant
Dim l0358 As Variant
Dim l035C As Variant
Dim l0360 As Variant
Dim l0364 As Variant
Dim l0368 As Variant
Dim l036C As Variant
Dim l0370 As Variant
Dim l0374 As Variant
l02F8 = extfn1018(0&, "X-Zone")
If  l02F8 > 0 Then
l02FC = extfnFE0(l02F8, 16, 0, 0)
MsgBox "Just Killed X-Zone!", 16
End If
If  main.Caption <> fn4F0() Then GoTo L173A8
l0302 = extfn1018("AOL FRAME25", 0&)
l0308 = extfn1790(l0302)
l030C = extfn17C8(l030C, 1)
l0314 = extfn1800(l030C, 1)
l0318 = extfn10F8(l0302, "Welcome")
l031C$ = String(30, 0&)
l031E = extfnFA8(l0318, l031C$, 250)
m002A = Mid(l031C$, 10, 10)
l0322 = InStr(m002A, "!")
If  l0322 Then
m002A = Mid(m002A, 1, l0322 - 1)
End If
l0326 = Mid(m002A, 1, 8)
l032A = InStr(1, l0326, "catwatch", 1)
If  l032A Then
End
End If
l0326 = Mid(m002A, 1, 5)
l032A = InStr(1, l0326, "steve", 1)
If  l032A Then
End
End If
l032A = InStr(1, m002A, "hein", 1)
If  l032A Then
MsgBox "Sorry HeiN, your not allowed in!", 16
End
End If
l0326 = Mid(m002A, 1, 5)
l032A = InStr(1, l0326, "guide", 1)
If  l032A Then
End
End If
l0326 = Mid(m002A, 1, 3)
l032A = InStr(1, l0326, "Tos", 1)
If  l032A Then
End
End If
extsub1050 main.hWnd, -1, 0, 0, 0, 0, 3
On Error Resume Next
SetAttr "fate.ini", 0
Open "fate.ini" For Binary As #2
l0302 = LOF(2)
l032E$ = String(l0302, 0)
Get #2, 1, l032E$
Close #2
l0330 = InStr(1, l032E$, "WAOLDir=", 1)
l031E = Mid(l032E$, l0330 + 8, Len(l032E$))
m002A = InStr(l031E, Chr(13))
l0322 = Mid(l031E, 1, m002A - 1)
l0334 = l0322
l0326 = InStr(1, l032E$, "Intro=", 1)
l032A = Mid(l032E$, l0326 + 6, Len(l032E$))
l0338 = InStr(l032A, Chr(13))
l033C = Mid(l032A, 1, l0338 - 1)
l0340 = InStr(1, l033C, "yes", 1)
If  l0340 Then
Else
l0302 = extfn1210("final.wav", 1)
End If
l0344 = InStr(1, l032E$, "Load=", 1)
l0348 = Mid(l032E$, l0344 + 5, Len(l032E$))
l034C = InStr(l0348, Chr(13))
m004A = Mid(l0348, 1, l034C - 1)
l0350 = InStr(1, m004A, "yes", 1)
If  l0350 Then
On Error Resume Next
If  l0334 = "" Then GoTo 44
l0354 = Shell(l0334 + "\waol.exe", 1)
If  Err Then
MsgBox "An error occurred: AOL directory doesn't exist", 16
End If
Else
End If

44 :
l0358 = InStr(1, l032E$, "times=", 1)
l035C = Mid(l032E$, l0358 + 6, Len(l032E$))
l0360 = InStr(l035C, Chr(13))
l0364 = Mid(l035C, 1, l0360 - 1)
m003A = Int(l0364)
If  m003A = 0 Then
options.Show
MsgBox "Thank you for choosing Fate Ultra. For all of you who didn't know, this program works for America Online 2.5/3.0 with or without Master.aol loaded.", 64
GoTo 19
End If
19 :
l0368 = m003A + 1
l036C = Str(l0368)
l0370 = InStr(1, l032E$, "Times=", 1)
m001A = Mid(l032E$, 1, l0370 - 1)
l02F8 = InStr(1, l032E$, "Load=", 1)
l02FC = Mid(l032E$, l02F8, Len(l032E$))
Open "fate.ini" For Output As #4
l0374 = Chr(13) + Chr(10)
Print #4, m001A + "Times=" + Trim(l036C) + Chr(13) + Chr(10) + l02FC
Close #4
GoTo L17420
L173A8:
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
L17420:
If  Err Then
sub3A0
End If
End Sub

Sub freewarez_Click ()
free.Show
End Sub

Sub getonm_Click ()
Dim l037C As Variant
Dim l0380 As Variant
Dim l0384 As String
Dim l0386 As Variant
l037C = extfn1018("AOL FRAME25", 0&)
l0380 = extfn10F8(l037C, "Welcome")
l0384$ = String(30, 0)
l0386 = extfnFA8(l0380, l0384$, 250)
If  l0386 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
getmm.Show
End Sub

Sub sub2968 ()
Dim l038E As Variant
Dim l0392 As Variant
Dim l0396 As String
Dim l0398 As Variant
l038E = extfn1018("AOL FRAME25", 0&)
l0392 = extfn10F8(l038E, "Welcome")
l0396$ = String(30, 0)
l0398 = extfnFA8(l0392, l0396$, 250)
If  l0398 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
smartbot.Show
End Sub

Sub greets_Click ()
Dim l03A0 As Variant
Dim l03A4 As Variant
Dim l03A8 As String
Dim l03AA As Variant
l03A0 = extfn1018("AOL FRAME25", 0&)
l03A4 = extfn10F8(l03A0, "Welcome")
l03A8$ = String(30, 0)
l03AA = extfnFA8(l03A4, l03A8$, 250)
If  l03AA <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
On Error Resume Next
smartbot.Show
End Sub

Sub helper_Click (Index%)
Dim l03B0 As Variant
Dim l03B4 As Variant
Dim l03B8 As String
Dim l03BA As Variant
Dim l03BE As Variant
Dim l03C2 As Variant
Dim l03C6 As Variant
Select Case Index
Case 0
Case 1
l03B0 = extfn1018("AOL FRAME25", 0&)
l03B4 = extfn10F8(l03B0, "Welcome")
l03B8$ = String(30, 0)
l03BA = extfnFA8(l03B4, l03B8$, 250)
If  l03BA <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  fn138() = 5 Then
l03BE = fn5D0("&Go To", "Exit Free Area")
Else
l03BE = fn5D0("&Go To", "Exit Unlimited Use area")
End If
l03C2 = extfn1018("#32770", "America Online")
If  l03C2 > 0 Then
l03BE = extfn10F8(l03C2, "&No")
l03C6 = fn1480(l03BE)
End If
End Select
End Sub

Sub ib_Click ()
Dim l03CA As Variant
Dim l03CE As Variant
Dim l03D2 As String
Dim l03D4 As Variant
l03CA = extfn1018("AOL FRAME25", 0&)
l03CE = extfn10F8(l03CA, "Welcome")
l03D2$ = String(30, 0)
l03D4 = extfnFA8(l03CE, l03D2$, 250)
If  l03D4 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
invitebo.Show
End Sub

Sub ima_Click ()
Dim l03DC As Variant
Dim l03E0 As Variant
Dim l03E4 As String
Dim l03E6 As Variant
l03DC = extfn1018("AOL FRAME25", 0&)
l03E0 = extfn10F8(l03DC, "Welcome")
l03E4$ = String(30, 0)
l03E6 = extfnFA8(l03E0, l03E4$, 250)
If  l03E6 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
manage.Show
End Sub

Sub imb_Click ()
Dim l03EE As Variant
Dim l03F2 As Variant
Dim l03F6 As String
Dim l03F8 As Variant
l03EE = extfn1018("AOL FRAME25", 0&)
l03F2 = extfn10F8(l03EE, "Welcome")
l03F6$ = String(30, 0)
l03F8 = extfnFA8(l03F2, l03F6$, 250)
If  l03F8 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
punt.Show
End Sub

Sub sub2AB8 ()
End Sub

Sub imi_Click ()
Dim l0400 As Variant
Dim l0404 As Variant
Dim l0408 As String
Dim l040A As Variant
l0400 = extfn1018("AOL FRAME25", 0&)
l0404 = extfn10F8(l0400, "Welcome")
l0408$ = String(30, 0)
l040A = extfnFA8(l0404, l0408$, 250)
If  l040A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
ignore.Show
End Sub

Sub sub2B28 ()
End Sub

Sub isitmagus_Click ()
verify.Show
End Sub

Sub jaap_Click ()
jphish.Show
End Sub

Sub jmb_Click ()
Junobomb.Show
End Sub

Sub lilhelper_Click (Index%)
help1.Show
help1.List1.Selected(0) = True
End Sub

Sub sub2C40 ()
Dim l0424 As Variant
Dim l0428 As Variant
Dim l042C As String
Dim l042E As Variant
l0424 = extfn1018("AOL FRAME25", 0&)
l0428 = extfn10F8(l0424, "Welcome")
l042C$ = String(30, 0)
l042E = extfnFA8(l0428, l042C$, 250)
If  l042E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
sugg.Show
End Sub

Sub mailb_Click ()
Dim l0432 As Variant
Dim l0436 As Variant
Dim l043A As String
Dim l043C As Variant
l0432 = extfn1018("AOL FRAME25", 0&)
l0436 = extfn10F8(l0432, "Welcome")
l043A$ = String(30, 0)
l043C = extfnFA8(l0436, l043A$, 250)
If  l043C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
mail.Show
End Sub

Sub mailf_Click ()
Dim l0444 As Variant
Dim l0448 As Variant
Dim l044C As String
Dim l044E As Variant
l0444 = extfn1018("AOL FRAME25", 0&)
l0448 = extfn10F8(l0444, "Welcome")
l044C$ = String(30, 0)
l044E = extfnFA8(l0448, l044C$, 250)
If  l044E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
On Error Resume Next
if980c.Show
End Sub

Sub mailma_Click ()
Dim l0456 As Variant
Dim l045A As Variant
Dim l045E As String
Dim l0460 As Variant
l0456 = extfn1018("AOL FRAME25", 0&)
l045A = extfn10F8(l0456, "Welcome")
l045E$ = String(30, 0)
l0460 = extfnFA8(l045A, l045E$, 250)
If  l0460 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
maillist.Show
End Sub

Sub sub2D20 ()
End Sub

Sub mailz_Click ()
Dim l0468 As Variant
Dim l046C As Variant
Dim l0470 As String
Dim l0472 As Variant
l0468 = extfn1018("AOL FRAME25", 0&)
l046C = extfn10F8(l0468, "Welcome")
l0470$ = String(30, 0)
l0472 = extfnFA8(l046C, l0470$, 250)
If  l0472 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
search.Show
End Sub

Sub sub2D90 ()
End Sub

Sub massem_Click ()
Dim l047A As Variant
Dim l047E As Variant
Dim l0482 As String
Dim l0484 As Variant
l047A = extfn1018("AOL FRAME25", 0&)
l047E = extfn10F8(l047A, "Welcome")
l0482$ = String(30, 0)
l0484 = extfnFA8(l047E, l0482$, 250)
If  l0484 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
MAILBOX.Show
End Sub

Sub massim_Click ()
Dim l048C As Variant
Dim l0490 As Variant
Dim l0494 As String
Dim l0496 As Variant
l048C = extfn1018("AOL FRAME25", 0&)
l0490 = extfn10F8(l048C, "Welcome")
l0494$ = String(30, 0)
l0496 = extfnFA8(l0490, l0494$, 250)
If  l0496 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
imform.Show
End Sub

Sub masteraola_Click ()
unloads.Show
End Sub

Sub sub2E70 ()
Dim l04A2 As Variant
Dim l04A6 As Variant
Dim l04AA As String
Dim l04AC As Variant
Dim l04B0 As Variant
l04A2 = extfn1018("AOL FRAME25", 0&)
l04A6 = extfn10F8(l04A2, "Welcome")
l04AA$ = String(30, 0)
l04AC = extfnFA8(l04A6, l04AA$, 250)
If  l04AC <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
On Error Resume Next
l04B0 = Shell("mchat30.exe", 1)
If  Err Then
MsgBox "M-Chat Could Not Be Found!", 16
Exit Sub
End If
End Sub

Sub mdraw_Click ()
drawman.Show
End Sub

Sub minimax_Click ()
If  minimax.Caption = "Roll Up" Then
minimax.Caption = "Roll Down"
main.Height = 580
Load helper(1)
helper(1).Caption = "Kill Wait"
Else
Unload helper(1)
minimax.Caption = "Roll Up"
main.Height = 1475
End If
End Sub

Sub sub2F18 ()
Dim l04C4 As Variant
Dim l04C8 As Variant
Dim l04CC As String
Dim l04CE As Variant
l04C4 = extfn1018("AOL FRAME25", 0&)
l04C8 = extfn10F8(l04C4, "Welcome")
l04CC$ = String(30, 0)
l04CE = extfnFA8(l04C8, l04CC$, 250)
If  l04CE <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
mmER.Show
End Sub

Sub Mmers_Click ()
Dim l04D6 As Variant
Dim l04DA As Variant
Dim l04DE As String
Dim l04E0 As Variant
l04D6 = extfn1018("AOL FRAME25", 0&)
l04DA = extfn10F8(l04D6, "Welcome")
l04DE$ = String(30, 0)
l04E0 = extfnFA8(l04DA, l04DE$, 250)
If  l04E0 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
mmER.Show
End Sub

Sub multim_Click ()
Dim l04E4 As Variant
Dim l04E8 As Variant
Dim l04EC As String
Dim l04EE As Variant
l04E4 = extfn1018("AOL FRAME25", 0&)
l04E8 = extfn10F8(l04E4, "Welcome")
l04EC$ = String(30, 0)
l04EE = extfnFA8(l04E8, l04EC$, 250)
If  l04EE <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
multimb.Show
End Sub

Sub sub2FC0 ()
Dim l04F6 As Variant
Dim l04FA As Variant
Dim l04FE As String
Dim l0500 As Variant
l04F6 = extfn1018("AOL FRAME25", 0&)
l04FA = extfn10F8(l04F6, "Welcome")
l04FE$ = String(30, 0)
l0500 = extfnFA8(l04FA, l04FE$, 250)
If  l0500 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
phish.Show
End Sub

Sub phisher_Click ()
Dim l0508 As Variant
Dim l050C As Variant
Dim l0510 As String
Dim l0512 As Variant
l0508 = extfn1018("AOL FRAME25", 0&)
l050C = extfn10F8(l0508, "Welcome")
l0510$ = String(30, 0)
l0512 = extfnFA8(l050C, l0510$, 250)
If  l0512 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
phish.Show
End Sub

Sub rbust_Click ()
Dim l0516 As Variant
Dim l051A As Variant
Dim l051E As String
Dim l0520 As Variant
l0516 = extfn1018("AOL FRAME25", 0&)
l051A = extfn10F8(l0516, "Welcome")
l051E$ = String(30, 0)
l0520 = extfnFA8(l051A, l051E$, 250)
If  l0520 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
buster.Show
End Sub

Sub ren_Click ()
Dim l0528 As Variant
Dim l052C As Variant
Dim l0530 As String
Dim l0532 As Variant
l0528 = extfn1018("AOL FRAME25", 0&)
l052C = extfn10F8(l0528, "Welcome")
l0530$ = String(30, 0)
l0532 = extfnFA8(l052C, l0530$, 250)
If  l0532 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
reenter.Show
End Sub

Sub rs_Click ()
mac.Show
End Sub

Sub scrambl_Click ()
Dim l053E As Variant
Dim l0542 As Variant
Dim l0546 As String
Dim l0548 As Variant
l053E = extfn1018("AOL FRAME25", 0&)
l0542 = extfn10F8(l053E, "Welcome")
l0546$ = String(30, 0)
l0548 = extfnFA8(l0542, l0546$, 250)
If  l0548 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
scram.Show
End Sub

Sub scroll_Click ()
Dim l0550 As Variant
Dim l0554 As Variant
Dim l0558 As String
Dim l055A As Variant
l0550 = extfn1018("AOL FRAME25", 0&)
l0554 = extfn10F8(l0550, "Welcome")
l0558$ = String(30, 0)
l055A = extfnFA8(l0554, l0558$, 250)
If  l055A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
customkill.Show
End Sub

Sub servinman_Click ()
Dim l0562 As Variant
Dim l0566 As Variant
Dim l056A As String
Dim l056C As Variant
l0562 = extfn1018("AOL FRAME25", 0&)
l0566 = extfn10F8(l0562, "Welcome")
l056A$ = String(30, 0)
l056C = extfnFA8(l0566, l056A$, 250)
If  l056C <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
server.Show
End Sub

Sub snmailfix_Click ()
mailfix.Show
End Sub

Sub sub31B8 ()
End Sub

Sub tosb_Click ()
Dim l0578 As Variant
Dim l057C As Variant
Dim l0580 As String
Dim l0582 As Variant
l0578 = extfn1018("AOL FRAME25", 0&)
l057C = extfn10F8(l0578, "Welcome")
l0580$ = String(30, 0)
l0582 = extfnFA8(l057C, l0580$, 250)
If  l0582 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
tos.Show
End Sub

Sub validity_Click ()
valid.Show
End Sub

Sub voterb_Click ()
Dim l058E As Variant
Dim l0592 As Variant
Dim l0596 As String
Dim l0598 As Variant
l058E = extfn1018("AOL FRAME25", 0&)
l0592 = extfn10F8(l058E, "Welcome")
l0596$ = String(30, 0)
l0598 = extfnFA8(l0592, l0596$, 250)
If  l0598 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
vote.Show
End Sub

Sub warezb_Click ()
Dim l05A0 As Variant
Dim l05A4 As Variant
Dim l05A8 As String
Dim l05AA As Variant
l05A0 = extfn1018("AOL FRAME25", 0&)
l05A4 = extfn10F8(l05A0, "Welcome")
l05A8$ = String(30, 0)
l05AA = extfnFA8(l05A4, l05A8$, 250)
If  l05AA <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
warezre.Show
End Sub
