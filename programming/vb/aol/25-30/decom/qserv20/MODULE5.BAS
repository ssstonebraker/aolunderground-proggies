' Module5
Option Explicit
Const c001E = 513 ' &H201%
Const c0026 = 514 ' &H202%
Const c002E = 1036 ' &H40C%
Const c004E = 1035 ' &H40B%
Const c0056 = 1034 ' &H40A%
Const c0104 = 12 ' &HC%
Const c011C = 513 ' &H201%
Const c0122 = 514 ' &H202%
Const c0186 = 1031 ' &H407%
Const c01A6 = 16 ' &H10%
Const c01B0 = 2 ' &H2%

Sub sub9570 ()
Dim l000E As Integer
Dim l0010 As String
Dim l0014 As Integer
Dim l0016 As Variant
Dim l001A As Integer
Dim l001C As Integer
' Const c001E = 513 ' &H201%
Dim l0022 As Variant
' Const c0026 = 514 ' &H202%
Dim l002A As Variant
' Const c002E = 1036 ' &H40C%
Dim l0032 As Variant
Dim l0036 As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As Variant
Dim l0046 As Integer
Dim l0048 As Integer
Dim l004C As Integer
' Const c004E = 1035 ' &H40B%
Dim l0050 As Variant
Dim l0054 As String
' Const c0056 = 1034 ' &H40A%
Dim l005C As Variant
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting FlashMail)"
If  extfn250(extfn288("AOL Frame25", "America  Online"), "Incoming FlashMail") = False Then
l000E% = extfn288("AOL Frame25", "America  Online")
l0010$ = "Read &Incoming FlashMail"
l0014% = fn9C0("&Mail", l0010$)
Do
DoEvents
l0016 = extfn250(extfn288("AOL Frame25", "America  Online"), "Incoming FlashMail")
l000E% = extfn288("AOL Frame25", "America  Online")
l001A% = extfn288("#32770", 0&)
l001C% = extfn250(l001A%, "OK")
If  l001C% Then
    l0022 = extfn6E8(l001C%, c001E, 0&, 0&)
    l0022 = extfn6E8(l001C%, c0026, 0&, 0&)
    Exit Sub
End If
Loop Until l0016
l002A = extfn218(l0016, "_AOL_Tree")
Do
DoEvents
l002A = extfn218(l0016, "_AOL_Tree")
l0032 = extfn6B0(l002A, c002E, 0, 0&)
subEC8 (.1#)
l002A = extfn218(l0016, "_AOL_Tree")
l0036 = extfn6B0(l002A, c002E, 0, 0&)
l002A = extfn218(l0016, "_AOL_Tree")
subEC8 (.1#)
l002A = extfn218(l0016, "_AOL_Tree")
l003A = extfn6B0(l002A, c002E, 0, 0&)
Loop Until l0032 = l0036 And l003A = l0032 And l0036 = l003A
Else
l0016 = extfn250(extfn288("AOL Frame25", "America  Online"), "Incoming FlashMail")
    l002A = extfn218(l0016, "_AOL_Tree")
    l0032 = extfn6B0(l002A, c002E, 0, 0&)
End If
l0016 = extfn250(extfn288("AOL Frame25", "America  Online"), "Incoming FlashMail")
l003E = l0032
server.List10.AddItem "Incoming FlashMail:" & server.List3.ListCount
For l0042 = 0 To l0032 - 1
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting FlashMail) [" & l0042 + "]"
l0046% = extfn288("AOL Frame25", 0&)
l0048% = extfn250(extfn288("AOL Frame25", "America  Online"), "Incoming FlashMail")
l004C% = extfn1D88(l0048%, "_AOL_Tree")
l0050 = extfn6E8(l004C%, c004E, l0042, 0)
l0054$ = String$(l0050, 0)
l0050 = extfn720(l004C%, c0056, l0042, l0054$)
If  server.List3.ListCount < 1000 Then server.List3.AddItem l0054$
If  server.List3.ListCount > 1000 Then server.List11.AddItem l0054$
Next l0042
l005C = extfn6078(l0016, 0)
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL 3.0"
server.rede = "go"
End Sub

Sub sub95A8 ()
Dim l0060 As Integer
Dim l0062 As String
Dim l0064 As Integer
Dim l0066 As Variant
Dim l006A As Integer
Dim l006C As Integer
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Integer
Dim l008C As Integer
Dim l008E As Integer
Dim l0090 As Variant
Dim l0094 As String
Dim l0096 As Variant
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting New Mail)"
If  extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail") = False Then
l0060% = extfn288("AOL Frame25", "America  Online")
l0062$ = "Read &New Mail" + Chr$(9) + "Ctrl+R"
l0064% = fn9C0("&Mail", l0062$)
Do
DoEvents
l0066 = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
l0060% = extfn288("AOL Frame25", "America  Online")
l006A% = extfn288("#32770", 0&)
l006C% = extfn250(l006A%, "OK")
If  l006C% Then
    l006E = extfn6E8(l006C%, c001E, 0&, 0&)
    l006E = extfn6E8(l006C%, c0026, 0&, 0&)
    Exit Sub
End If
Loop Until l0066
l0072 = extfn218(l0066, "_AOL_Tree")
Do
DoEvents
l0072 = extfn218(l0066, "_AOL_Tree")
l0076 = extfn6B0(l0072, c002E, 0, 0&)
subEC8 (1)
l0072 = extfn218(l0066, "_AOL_Tree")
l007A = extfn6B0(l0072, c002E, 0, 0&)
l0072 = extfn218(l0066, "_AOL_Tree")
subEC8 (1)
l0072 = extfn218(l0066, "_AOL_Tree")
l007E = extfn6B0(l0072, c002E, 0, 0&)
Loop Until l0076 = l007A And l007E = l0076 And l007A = l007E
Else
    l0066 = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
    l0072 = extfn218(l0066, "_AOL_Tree")
    l0076 = extfn6B0(l0072, c002E, 0, 0&)
End If
    l0066 = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
l0082 = l0076
server.List10.AddItem "New Mail:" & server.List3.ListCount
For l0086 = 0 To l0076 - 1
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting New Mail) [" & l0086 + "]"
l008A% = extfn288("AOL Frame25", 0&)
l008C% = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
l008E% = extfn1D88(l008C%, "_AOL_Tree")
l0090 = extfn6E8(l008E%, c004E, l0086, 0)
l0094$ = String$(l0090, 0)
l0090 = extfn720(l008E%, c0056, l0086, l0094$)
server.List3.AddItem l0094$
Next l0086
l0096 = extfn6078(l0066, 0)
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL 3.0"
server.rede = "go"
End Sub

Sub sub95E0 ()
Dim l009A As Integer
Dim l009C As String
Dim l009E As Integer
Dim l00A0 As Variant
Dim l00A4 As Integer
Dim l00A6 As Integer
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Integer
Dim l00C6 As Integer
Dim l00C8 As Integer
Dim l00CA As Variant
Dim l00CE As String
Dim l00D0 As Variant
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting Old Mail)"
If  extfn250(extfn288("AOL Frame25", "America  Online"), "Old Mail") = False Then
l009A% = extfn288("AOL Frame25", "America  Online")
l009C$ = "Check Mail You've &Read"
l009E% = fn9C0("&Mail", l009C$)
Do
DoEvents
l00A0 = extfn250(extfn288("AOL Frame25", "America  Online"), "Old Mail")
l009A% = extfn288("AOL Frame25", "America  Online")
l00A4% = extfn288("#32770", 0&)
l00A6% = extfn250(l00A4%, "OK")
If  l00A6% Then
    l00A8 = extfn6E8(l00A6%, c001E, 0&, 0&)
    l00A8 = extfn6E8(l00A6%, c0026, 0&, 0&)
    Exit Sub
End If
Loop Until l00A0
l00AC = extfn218(l00A0, "_AOL_Tree")
Do
DoEvents
l00AC = extfn218(l00A0, "_AOL_Tree")
l00B0 = extfn6B0(l00AC, c002E, 0, 0&)
subEC8 (1)
l00AC = extfn218(l00A0, "_AOL_Tree")
l00B4 = extfn6B0(l00AC, c002E, 0, 0&)
l00AC = extfn218(l00A0, "_AOL_Tree")
subEC8 (1)
l00AC = extfn218(l00A0, "_AOL_Tree")
l00B8 = extfn6B0(l00AC, c002E, 0, 0&)
Loop Until l00B0 = l00B4 And l00B8 = l00B0 And l00B4 = l00B8
Else
l00A0 = extfn250(extfn288("AOL Frame25", "America  Online"), "Old Mail")
    l00AC = extfn218(l00A0, "_AOL_Tree")
    l00B0 = extfn6B0(l00AC, c002E, 0, 0&)
End If
l00A0 = extfn250(extfn288("AOL Frame25", "America  Online"), "Old Mail")
l00BC = l00B0
server.List10.AddItem "Old Mail:" & server.List3.ListCount
For l00C0 = 0 To l00B0 - 1
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting Old Mail) [" & l00C0 + "]"
l00C4% = extfn288("AOL Frame25", 0&)
l00C6% = extfn250(extfn288("AOL Frame25", "America  Online"), "Old Mail")
l00C8% = extfn1D88(l00C6%, "_AOL_Tree")
l00CA = extfn6E8(l00C8%, c004E, l00C0, 0)
l00CE$ = String$(l00CA, 0)
l00CA = extfn720(l00C8%, c0056, l00C0, l00CE$)
server.List3.AddItem l00CE$
Next l00C0
l00D0 = extfn6078(l00A0, 0)
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL 3.0"
server.rede = "go"
End Sub

Sub sub9618 ()
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
If  server.Command1.Caption = "Stop" Then
On Error Resume Next
Kill "c:\mail1.dat"
Kill "c:\mail2.dat"
Kill "c:\mail3.dat"
Kill "c:\mail4.dat"
Open l00D4 + "c:\mail1.dat" For Output As 1
Open l00D4 + "c:\mail2.dat" For Output As 2
Open l00D4 + "c:\mail3.dat" For Output As 3
Open l00D4 + "c:\mail4.dat" For Output As 4
For l00D8 = 0 To server.List3.ListCount - 1
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Prepping Mail) [" & l00D8 + "]"
l00DC = InStr(server.List3.List(l00D8), Chr(9))
l00E0 = InStr(l00DC + 1, server.List3.List(l00D8), Chr(9), 1)
l00E4 = Mid(server.List3.List(l00D8), l00E0 + 1, Len(server.List3.List(l00D8)))
If  l00D8 < 550 Then Print #1, (l00D8 + 1) & ") " + l00E4
If  l00D8 > 549 And l00D8 < 1100 Then Print #2, (l00D8 + 1) & ") " + l00E4
If  l00D8 > 1099 And l00D8 < 1650 Then Print #3, (l00D8 + 1) & ") " + l00E4
If  l00D8 > 1649 And l00D8 < 2200 Then Print #4, (l00D8 + 1) & ") " + l00E4
Next l00D8
Close 1
Close 2
Close 3
Close 4
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL3.0"
End If
End Sub

Sub sub9650 ()
Dim l00E8 As String
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
' Const c0104 = 12 ' &HC%
Dim l0106 As Variant
Dim l010A As Variant
Dim l0110 As Variant
Dim l0114 As Variant
Dim l0118 As Variant
' Const c011C = 513 ' &H201%
Dim l011E As Variant
' Const c0122 = 514 ' &H202%
Dim l0124 As Variant
Dim l0128 As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Integer
Dim l013C As Variant
Dim l0140 As Variant
Dim l0144 As Variant
server.rede3 = "stop"
On Error Resume Next
If  server.List6.ListCount = 0 Then Exit Sub
For l00EA = 0 To server.List6.ListCount - 1
l00E8 = l00E8 + server.List6.List(l00EA) + ", "
Next l00EA
l00EE = Mid(l00E8, 1, Len(l00E8) - 1)
Open l00F2 + "c:\mail4.dat" For Input As 1
l00F2 = extfn8688("AOL FRAME25", 0&)
l00F8 = extfn250(l00F2, "$Four")
l00FC = extfn218(l00F8, "RICHCNTL")
If  l00FC = 0 Then l00FC = extfn218(l00F8, "_AOL_VIEW")
l0100 = Input(LOF(1), 1)
Close 1
l00F2 = extfn8688("AOL FRAME25", 0&)
l00F8 = extfn250(l00F2, "$Four")
l00FC = extfn218(l00F8, "RICHCNTL")
If  l00FC = 0 Then l00FC = extfn218(l00F8, "_AOL_VIEW")
l0106 = extfn720(l00FC, c0104, 0, l0100)
L9492:
l010A = extfn218(l00F8, "_AOL_EDIT")
l0106 = extfn720(l010A, c0104, 0, l00E8)
l0110 = extfn87D8(l010A, 2)
l0110 = extfn87D8(l0110, 2)
l0110 = extfn87D8(l0110, 2)
l0110 = extfn87D8(l0110, 2)
l0114 = "(Mails 1101 - " & (server.List11.ListCount + server.List3.ListCount) + ")"
l0106 = extfn720(l0110, c0104, 0, "-=‡» ¡Queso! \\'arez Server v2.0 List " + l0114 + ") «‡=-")
DoEvents
l0118 = extfn218(l00F8, "_AOL_ICON")
l011E = extfn6E8(l0118, c011C, 0, 0)
l011E = extfn6E8(l0118, c0122, 0, 0)
Do
l00F2 = extfn8688("AOL FRAME25", 0&)
l00F8 = extfn250(l00F2, "$Four")
l0124 = extfn250(l00F2, "Error")
DoEvents
Loop Until l00F8 = 0 Or l0124
If  l0124 Then
   l0128 = extfn218(l0124, "_AOL_VIEW")
   l012E = extfn138(l0128)
l0132 = extfn250(l0124, "OK")
subFE0 (l0132)
DoEvents
l0136 = server.List6.ListCount - 1
l013A = -1
l013C = ""
Do
l013A = l013A + 1
l00F2 = InStr(l012E, server.List6.List(l013A))
If  l00F2 Then
l013C = l013C + server.List6.List(l013A) + ", "
DoEvents
server.List6.RemoveItem l013A
l013A = -1
End If
Loop Until l013A = server.List6.ListCount - 1
l00E8 = ""
For l0140 = 0 To server.List6.ListCount - 1
l00E8 = l00E8 + server.List6.List(l0140) + ", "
Next l0140
sub1C28 "-=‡» " + l013C + "Your MailBox Is FULL «‡=-"
DoEvents
If  Len(Trim(l00E8)) > 0 Then GoTo L9492
GoTo L9B5E:
End If
server.List2.AddItem "1"
If  Len(Trim(l00E8)) > 0 Then server.List2.AddItem "-=‡» " + l00E8 + " List " + l0114 + " has been sent «‡=-" Else server.List2.AddItem "-=‡» " + l0144 + " List " + l0114 + " has been sent «‡=-"
DoEvents
L9B5E:
DoEvents
server.rede = "go"
Do
DoEvents
Loop Until server.List2.ListCount = 0
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL3.0"
server.rede3 = "go"
End Sub

Sub sub9688 (p0148 As Integer)
Dim l014A As String
Dim l014E As Variant
Dim l0152 As Variant
Dim l0156 As Variant
Dim l015A As Variant
Dim l015E As Variant
Dim l0162 As Variant
Dim l0166 As Integer
Dim l0168 As Variant
Dim l016C As Variant
Dim l0170 As Variant
Dim l0174 As Variant
Dim l0178 As Variant
Dim l017C As Variant
Dim l0182 As Variant
' Const c0186 = 1031 ' &H407%
Dim l0188 As Variant
Dim l018C As Variant
Dim l0190 As Integer
Dim l0192 As String
Dim l0194 As Integer
Dim l0196 As Variant
Dim l019A As Variant
Dim l019E As Variant
Dim l01A2 As Variant
' Const c01A6 = 16 ' &H10%
Dim l01A8 As Variant
Dim l01AC As Variant
' Const c01B0 = 2 ' &H2%
Dim l01B4 As Variant
Dim l01B8 As Variant
Dim l01BC As Variant
Dim l01C0 As Variant
Dim l01C4 As Variant
Dim l01C8 As Variant
Dim l01CC As Variant
Dim l01D0 As Variant
Dim l01D4 As Variant
Dim l01D8 As Variant
Dim l01DC As Variant
Dim l01E0 As Variant
Dim l01E4 As Variant
Dim l01E8 As Variant
Dim l01EC As Variant
Dim l01F0 As Variant
Dim l01F4 As Variant
Dim l01F8 As Variant
Dim l01FC As Variant
Dim l0200 As Variant
Dim l0206 As Variant
Dim l020C As Variant
Dim l0210 As Variant
Dim l0214 As Variant
Dim l0218 As Variant
Dim l021C As Variant
Dim l0220 As Variant
Dim l0224 As Variant
If  server.Command1.Caption = "Start" Then Exit Sub
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Sending #" & (p0148 + 1) & ")"
server.Frame3D3.Caption = "Want #" & (p0148 + 1)
l014A = ""
For l014E = 0 To server.List1.ListCount - 1
l0152 = InStr(server.List1.List(l014E), "-#")
If  l0152 Then
l0156 = Mid(server.List1.List(l014E), l0152 + 2, Len(server.List1.List(l014E)))
If  Val(l0156) - 1 = p0148 Then
l015A = Mid(server.List1.List(l014E), 1, l0152 - 1)
server.List6.AddItem l015A
l014A = l014A + l015A + ", "
server.List1.RemoveItem l014E
l015E = l015E - 1
End If
End If
Next l014E
l0162 = Mid(l014A, 1, Len(l014A) - 1)
If  l014A = "" Then Exit Sub
DoEvents
For l0168 = 0 To server.List10.ListCount - 1
l016C = InStr(server.List10.List(l0168), ":")
l0170 = Mid(server.List10.List(l0168), l016C + 1, Len(server.List10.List(l0168)))
DoEvents
If  p0148 >= l0170 Then
l0166 = l0168
End If
Next l0168
DoEvents
l016C = InStr(server.List10.List(l0166), ":")
l0170 = Mid(server.List10.List(l0166), l016C + 1, Len(server.List10.List(l0166)))
l0174 = Mid(server.List10.List(l0166), 1, l016C - 1)
l0178 = extfn250(extfn288("AOL Frame25", "America  Online"), l0174)
DoEvents
    l017C = extfn218(l0178, "_AOL_Tree")
l0182 = extfn1678(l017C)
DoEvents
l0188 = extfn6E8(l017C, c0186, p0148 - l0170, 0)
DoEvents
l018C = extfn250(l0182, "Read")
DoEvents
Do
If  l0174 = "Incoming FlashMail" Then l018C = extfn250(l0182, "Open")
subFE0 (l018C)
DoEvents
DoEvents
l0190% = extfn288("AOL Frame25", "America  Online")
l0192$ = "Stop Incoming Text" & Chr(9) + "Esc"
l0194% = fn9C0("&File", l0192$)
DoEvents
l0194% = fn9C0("&File", l0192$)
DoEvents
l0194% = fn9C0("&File", l0192$)
DoEvents
l0194% = fn9C0("&File", l0192$)
DoEvents
DoEvents
l0196 = extfn288("AOL FRAME25", 0&)
l019A = extfn250(l0196, "Forward")
Loop Until l019A
l019E = extfn1678(l019A)
Do
l0190% = extfn288("AOL Frame25", "America  Online")
l0192$ = "Get a Member's Profile" + Chr$(9) + "Ctrl+G"
l0194% = fn9C0("Mem&bers", l0192$)
Do
l01A2 = extfn250(l0190%, "Get a Member's Profile")
DoEvents
Loop Until l01A2
l01A8 = extfn6E8(l01A2, c01A6, 0, 0)
DoEvents
l01AC = extfn218(l019E, "_AOL_ICON")
    l01B4 = extfn448(l01AC, c01B0)
    l01B8 = extfn448(l01B4, c01B0)
subFE0 (l01B8)
subEC8 (.4#)
l01BC = extfn250(l0196, "Send Now")
DoEvents
Loop Until l01BC
l01C0 = extfn6E8(l019E, c01A6, 0, 0)
Do
l01A2 = extfn288("AOL FRAME25", 0&)
l01A8 = extfn250(l01A2, "Send Now")
DoEvents
Loop Until l01A8

LA6C8:
l01C4 = extfn1678(l01A8)
l01C8 = extfn218(l01C4, "_AOL_EDIT")
l01CC = extfn720(l01C8, c0104, 0, l014A)
l01D0 = extfn250(l01C4, "SUBJECT:")
l01D4 = extfn720(l01C4, c0104, 0, "¡Queso!")
Do
l01D8 = extfn288("AOL FRAME25", 0&)
l01DC = extfn250(l01D8, "¡Queso!")
DoEvents
Loop Until l01A8
l01D0 = extfn87D8(l01D0, 2)
l01E0 = extfn138(l01D0)
l01E4 = Mid(l01E0, 6, Len(l01E0))
l01C8 = extfn720(l01D0, c0104, 0, l01E4)
l01E8 = extfn218(l01C4, "RICHCNTL")
If  l01E8 = 0 Then l01E8 = extfn218(l01A8, "_AOL_VIEW")
If  Len(Trim(server.Text1)) = 0 Then
Else
l01C8 = extfn720(l01E8, c0104, 0, server.Text1 & Chr(13))
End If
l01C8 = extfn720(l01E8, c0104, 0, "-=‡» ¡Queso! \\'arez Server v2.0 by LeidenKranz «‡=-" & Chr(13) & Chr(13))
DoEvents
l01EC = extfn218(l01C4, "_AOL_ICON")
l01F0 = extfn6E8(l01EC, c011C, 0, 0)
l01F0 = extfn6E8(l01EC, c0122, 0, 0)
l0190% = extfn288("AOL Frame25", "America  Online")
l0192$ = "Get a Member's Profile" + Chr$(9) + "Ctrl+G"
l0194% = fn9C0("Mem&bers", l0192$)
Do
l01A2 = extfn250(l0190%, "Get a Member's Profile")
DoEvents
Loop Until l01A2
l01A8 = extfn6E8(l01A2, c01A6, 0, 0)
l01A8 = extfn6E8(l01C4, c01A6, 0, 0)
Do
l01A2 = extfn288("AOL FRAME25", 0&)
l01A8 = extfn250(l01A2, "¡Queso!")
l01F4 = extfn250(l01A2, "Error")
DoEvents
Loop Until l01A8 = 0 Or l01F4
DoEvents
    l0178 = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
If  server.Check3D1.Value = True Then l01F8 = extfn250(l0178, "Keep As New")
subFE0 (l01F8)
DoEvents
If  l01F4 Then
l01FC = extfn6E8(l01F4, c01A6, 0, 0)
GoTo LB0E6
   l0200 = extfn218(l01F4, "_AOL_VIEW")
   l0206 = extfn17C8(l0200, &HE, 0, 0&)
   l0156 = String$(l0206, 0)
   l020C = extfn1800(l0200, &HD, l0206, l0156)
For l0210 = 0 To server.List6.ListCount - 1
l01A2 = InStr(l0156, server.List6.List(l0210))
If  l01A2 Then
l0214 = l0214 + server.List6.List(l0210) + " ,"
server.List6.RemoveItem l0210
l0210 = l0210 - 1
End If
Next l0210
l014A = ""
For l0218 = 0 To server.List6.ListCount - 1
l014A = l014A + server.List6.List(l0218) + ", "
Next l0218
server.List2.AddItem "1"
server.List2.AddItem "-=‡» " + l0214 + " the Your MailBox Is FULL «‡=-"
DoEvents
If  server.List6.ListCount > 0 Then GoTo LA6C8
End If
LB0E6:
Do
l01A2 = extfn288("AOL FRAME25", 0&)
l01A8 = extfn250(l01A2, "¡Queso!")
DoEvents
Loop Until l01A8 = 0
Do
l01A2 = extfn288("AOL FRAME25", 0&)
l01BC = extfn250(l01A2, "Send Now")
If  l01BC Then
l021C = extfn1678(l01BC)
l0220 = extfn6E8(l021C, c01A6, 0, 0)
End If
Loop Until l01BC = 0
    l0178 = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
If  server.Check3D1.Value = True Then l01F8 = extfn250(l0178, "Keep As New")
subFE0 (l01F8)
DoEvents
server.List2.AddItem "1"
server.List2.AddItem "-=‡» " + l0162 + " #" & (p0148 + 1) & " has been sent «‡=-"
DoEvents
For l0224 = 0 To server.List6.ListCount - 1
server.list9.AddItem server.List6.List(l0224) + " #" & p0148 + 1
Next l0224
server.List6.Clear
server.Frame3D3.Caption = "Server Avail."
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL3.0"
server.Frame3D6.Caption = "Served " & server.list9.ListCount
server.rede = "go"
Do
DoEvents
Loop Until server.List2.ListCount = 0
End Sub

Sub sub96C0 ()
Dim l0228 As String
Dim l022A As Variant
Dim l022E As Variant
Dim l0232 As Variant
Dim l0236 As Variant
Dim l023A As Variant
Dim l023E As Variant
Dim l0242 As Variant
Dim l0246 As Variant
Dim l024A As Variant
Dim l024E As Variant
Dim l0252 As Variant
Dim l0256 As Variant
Dim l025A As Variant
Dim l025E As Variant
Dim l0262 As Variant
Dim l0266 As Variant
Dim l026A As Variant
Dim l026E As Integer
Dim l0270 As Variant
Dim l0274 As Variant
Dim l0278 As Variant
server.rede3 = "stop"
On Error Resume Next
If  server.List6.ListCount = 0 Then Exit Sub
For l022A = 0 To server.List6.ListCount - 1
l0228 = l0228 + server.List6.List(l022A) + ", "
Next l022A
l022E = Mid(l0228, 1, Len(l0228) - 1)
Open l0232 + "c:\mail1.dat" For Input As 1
l0232 = extfn8688("AOL FRAME25", 0&)
l0236 = extfn250(l0232, "$One")
l023A = extfn218(l0236, "RICHCNTL")
If  l023A = 0 Then l023A = extfn218(l0236, "_AOL_VIEW")
l023E = Input(LOF(1), 1)
Close 1
LB784:
l0242 = extfn218(l0236, "_AOL_EDIT")
l0246 = extfn720(l0242, c0104, 0, l0228)
l024A = extfn87D8(l0242, 2)
l024A = extfn87D8(l024A, 2)
l024A = extfn87D8(l024A, 2)
l024A = extfn87D8(l024A, 2)
If  server.List3.ListCount < 551 Then l024E = "(Mails 1 - " & server.List3.ListCount + ")"
If  server.List3.ListCount > 550 Then l024E = "(Mails 1 - 550)"
l0246 = extfn720(l024A, c0104, 0, "-=‡» ¡Queso! \\'arez Server v2.0 List " + l024E + " «‡=-")
DoEvents
l023A = extfn218(l0236, "RICHCNTL")
l0246 = extfn720(l023A, c0104, 0, l023E)
l0252 = extfn218(l0236, "_AOL_ICON")
l0256 = extfn6E8(l0252, c011C, 0, 0)
l0256 = extfn6E8(l0252, c0122, 0, 0)
Do
l0232 = extfn8688("AOL FRAME25", 0&)
l0236 = extfn250(l0232, "$One")
l025A = extfn250(l0232, "Error")
DoEvents
Loop Until l0236 = 0 Or l025A
If  l025A Then
   l025E = extfn218(l025A, "_AOL_VIEW")
   l0262 = extfn138(l025E)
l0266 = extfn250(l025A, "OK")
subFE0 (l0266)
DoEvents
l026A = server.List6.ListCount - 1
l026E = -1
l0270 = ""
Do
l026E = l026E + 1
l0232 = InStr(l0262, server.List6.List(l026E))
If  l0232 Then
l0270 = l0270 + server.List6.List(l026E) + ", "
DoEvents
server.List6.RemoveItem l026E
l026E = -1
End If
Loop Until l026E = server.List6.ListCount - 1
l0228 = ""
For l0274 = 0 To server.List6.ListCount - 1
l0228 = l0228 + server.List6.List(l0274) + ", "
Next l0274
sub1C28 "-=‡» " + l0270 + "Your MailBox Is FULL «‡=-"
DoEvents
If  Len(Trim(l0228)) > 0 Then GoTo LB784
GoTo LBEFC:
End If
server.List2.AddItem "1"
If  Len(Trim(l0228)) > 0 Then server.List2.AddItem "-=‡» " + l0228 + " List " + l024E + " has been sent «‡=-" Else server.List2.AddItem "-=‡» " + l0278 + " List " + l024E + " has been sent «‡=-"
DoEvents
LBEFC:
DoEvents
server.rede = "go"
Do
DoEvents
Loop Until server.List2.ListCount = 0
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL3.0"
server.rede3 = "go"
End Sub

Sub sub96F8 ()
Dim l027C As String
Dim l027E As Variant
Dim l0282 As Variant
Dim l0286 As Variant
Dim l028A As Variant
Dim l028E As Variant
Dim l0292 As Variant
Dim l0296 As Variant
Dim l029A As Variant
Dim l029E As Variant
Dim l02A2 As Variant
Dim l02A6 As Variant
Dim l02AA As Variant
Dim l02AE As Variant
Dim l02B2 As Variant
Dim l02B6 As Variant
Dim l02BA As Variant
Dim l02BE As Variant
Dim l02C2 As Integer
Dim l02C4 As Variant
Dim l02C8 As Variant
Dim l02CC As Variant
server.rede3 = "stop"
On Error Resume Next
If  server.List6.ListCount = 0 Then Exit Sub
For l027E = 0 To server.List6.ListCount - 1
l027C = l027C + server.List6.List(l027E) + ", "
Next l027E
l0282 = Mid(l027C, 1, Len(l027C) - 1)
Open l0286 + "c:\mail3.dat" For Input As 1
l0286 = extfn8688("AOL FRAME25", 0&)
l028A = extfn250(l0286, "$Three")
l028E = extfn218(l028A, "RICHCNTL")
If  l028E = 0 Then l028E = extfn218(l028A, "_AOL_VIEW")
l0292 = Input(LOF(1), 1)
Close 1
l0286 = extfn8688("AOL FRAME25", 0&)
l028A = extfn250(l0286, "$Three")
l028E = extfn218(l028A, "RICHCNTL")
If  l028E = 0 Then l028E = extfn218(l028A, "_AOL_VIEW")
l0296 = extfn720(l028E, c0104, 0, l0292)
LC2DE:
l029A = extfn218(l028A, "_AOL_EDIT")
l0296 = extfn720(l029A, c0104, 0, l027C)
l029E = extfn87D8(l029A, 2)
l029E = extfn87D8(l029E, 2)
l029E = extfn87D8(l029E, 2)
l029E = extfn87D8(l029E, 2)
If  server.List11.ListCount < 551 Then l02A2 = "(Mails 1101 - " & server.List11.ListCount + ")"
If  server.List11.ListCount > 550 Then l02A2 = "(Mails 1101 - 1650)"
l0296 = extfn720(l029E, c0104, 0, "-=‡» ¡Queso! \\'arez Server v2.0 List " + l02A2 + " «‡=-")
DoEvents
l02A6 = extfn218(l028A, "_AOL_ICON")
l02AA = extfn6E8(l02A6, c011C, 0, 0)
l02AA = extfn6E8(l02A6, c0122, 0, 0)
Do
l0286 = extfn8688("AOL FRAME25", 0&)
l028A = extfn250(l0286, "$Three")
l02AE = extfn250(l0286, "Error")
DoEvents
Loop Until l028A = 0 Or l02AE
If  l02AE Then
   l02B2 = extfn218(l02AE, "_AOL_VIEW")
   l02B6 = extfn138(l02B2)
l02BA = extfn250(l02AE, "OK")
subFE0 (l02BA)
DoEvents
l02BE = server.List6.ListCount - 1
l02C2 = -1
l02C4 = ""
Do
l02C2 = l02C2 + 1
l0286 = InStr(l02B6, server.List6.List(l02C2))
If  l0286 Then
l02C4 = l02C4 + server.List6.List(l02C2) + ", "
DoEvents
server.List6.RemoveItem l02C2
l02C2 = -1
End If
Loop Until l02C2 = server.List6.ListCount - 1
l027C = ""
For l02C8 = 0 To server.List6.ListCount - 1
l027C = l027C + server.List6.List(l02C8) + ", "
Next l02C8
sub1C28 "-=‡» " + l02C4 + "Your MailBox Is FULL «‡=-"
DoEvents
If  Len(Trim(l027C)) > 0 Then GoTo LC2DE
GoTo LC9EE:
End If
server.List2.AddItem "1"
If  Len(Trim(l027C)) > 0 Then server.List2.AddItem "-=‡» " + l027C + " List " + l02A2 + " has been sent «‡=-" Else server.List2.AddItem "-=‡» " + l02CC + " List " + l02A2 + " has been sent «‡=-"
DoEvents
LC9EE:
DoEvents
server.rede = "go"
Do
DoEvents
Loop Until server.List2.ListCount = 0
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL3.0"
server.rede3 = "go"
End Sub

Sub sub9730 ()
Dim l02D0 As String
Dim l02D2 As Variant
Dim l02D6 As Variant
Dim l02DA As Variant
Dim l02DE As Variant
Dim l02E2 As Variant
Dim l02E6 As Variant
Dim l02EA As Variant
Dim l02EE As Variant
Dim l02F2 As Variant
Dim l02F6 As Variant
Dim l02FA As Variant
Dim l02FE As Variant
Dim l0302 As Variant
Dim l0306 As Variant
Dim l030A As Variant
Dim l030E As Variant
Dim l0312 As Variant
Dim l0316 As Variant
Dim l031A As Integer
Dim l031C As Variant
Dim l0320 As Variant
Dim l0324 As Variant
server.rede3 = "stop"
On Error Resume Next
If  server.List6.ListCount = 0 Then MsgBox l02D2: Exit Sub
For l02D6 = 0 To server.List6.ListCount - 1
l02D0 = l02D0 + server.List6.List(l02D6) + ", "
Next l02D6
Open l02DA + "c:\mail2.dat" For Input As 1
l02DA = extfn8688("AOL FRAME25", 0&)
l02DE = extfn250(l02DA, "$Two")
l02E2 = extfn218(l02DE, "RICHCNTL")
If  l02E2 = 0 Then l02E2 = extfn218(l02DE, "_AOL_VIEW")
l02E6 = Input(LOF(1), 1)
Close 1
l02EA = Mid(l02D0, 1, Len(l02D0) - 1)
l02DA = extfn8688("AOL FRAME25", 0&)
l02DE = extfn250(l02DA, "$Two")
l02E2 = extfn218(l02DE, "RICHCNTL")
If  l02E2 = 0 Then l02E2 = extfn218(l02DE, "_AOL_VIEW")
l02EE = extfn720(l02E2, c0104, 0, l02E6)
LCDDA:
l02F2 = extfn218(l02DE, "_AOL_EDIT")
l02EE = extfn720(l02F2, c0104, 0, l02D0)
l02F6 = extfn87D8(l02F2, 2)
l02F6 = extfn87D8(l02F6, 2)
l02F6 = extfn87D8(l02F6, 2)
l02F6 = extfn87D8(l02F6, 2)
If  server.List11.ListCount < 551 Then l02FA = "(Mails 551 - " & server.List3.ListCount + ")"
If  server.List11.ListCount > 550 Then l02FA = "(Mails 551 - 1101)"
l02EE = extfn720(l02F6, c0104, 0, "-=‡» ¡Queso! \\'arez Server v2.0 List " + l02FA + " «‡=-")
DoEvents
l02FE = extfn218(l02DE, "_AOL_ICON")
l0302 = extfn6E8(l02FE, c011C, 0, 0)
l0302 = extfn6E8(l02FE, c0122, 0, 0)
Do
l02DA = extfn8688("AOL FRAME25", 0&)
l02DE = extfn250(l02DA, "$Two")
l0306 = extfn250(l02DA, "Error")
DoEvents
Loop Until l02DE = 0 Or l0306
If  l0306 Then
   l030A = extfn218(l0306, "_AOL_VIEW")
   l030E = extfn138(l030A)
l0312 = extfn250(l0306, "OK")
subFE0 (l0312)
DoEvents
l0316 = server.List6.ListCount - 1
l031A = -1
l031C = ""
Do
l031A = l031A + 1
l02DA = InStr(l030E, server.List6.List(l031A))
If  l02DA Then
l031C = l031C + server.List6.List(l031A) + ", "
DoEvents
server.List6.RemoveItem l031A
l031A = -1
End If
Loop Until l031A = server.List6.ListCount - 1
l02D0 = ""
For l0320 = 0 To server.List6.ListCount - 1
l02D0 = l02D0 + server.List6.List(l0320) + ", "
Next l0320
sub1C28 "-=‡» " + l031C + "Your MailBox Is FULL «‡=-"
DoEvents
If  Len(Trim(l02D0)) > 0 Then GoTo LCDDA
GoTo LD4E6:
End If
server.List2.AddItem "1"
If  Len(Trim(l02D0)) > 0 Then server.List2.AddItem "-=‡» " + l02D0 + " List " + l02FA + " has been sent «‡=-" Else server.List2.AddItem "-=‡» " + l0324 + " List " + l02FA + " has been sent «‡=-"
DoEvents
LD4E6:
DoEvents
server.rede = "go"
Do
DoEvents
Loop Until server.List2.ListCount = 0
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL3.0"
server.rede3 = "go"
End Sub

Sub sub9768 ()
Dim l0328 As Integer
Dim l032A As String
Dim l032C As Integer
Dim l032E As Variant
Dim l0332 As Integer
Dim l0334 As Integer
Dim l0336 As Variant
Dim l033A As Variant
Dim l033E As Variant
Dim l0342 As Variant
Dim l0346 As Variant
Dim l034A As Variant
Dim l034E As Variant
Dim l0352 As Integer
Dim l0354 As Integer
Dim l0356 As Integer
Dim l0358 As Variant
Dim l035C As String
Dim l035E As Variant
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting Sent Mail)"
If  extfn250(extfn288("AOL Frame25", "America  Online"), "Outgoing Mail") = False Then
l0328% = extfn288("AOL Frame25", "America  Online")
l032A$ = "Check Mail You've &Sent"
l032C% = fn9C0("&Mail", l032A$)
Do
DoEvents
l032E = extfn250(extfn288("AOL Frame25", "America  Online"), "Outgoing Mail")
l0328% = extfn288("AOL Frame25", "America  Online")
l0332% = extfn288("#32770", 0&)
l0334% = extfn250(l0332%, "OK")
If  l0334% Then
    l0336 = extfn6E8(l0334%, c001E, 0&, 0&)
    l0336 = extfn6E8(l0334%, c0026, 0&, 0&)
    Exit Sub
End If
Loop Until l032E
l033A = extfn218(l032E, "_AOL_Tree")
Do
DoEvents
l033A = extfn218(l032E, "_AOL_Tree")
l033E = extfn6B0(l033A, c002E, 0, 0&)
subEC8 (1)
l033A = extfn218(l032E, "_AOL_Tree")
l0342 = extfn6B0(l033A, c002E, 0, 0&)
l033A = extfn218(l032E, "_AOL_Tree")
subEC8 (1)
l033A = extfn218(l032E, "_AOL_Tree")
l0346 = extfn6B0(l033A, c002E, 0, 0&)
Loop Until l033E = l0342 And l0346 = l033E And l0342 = l0346
Else
l032E = extfn250(extfn288("AOL Frame25", "America  Online"), "Outgoing Mail")
    l033A = extfn218(l032E, "_AOL_Tree")
    l033E = extfn6B0(l033A, c002E, 0, 0&)
End If
l032E = extfn250(extfn288("AOL Frame25", "America  Online"), "Outgoing Mail")
l034A = l033E
server.List10.AddItem "Outgoing Mail:" & server.List3.ListCount
For l034E = 0 To l033E - 1
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Hold On, Getting Sent Mail) [" & l034E + "]"
l0352% = extfn288("AOL Frame25", 0&)
l0354% = extfn250(extfn288("AOL Frame25", "America  Online"), "Outgoing Mail")
l0356% = extfn1D88(l0354%, "_AOL_Tree")
l0358 = extfn6E8(l0356%, c004E, l034E, 0)
l035C$ = String$(l0358, 0)
l0358 = extfn720(l0356%, c0056, l034E, l035C$)
server.List3.AddItem l035C$
Next l034E
l035E = extfn6078(l032E, 0)
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz for 16bit AOL 3.0"
server.rede = "go"
End Sub

Sub sub97A0 ()
Dim l0362 As Variant
Dim l0366 As Variant
Dim l036A As Variant
Dim l036E As Variant
Dim l0372 As Variant
Dim l0376 As Variant
Dim l037A As Integer
Dim l037C As String
Dim l037E As Integer
If  server.Command1.Caption = "Start" Then Exit Sub
server.Caption = "¡Queso! \\'arez Server v2.0 by LeidenKranz (Sending List)"
For l0362 = 0 To server.List1.ListCount - 1
l0366 = InStr(server.List1.List(l0362), "-List")
If  l0366 Then
l036A = Mid(server.List1.List(l0362), 1, l0366 - 1)
server.List6.AddItem l036A
server.List1.RemoveItem l0362
l0362 = l0362 - 1
End If
Next l0362
DoEvents
server.Frame3D3.Caption = "Want List"
l036E = extfn8688("AOL FRAME25", 0&)
l0372 = extfn250(l036E, "Compose Mail")
If  l0372 Then l0376 = extfn6E8(l0372, c01A6, 0, 0)
DoEvents
l037A% = extfn288("AOL Frame25", "America  Online")
l037C$ = "&Compose Mail" + Chr$(9) + "Ctrl+M"
l037E% = fn9C0("&Mail", l037C$)
Do
l036E = extfn8688("AOL FRAME25", 0&)
l0372 = extfn250(l036E, "Compose Mail")
DoEvents
Loop Until l0372
If  l0372 Then l0376 = extfn720(l0372, c0104, 0, "$One")
DoEvents
Call sub96C0
Do
DoEvents
Loop Until server.rede3 = "go"
If  server.List3.ListCount > 500 Then
l037A% = extfn288("AOL Frame25", "America  Online")
l037C$ = "&Compose Mail" + Chr$(9) + "Ctrl+M"
l037E% = fn9C0("&Mail", l037C$)
Do
l036E = extfn8688("AOL FRAME25", 0&)
l0372 = extfn250(l036E, "Compose Mail")
DoEvents
Loop Until l0372
If  l0372 Then l0376 = extfn720(l0372, c0104, 0, "$Two")
DoEvents
Call sub9730
Do
DoEvents
Loop Until server.rede3 = "go"
End If
If  server.List11.ListCount < 500 And server.List11.ListCount > 0 Then
l037A% = extfn288("AOL Frame25", "America  Online")
l037C$ = "&Compose Mail" + Chr$(9) + "Ctrl+M"
l037E% = fn9C0("&Mail", l037C$)
Do
l036E = extfn8688("AOL FRAME25", 0&)
l0372 = extfn250(l036E, "Compose Mail")
DoEvents
Loop Until l0372
If  l0372 Then l0376 = extfn720(l0372, c0104, 0, "$Three")
DoEvents
Call sub96F8
Do
DoEvents
Loop Until server.rede3 = "go"
End If
If  server.List11.ListCount > 500 Then
l037A% = extfn288("AOL Frame25", "America  Online")
l037C$ = "&Compose Mail" + Chr$(9) + "Ctrl+M"
l037E% = fn9C0("&Mail", l037C$)
Do
l036E = extfn8688("AOL FRAME25", 0&)
l0372 = extfn250(l036E, "Compose Mail")
DoEvents
Loop Until l0372
If  l0372 Then l0376 = extfn720(l0372, c0104, 0, "$Four")
DoEvents
Call sub9650
Do
DoEvents
Loop Until server.rede3 = "go"
End If
server.List6.Clear
server.Frame3D3.Caption = "Server Avail."
End Sub
