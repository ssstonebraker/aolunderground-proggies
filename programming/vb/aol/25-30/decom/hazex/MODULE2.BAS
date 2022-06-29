' Module2
Option Explicit
Const c001E = 12 ' &HC%
Const c0026 = 258 ' &H102%
Const c0036 = 513 ' &H201%
Const c003E = 514 ' &H202%
Const c0042 = 16 ' &H10%
Const c0054 = 2 ' &H2%

Sub sub3128 ()
End Sub

Sub sub3160 ()
Dim l0010 As Integer
Dim l0014 As Variant
Dim l001A As Variant
' Const c001E = 12 ' &HC%
Dim l0022 As Variant
' Const c0026 = 258 ' &H102%
Dim l002A As Variant
Dim l002E As Integer
Dim l0030 As Variant
Dim l0034 As Integer
' Const c0036 = 513 ' &H201%
Dim l003A As Variant
' Const c003E = 514 ' &H202%
Dim l0040 As Integer
' Const c0042 = 16 ' &H10%
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As Variant
' Const c0054 = 2 ' &H2%
Dim l0058 As Variant
Dim l005C As Variant
Dim l0060 As Variant
Dim l0064 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› " & Form1.Room & " Via Guidepager"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Meth. - Mass TOS!"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Phrase - " & Form1.Panel3D5
Call sub3278
l0010% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0014 = extfn24E0(l0010%, "Keyword")
Loop Until l0014 <> 0
l001A = extfn2518(l0014, "_AOL_Edit")
l0022 = extfn1718(l001A, c001E, 0, "GUIDEPAGER")
l002A = extfn16E0(l001A, c0026, 13, 0)
Do
DoEvents
l002E% = extfn1520("AOL Frame25", 0&)
l0030 = extfn24E0(l002E%, "I Need Help!")
Loop Until l0030 > 0
l0034% = extfn2518(l0030, "_AOL_ICON")
l003A = extfn1750(l0034%, c0036, 0, 0)
l003A = extfn1750(l0034%, c003E, 0, 0)
subF38 .1#
l0040% = extfn24E0(l0010%, "I Need Help!")
l0044 = extfn16E0(l0040%, c0042, 0, 0)
Do
DoEvents
l0048 = extfn24E0(l0010%, "Report Password Solicitations")
Loop Until l0048 <> 0
l0022 = extfn16E0(l004C, c0042, 0, 0)
l0050 = extfn2518(l0010%, "_AOL_Edit")
l0058 = extfn14E8(l0050, c0054)
l005C = extfn14E8(l0058, c0054)
l0060 = extfn14E8(l005C, c0054)
l0064 = extfn14E8(l0060, c0054)
l0068 = extfn14E8(l0064, c0054)
l006C = extfn14E8(l0068, c0054)
l0070 = extfn14E8(l006C, c0054)
l0074 = extfn14E8(l0070, c0054)
l0078 = extfn1718(l006C, c001E, 0, Form1.Room)
l0078 = extfn1718(l0074, c001E, 0, Form1.Room + ":" & " " & "" & Form1.Text1)
subF38 .3#
l002A = extfn16E0(l006C, c0026, 13, 0)
subF38 .5#
subFE0
If  Form1.Combo1.ListCount = 0 Then
Form1.Combo1.AddItem l007C
Exit Sub
End If
End Sub

Sub sub3198 ()
Dim l0080 As Integer
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Integer
Dim l0094 As Variant
Dim l0098 As Integer
Dim l009A As Variant
Dim l009E As Integer
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› " & Form1.Room & " Via KO HELP"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Meth. - Mass TOS!"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Phrase - " & Form1.Panel3D5
Call sub3278
l0080% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0082 = extfn24E0(l0080%, "Keyword")
Loop Until l0082 <> 0
l0086 = extfn2518(l0082, "_AOL_Edit")
l008A = extfn1718(l0086, c001E, 0, "KO HELP")
l008E = extfn16E0(l0086, c0026, 13, 0)
Do
DoEvents
l0092% = extfn1520("AOL Frame25", 0&)
l0094 = extfn24E0(l0092%, "I Need Help!")
Loop Until l0094 > 0
l0098% = extfn2518(l0094, "_AOL_ICON")
l009A = extfn1750(l0098%, c0036, 0, 0)
l009A = extfn1750(l0098%, c003E, 0, 0)
subF38 .1#
l009E% = extfn24E0(l0080%, "I Need Help!")
l00A0 = extfn16E0(l009E%, c0042, 0, 0)
Do
DoEvents
l00A4 = extfn24E0(l0080%, "Report Password Solicitations")
Loop Until l00A4 <> 0
l008A = extfn16E0(l00A8, c0042, 0, 0)
l00AC = extfn2518(l0080%, "_AOL_Edit")
l00B0 = extfn14E8(l00AC, c0054)
l00B4 = extfn14E8(l00B0, c0054)
l00B8 = extfn14E8(l00B4, c0054)
l00BC = extfn14E8(l00B8, c0054)
l00C0 = extfn14E8(l00BC, c0054)
l00C4 = extfn14E8(l00C0, c0054)
l00C8 = extfn14E8(l00C4, c0054)
l00CC = extfn14E8(l00C8, c0054)
l00D0 = extfn1718(l00C4, c001E, 0, Form1.Room)
l00D0 = extfn1718(l00CC, c001E, 0, Form1.Room + ":" & " " & "" & Form1.Text1)
subF38 .3#
l008E = extfn16E0(l00C4, c0026, 13, 0)
subF38 .5#
subFE0
If  Form1.Combo1.ListCount = 0 Then
Form1.Combo1.AddItem l00D4
Exit Sub
End If
End Sub

Sub sub31D0 ()
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As String
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Integer
Dim l00F0 As Variant
Dim l00F4 As Variant
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
Dim l0104 As Variant
Dim l010A As Variant
Dim l010E As Variant
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126 As Variant
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
Randomize
For l00D8 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l00DC = 6
    l00E0 = Int(Rnd * l00DC + 1)
      If  l00E0 = 1 Then l00E4$ = "17"
      If  l00E0 = 2 Then l00E4$ = "2"
      If  l00E0 = 3 Then l00E4$ = "72"
      If  l00E0 = 4 Then l00E4$ = "53"
      If  l00E0 = 5 Then l00E4$ = "45"
      If  l00E0 = 6 Then l00E4$ = "11"
      If  l00E0 = 7 Then l00E4$ = "65"
      Next
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› " & l00E6 & " Via IM Violation"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Meth. - Mass TOS!"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Phrase - " & l00EA
Call sub3278
Call subD08("Keyword...", "&Go To")
l00EE% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l00F0 = extfn24E0(l00EE%, "Keyword")
Loop Until l00F0 <> 0
l00F4 = extfn2518(l00F0, "_AOL_Edit")
l00F8 = extfn1718(l00F4, c001E, 0, "Guidepager")
l00FC = extfn16E0(l00F4, c0026, 13, 0)
Do
DoEvents
l0100 = extfn24E0(l00EE%, "I Need Help!")
Loop Until l0100 <> 0
l0104 = extfn2518(l0100, "_AOL_Icon")
l010A = extfn13D0(l0104, 2)
l010E = extfn13D0(l010A, 2)
l0112 = extfn13D0(l010E, 2)
l0116 = extfn13D0(l0112, 2)
l011A = extfn1750(l010E, c0036, 0, 0&)
l011E = extfn1750(l010E, c003E, 0, 0&)
subF38 .1#
l0124% = extfn24E0(l0122%, "I Need Help!")
Do
DoEvents
l0126 = extfn24E0(l00EE%, "Report A Violation")
Loop Until l0126 <> 0
l012A = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l012E = extfn16E0(l0124%, c0042, 0, 0)
l0132 = extfn16E0(10634, c0042, 0, 0)
Do
DoEvents
l0136 = extfn24E0(l00EE%, "Violations via instant Messages")
Loop Until l0136 <> 0
subF38 .1#
l0124% = extfn24E0(l0122%, "Report A Violation")
l012E = extfn16E0(l0124%, c0042, 0, 0)
l00F4 = extfn2518(l0136, "_AOL_Edit")
l013A = extfn13D0(l00F4, 2)
l013E = extfn13D0(l013A, 2)
l0142 = extfn13D0(l013E, 2)
l0146 = extfn13D0(l0142, 2)
l014A = extfn13D0(l0146, 2)
l014E = extfn13D0(l014A, 2)
l0152 = extfn13D0(l014E, 2)
l0156 = extfn13D0(l0152, 2)
l00F8 = extfn1718(l00F4, c001E, 0, Date)
l00F8 = extfn1718(l013E, c001E, 0, Time)
l00F8 = extfn1718(l0146, c001E, 0, "Town Square")
l00F8 = extfn1718(l014E, c001E, 0, "Lobby " & l00E4$)
l00F8 = extfn1718(l0156, c001E, 0, Form1.Room + ":" & " " + Form1.Text1)
subF38 .3#
l015A = extfn2518(l0136, "_AOL_Button")
l011A = extfn1750(l015A, c0036, 0, 0&)
l011E = extfn1750(l015A, c003E, 0, 0&)
l015E = extfn16E0(l0136, c0042, 0, 0)
l015E = extfn16E0(l0100, c0042, 0, 0)
l015E = extfn16E0(l0126, c0042, 0, 0)
If  Form1.Combo1.ListCount = 0 Then
Form1.Combo1.AddItem l00E6
Exit Sub
End If
End Sub

Sub sub3208 ()
Dim l0162 As Variant
Dim l0166 As Variant
Dim l016A As Variant
Dim l016E As String
Randomize
For l0162 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0166 = 7
    l016A = Int(Rnd * l0166 + 1)
      If  l016A = 1 Then l016E$ = "Damn wh0re."
      If  l016A = 2 Then l016E$ = "Lame bitch."
      If  l016A = 3 Then l016E$ = "CrAcK head."
      If  l016A = 4 Then l016E$ = "Sleep bitch."
      If  l016A = 5 Then l016E$ = "Dirty hoe."
      If  l016A = 6 Then l016E$ = "Lamer den BeaSte."
      If  l016A = 7 Then l016E$ = "Gawd damn lamer."
      Next
sub30B0 "•-^v•HaZe-X²·³ [gold] • " & Form1.Room & " " & l016E$
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSed at - " & Time & "-" & Date
End Sub

Sub sub3240 ()
End Sub

Sub sub30B0 (p0170 As Variant)
Dim l0174 As Integer
Dim l0176 As Integer
Dim l017A As Integer
Dim l017C As Integer
Dim l017E As Integer
Dim l0180 As Variant
l0174% = extfn1520("AOL FRAME25", 0&)
l0176% = extfn2518(l0174%, "_AOL_Listbox")
If  l0176% = 0 Then
Exit Sub
End If
l017A% = extfn1600(l0176%)
l017C% = extfn2518(l017A%, "_AOL_EDIT")
l017E% = extfn24E0(l017A%, "Send")
l0180 = extfn1718(l017C%, c001E, 0, p0170)
l0180 = extfn1750(l017C%, c0026, 13, 0)
subF38 1E-300#
End Sub

Sub sub3278 ()
Dim l0184 As Variant
Dim l0188 As Variant
Dim l018C As Variant
Dim l0190 As String
Randomize
For l0184 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0188 = 33
    l018C = Int(Rnd * l0188 + 1)
      If  l018C = 1 Then l0190$ = "Flip side over bitch."
      If  l018C = 2 Then l0190$ = "One less laymuh."
      If  l018C = 3 Then l0190$ = "Bow down ta this."
      If  l018C = 4 Then l0190$ = "You are owned."
      If  l018C = 5 Then l0190$ = "Ignorant bitch."
      If  l018C = 6 Then l0190$ = "Call me Daddy wh0re."
      If  l018C = 7 Then l0190$ = "Hush lil fuck."
      If  l018C = 8 Then l0190$ = "I own jo0."
      If  l018C = 9 Then l0190$ = "Who'z y0 dadyyyy."
      If  l018C = 10 Then l0190$ = "Ass CrAcK licker."
      If  l018C = 11 Then l0190$ = "Jo0 like my style'z?"
      If  l018C = 12 Then l0190$ = "FaTe own'z jo0."
      If  l018C = 13 Then l0190$ = "CHoPPeR iz jo0r g0d."
      If  l018C = 14 Then l0190$ = "KaOS iz my MASTER."
      If  l018C = 15 Then l0190$ = "I b0w b4 Teri."
      If  l018C = 16 Then l0190$ = "Suck my t0es bitch."
      If  l018C = 17 Then l0190$ = "Can you CouNT to 3?"
      If  l018C = 18 Then l0190$ = "Da-SlaYeR own'z me."
      If  l018C = 19 Then l0190$ = "come on you wanna layme."
      If  l018C = 20 Then l0190$ = "Lo0se ass queer."
      If  l018C = 21 Then l0190$ = "B0w b4 DUKE3D."
      If  l018C = 22 Then l0190$ = "Scrambled or TOSed?"
      If  l018C = 23 Then l0190$ = "Can you beleeve this."
      If  l018C = 24 Then l0190$ = "WAR loves you <3."
      If  l018C = 25 Then l0190$ = "Wanna play with KeN?"
      If  l018C = 26 Then l0190$ = "Zer0Puberty iz gay."
      If  l018C = 27 Then l0190$ = "PiMP iz y0 dadyyy."
      If  l018C = 28 Then l0190$ = "SHaDe me baby."
      If  l018C = 29 Then l0190$ = "CrAzY wh0re."
      If  l018C = 30 Then l0190$ = "I like DS nutt'z."
      If  l018C = 31 Then l0190$ = "AOL iz gonna trace you."
      If  l018C = 32 Then l0190$ = "Have a nice day :-)"
      If  l018C = 33 Then l0190$ = "Go away Hacker!!!"
      Next
sub30B0 "•-^v•HaZe-X²·³ [gold] • " & l0190$
End Sub

Sub sub32B0 ()
Dim l0190 As String
Exit Sub
End Sub
