' Module3
Option Explicit
Const c002E = 12 ' &HC%
Const c00B6 = 513 ' &H201%
Const c00BE = 514 ' &H202%
Const c00CC = 13 ' &HD%
Const c00EE = 2 ' &H2%
Const c013C = 16 ' &H10%

Sub sub3328 (p0010 As String, p0012 As String)
Dim l0016 As Integer
Dim l001A As Integer
Dim l001E As Integer
Dim l0020 As Integer
l0016% = extfn1520("AOL Frame25", 0&)
l001A% = extfn2518(l0016%, "_AOL_LISTBOX")
l001E% = extfn1600(l001A%)
l0020% = extfn2518(l001E%, "_AOL_VIEW")
If  l0020% = 0 Then Exit Sub
sub30B0 (p0010)
subF38 (.001#)
sub30B0 (p0012)
subF38 (.001#)
End Sub

Sub sub3360 ()
End Sub

Sub sub3398 (ByVal p0022 As String)
Dim l0024 As Integer
Dim l0026 As Integer
Dim l0028 As Integer
Dim l002A As Integer
Dim l002C As String
' Const c002E = 12 ' &HC%
Dim l0032 As Integer
l0024% = extfn1520("AOL Frame25", 0&)
l0026% = extfn2518(l0024%, "_AOL_LISTBOX")
l0028% = extfn1600(l0026%)
l002A% = extfn2518(l0028%, "_AOL_VIEW")
l002C$ = Chr(13) + Chr(10) + Chr(13) + Chr(10) + "†" + Chr(9) + p0022 + Chr(13) + Chr(10) + Chr(13) + Chr(10)
l0032% = extfn1718(l002A%, c002E, 0, l002C$)
DoEvents
End Sub

Sub sub33D0 (p0036 As TextBox)
sub30B0 ("†    ïèñEñLñYñSñiñUñMñèïñóñv2.0ñóñï ƒÒ¸ﬂÔß")
sub30B0 ("†   †        ïNetherworldññóïv2.0ïñóñTOSserï")
sub30B0 ("†   ïèïóóóóóóóï In Effect ïóóóóóóïèï")
subF38 (2.5#)
sub30B0 ("†             ïóª Now TOSsing:†" + p0036 + " ´óï")
sub30B0 ("†       ïñª " + p0036 + ", welcome to the underworld! ´ñï")
End Sub

Sub sub3408 (p0038 As String, p003A As String, p003C As String, p003E As String)
Dim l0048 As Integer
Dim l004A As Integer
Dim l004C As Integer
Dim l004E As Integer
Dim l0050 As Integer
Dim l0052 As Integer
Dim l0054 As Integer
Dim l0056 As Integer
Dim l0058 As Integer
Dim l005A As Integer
Dim l005C As Integer
Dim l005E As Integer
Dim l0060 As Integer
Dim l0062 As Integer
Dim l0064 As Integer
Dim l0066 As Integer
Dim l0068 As Integer
Dim l006A As Integer
l0048% = fn1018("Compose Mail", "AOL Child")
l004A% = extfn14E8(l0048%, 5)
l004C% = extfn14E8(l004A%, 2)
l004E% = extfn14E8(l004C%, 2)
l0050% = extfn14E8(l004E%, 2)
l0052% = extfn14E8(l0050%, 2)
l0054% = extfn14E8(l0052%, 2)
l0056% = extfn14E8(l0054%, 2)
l0058% = extfn14E8(l0056%, 2)
l005A% = extfn14E8(l0058%, 2)
l005C% = extfn14E8(l005A%, 2)
l005E% = extfn14E8(l005C%, 2)
l0060% = extfn14E8(l005E%, 2)
l0062% = extfn14E8(l0060%, 2)
l0064% = extfn14E8(l0062%, 2)
l0066% = extfn14E8(l0064%, 2)
l0068% = extfn14E8(l0066%, 2)
l006A% = extfn14E8(l0068%, 2)
subF00 l005C%, CStr(p0038$)
subF00 l0060%, CStr(p003A$)
subF00 l0064%, CStr(p003C$)
subF00 l006A%, CStr(p003E$)
End Sub

Sub sub3440 (p0072 As TextBox, p0076 As TextBox)
Dim l0078 As Integer
Dim l007A As Integer
Dim l007C As Integer
Dim l007E As Integer
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Integer
Dim l0086 As Integer
Dim l0088 As Integer
Dim l008A As Integer
Dim l008C As Integer
Dim l008E As Integer
Dim l0090 As Integer
Dim l0092 As Integer
Dim l0094 As Integer
Dim l0096 As Integer
Dim l0098 As Integer
Dim l009A As Integer
Dim l009C As Integer
Dim l009E As Integer
Dim l00A0 As String
Dim l00A2 As Integer
Dim l00A4 As String
Dim l00A6 As String
Dim l00A8 As String
Dim l00AA As Integer
Dim l00AC As Variant
Dim l00B2 As Variant
' Const c00B6 = 513 ' &H201%
Dim l00BA As Variant
' Const c00BE = 514 ' &H202%
Dim l00C0 As Variant
Dim l00C4 As Integer
Dim l00C6 As Integer
Dim l00C8 As Integer
Dim l00CA As String
' Const c00CC = 13 ' &HD%
Dim l00CE As Integer
l0078% = extfn1520("AOL Frame25", 0&)
If  l0078% = 0 Then
    Beep
    MsgBox "You must have AOL 2.5 running before you use this feature."
    Exit Sub
End If
Call subD08("Compose Mail", "&Mail")
l0078% = extfn1520("AOL Frame25", 0&)
l007A% = fn1018("Compose Mail", "AOL Child")
l007C% = extfn14E8(l007A%, 5)
l007E% = extfn14E8(l007C%, 2)
l0080% = extfn14E8(l007E%, 2)
l0082% = extfn14E8(l0080%, 2)
l0084% = extfn14E8(l0082%, 2)
l0086% = extfn14E8(l0084%, 2)
l0088% = extfn14E8(l0086%, 2)
l008A% = extfn14E8(l0088%, 2)
l008C% = extfn14E8(l008A%, 2)
l008E% = extfn14E8(l008C%, 2)
l0090% = extfn14E8(l008E%, 2)
l0092% = extfn14E8(l0090%, 2)
l0094% = extfn14E8(l0092%, 2)
l0096% = extfn14E8(l0094%, 2)
l0098% = extfn14E8(l0096%, 2)
l009A% = extfn14E8(l0098%, 2)
l009C% = extfn14E8(l009A%, 2)
Randomize
l009E% = Int(12 * Rnd + 1)
If  l009E% = 1 Then l00A0$ = "Help!"
If  l009E% = 2 Then l00A0$ = "Help me!"
If  l009E% = 3 Then l00A0$ = "Here..."
If  l009E% = 4 Then l00A0$ = "Read this"
If  l009E% = 5 Then l00A0$ = "This happens way to often."
If  l009E% = 6 Then l00A0$ = "Argh..."
If  l009E% = 7 Then l00A0$ = "Is this real? Please read:"
If  l009E% = 8 Then l00A0$ = "Can you take care of this?"
If  l009E% = 9 Then l00A0$ = "..."
If  l009E% = 10 Then l00A0$ = "I need some help..."
If  l009E% = 11 Then l00A0$ = "Maybe you can help me..."
If  l009E% = 12 Then l00A0$ = "I got this IM..."
l00A2% = Int(30 * Rnd + 1)
If  l00A2% = 1 Then l00A4$ = "†    Hi. I got this IM from this guy - " + p0072 + "; Maybe you could read this and take care of it. I'm kinda new to America Online, so I didn't know what to do. Someone in the room said I should probably e-mail you guys since you'd take care of it."
If  l00A2% = 2 Then l00A4$ = "†    Hiya! Can you do something about this? I got this IM:"
If  l00A2% = 3 Then l00A4$ = "Help me! I almost believed this looser! He tried to steal my account information!"
If  l00A2% = 4 Then l00A4$ = "I'm new to AOL, and about ten minutes after I had signed on today, I got the following message. I took careful note of the warning on the IM window, so I copied what was said and sent it to you."
If  l00A2% = 5 Then l00A4$ = "Wow. This sure does happen enough! Can you do something about this?"
If  l00A2% = 6 Then l00A4$ = "†    Hello. Who is this guy?"
If  l00A2% = 7 Then l00A4$ = "Can you take care of this problem? I don't know what to do and some1 said you'd do something about this."
If  l00A2% = 8 Then l00A4$ = "Argh...read this:"
If  l00A2% = 9 Then l00A4$ = "Help me! Can you do something???"
If  l00A2% = 10 Then l00A4$ = "HELP! Read this:"
If  l00A2% = 11 Then l00A4$ = "This is like the millionth time this has happened to me."
If  l00A2% = 12 Then l00A4$ = "These guys that IM us and ask for our passwords and billing aren't for real, are they?"
If  l00A2% = 13 Then l00A4$ = "†          Why do they do this?  Do they steal our stuff???"
If  l00A2% = 14 Then l00A4$ = "Hi. I'm fairly new to all this (I'm not on very often). I got an IM requesting my billing info from " + p0072 + ". One of my friend's in the room said to goto Guidepager to deal with the problem, but I didn't know what she meant and she had left the room. I talked with someone else, and they told me to e-mail these addresses so I did; and to copy and paste the IM message below when I send it so here it is:"
If  l00A2% = 15 Then l00A4$ = "Argh... they won't leave me alone! I had to turn this one in - I just had to..."
If  l00A2% = 16 Then l00A4$ = "Help! This looser won't stop asking for people's stuff!"
If  l00A2% = 17 Then l00A4$ = "I wish there was an easier way to do this besides e-mailing you guys..."
If  l00A2% = 18 Then l00A4$ = "†     I like to e-mail you guys over GuidePager since they all just goto the same place, plus when I e-mail, it gurantees that my mail is read somewhere faster than it would via GuidePager."
If  l00A2% = 19 Then l00A4$ = "†  Can you give me a hand? He asked the whole room......"
If  l00A2% = 20 Then l00A4$ = "Help! This guy (or girl) won't leave me alone!!! Can you do something about this problem I'm having? Thanx!!! :-)"
If  l00A2% = 21 Then l00A4$ = "Um this person was asking for my important stuff."
If  l00A2% = 22 Then l00A4$ = "†     Here's another one for you! Your welcome! :-)"
If  l00A2% = 23 Then l00A4$ = "†   This happens way too often. I hate it. I can never real chat - I just keep getting this lame posers!"
If  l00A2% = 24 Then l00A4$ = ""
If  l00A2% = 25 Then l00A4$ = "Read:"
If  l00A2% = 26 Then l00A4$ = "Read this:               :-)"
If  l00A2% = 27 Then l00A4$ = "Read please:"
If  l00A2% = 28 Then l00A4$ = "Read this pleaze:"
If  l00A2% = 29 Then l00A4$ = "Read:          -Thanks!-"
If  l00A2% = 30 Then l00A4$ = "Help! Read this message I copied and pasted!"
l00A6$ = l00A4$ + Chr(13) + Chr(10) + Chr(13) + Chr(10) + p0072 + ":" + Chr(9) + p0076
l00A8$ = "(, TOSAdvisor, TOSEmail1, TOSEmail2, TOSEmail3, )"
l00AA% = extfn1718(l0096%, c002E, 0, l00A0$)
l00AA% = extfn1718(l009C%, c002E, 0, l00A6$)
LE898:
l00AA% = extfn1718(l008E%, c002E, 0, l00A8$)
sub170 (l007E%)
sub1050
DoEvents
l00AC = extfn1520("#32770", "America Online")
If  l00AC <> 0 Then
DoEvents
l00B2 = extfn24E0(l00AC, "OK")
l00BA = extfn1750(l00B2, c00B6, 0, 0&): DoEvents
l00C0 = extfn1750(l00B2, c00BE, 0, 0&): DoEvents
End If
subF38 (1)
l00C4% = extfn24E0(l0078%, "Error")
If  l00C4% = 0 Then Exit Sub
l00C6% = extfn2518(l00C4%, "_AOL_VIEW")
l00C8% = extfn2518(l00C4%, "_AOL_BUTTON")
l00CA$ = String(255, 0)
l00AA% = extfn1718(l00C6%, c00CC, 254, l00CA$)
l00CE% = InStr(l00CA$, "-")
l00CA$ = Left$(l00CA$, l00CE% - 2)
l00CA$ = fnF70(LCase(l00CA$))
l00CA$ = Right$(l00CA$, InStr(l00CA$, "The following problems occurred while processing your request:") + 7)
End Sub

Sub sub3478 (p00D4 As TextBox, p00D8 As TextBox)
Dim l00DA As Integer
Dim l00DC As Integer
Dim l00DE As Integer
Dim l00E0 As String
Dim l00E2 As Integer
Dim l00E4 As Integer
Dim l00E6 As Integer
Dim l00E8 As Integer
Dim l00EA As Integer
Dim l00EC As Integer
' Const c00EE = 2 ' &H2%
Dim l00F0 As Integer
Dim l00F2 As Integer
Dim l00F4 As Integer
Dim l00F6 As Integer
l00DA% = extfn1520("AOL Frame25", 0&)
If  l00DA% = 0 Then
    Beep
    MsgBox "America Online v2.5 must be running before you do this"
    Exit Sub
    End If
l00DA% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
Do Until l00DC% <> 0
    DoEvents
    l00DC% = extfn24E0(l00DA%, "Keyword")
    subF38 (.001#)
    Loop
l00DE% = extfn2518(l00DC%, "_AOL_EDIT")
l00E0$ = "GuidePager"
l00E2% = extfn1718(l00DE%, c002E, 0, CStr(l00E0$)): DoEvents
l00E4% = extfn2518(l00DC%, "_AOL_ICON")
sub170 (l00E4%): DoEvents
Do
    DoEvents
    l00E6% = extfn24E0(l00DA%, "I Need Help!"): DoEvents
    subF38 (.001#)
    Loop Until l00E6% <> 0
l00E8% = extfn2518(l00E6%, "_AOL_ICON")
sub170 (l00E8%): DoEvents
sub1050
Do Until l00EA% <> 0
    DoEvents
    l00EA% = extfn24E0(l00DA%, "Report Password Solicitations")
    subF38 (.001#)
    Loop
l00EC% = extfn24E0(l00EA%, "Screen Name of Member Soliciting Your Information:"): DoEvents
l00E2% = extfn1718(l00EC%, c002E, 0, "<-™[This guy blows goats:]™->"): DoEvents
l00F0% = extfn14E8(l00EC%, c00EE)
l00E2% = extfn1718(l00F0%, c002E, 0, p00D4): DoEvents
l00F2% = extfn24E0(l00EA%, "Copy and Paste the solicitation here:")
l00E2% = extfn1718(l00F2%, c002E, 0, "<-™[What a loser!]™->"): DoEvents
l00F4% = extfn14E8(l00F2%, c00EE)
l00E2% = extfn1718(l00F4%, c002E, 0, CStr(p00D8)): DoEvents
l00F6% = extfn24E0(l00EA%, "Send")
End Sub

Sub sub34B0 (ByVal p00F8 As String)
Dim l00FA As Integer
Dim l00FC As Integer
Dim l00FE As Integer
Dim l0100 As Integer
Dim l0102 As String
Dim l0104 As Integer
l00FA% = extfn1520("AOL Frame25", 0&)
l00FC% = extfn2518(l00FA%, "_AOL_LISTBOX")
l00FE% = extfn1600(l00FC%)
l0100% = extfn2518(l00FE%, "_AOL_VIEW")
l0102$ = Chr(13) + Chr(10) + Chr(13) + Chr(10) + "OnlineHost:" + Chr(9) + p00F8 + Chr(13) + Chr(10) + Chr(13) + Chr(10)
l0104% = extfn1718(l0100%, c002E, 0, l0102$)
DoEvents
End Sub

Function fn34E8 (p010A As String) As Variant
Dim l010C As String
Dim l010E As Integer
Dim l0110 As Integer
Dim l0112 As String
Dim l0114 As String
Let l010C$ = p010A
Let l010E% = Len(l010C$)
Do While l0110% <= l010E%
Let l0110% = l0110% + 1
Let l0112$ = Mid$(l010C$, l0110%, 1)
Let l0112$ = l0112$
Let l0114$ = l0114$ + l0112$
Loop
fn34E8 = l0114$
End Function

Sub sub3520 (p0118 As TextBox, p011C As TextBox)
Dim l011E As Integer
Dim l0120 As Integer
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126 As Integer
Dim l0128 As Integer
Dim l012A As Integer
Dim l012C As Integer
Dim l012E As Integer
Dim l0130 As Integer
Dim l0132 As Integer
Dim l0134 As Integer
Dim l0136 As Variant
Dim l013A As Integer
' Const c013C = 16 ' &H10%
Dim l0140 As Integer
Call subD08("Keyword...", "&Go To"): DoEvents
l011E% = extfn1520("AOL Frame25", 0&)
l0120% = extfn24E0(l011E%, "Keyword")
l0122% = extfn2518(l0120%, "_AOL_EDIT")
l0124% = extfn1718(l0122%, c002E, 0, "KO Help"): DoEvents
l0126% = extfn2518(l0120%, "_AOL_ICON")
sub170 (l0126%)
DoEvents
subF38 (1)
l0128% = 0
Do Until l0128% <> 0
    l0128% = extfn24E0(l011E%, "GUIDEPAGER FOR KIDS")
    DoEvents
    subF38 (.001#)
    Loop
sub1050
l012A% = extfn2518(l0128%, "_AOL_ICON")
l012C% = extfn14E8(l012A%, c00EE)
l012E% = extfn14E8(l012C%, c00EE)
sub170 (l012E%): DoEvents
sub1050
Do
l0130% = extfn24E0(l011E%, "Report Password Solicitation")
subF38 .2#
Loop Until l0130% <> 0
l0132% = extfn24E0(l0130%, "Screen Name of Member Soliciting Your Information:")
l0134% = extfn14E8(l0132%, c00EE)
l0136 = extfn1718(l0134%, c002E, 0, p0118)
l0132% = extfn24E0(l0130%, "Copy and Paste the solicitation here:")
l013A% = extfn14E8(l0132%, c00EE)
l0136 = extfn1718(l013A%, c002E, 0, p0118 & ":" & Chr(9) & p011C)
subF38 .1#
l0132% = extfn24E0(l0130%, "SEND")
sub170 (l0132%)
DoEvents
subF38 .2#
subFE0
l0124% = extfn16E0(l0130%, c013C, 0, 0): DoEvents
l0140% = extfn24E0(l011E%, "GUIDEPAGER FOR KIDS")
l0124% = extfn1750(l0140%, c013C, 0, 0): DoEvents
DoEvents
End Sub

Function fn3558 () As Integer
Dim l0144 As Integer
Dim l0146 As Integer
Dim l0148 As Integer
Dim l014A As Integer
Dim l014C As Integer
Dim l014E As Integer
Dim l0150 As Integer
Dim l0152 As Integer
Dim l0154 As Integer
Dim l0156 As Integer
Dim l0158 As Integer
l0144% = extfn1520("AOL Frame25", 0&)
Do Until fn3558() <> 0
    l0146% = extfn2518(l0144%, "_AOL_ICON")
    l0148% = extfn2518(l0144%, "_AOL_LISTBOX")
    l014A% = extfn2518(l0144%, "_AOL_VIEW")
    l014C% = extfn2518(l0144%, "_AOL_EDIT")
    l014E% = extfn2518(l0144%, "_AOL_GLYPH")
    l0150% = extfn1600(l0146%)
    l0152% = extfn1600(l0148%)
    l0154% = extfn1600(l014A%)
    l0156% = extfn1600(l014C%)
    l0158% = extfn1600(l014E%)
    If  l0150% = l0152% And l0152% = l0154% And l0154% = l0156% And l0156% = l0158% Then
        fn3558 = l0152%
        End If
    subF38 (.01#)
    Loop
End Function

Sub sub3590 ()
Dim l015A As Variant
Dim l015E As Variant
Dim l0162 As Variant
Dim l0166 As Variant
Dim l016A As Variant
Dim l016E As Variant
Dim l0172 As Variant
Dim l0176 As Variant
Dim l017A As Variant
l015A = "†  |\ØØ\ØØ\                      |ØØ|          /ØØ\"
l015E = "|ØØ|\__\   |                /ØØ/|__|Ø|Ø ||'Ø|    |\ \/ØØ\"
l0162 = "†>  >    |  |/Ø \|ØØ| |ØØ|\__'ØØ\|Ø|   ||  |    |  \/|    |"
l0166 = "|__|/ØØ/|__/\__\__\|   |ØØ|/__/|_|\___/\__\   |    |"
l016A = "|   /__/  |       '/ØØ\/   /|__|    ||  | |    |  |   | \/__/|"
l016E = "|_ |   |  /|___ /__/\__/ |    |__||_|\|__.|/\|_.| '|   | |"
l0172 = "†  |__|/         |   |  |   | |__|  v2.o - ƒÒ¸ﬂÔö |__|/"
l0176 = "†                 |__|/|__|/"
sub30B0 (l015A)
sub30B0 (l015E)
sub30B0 (l0162)
subF38 (1.9#)
sub30B0 (l0166)
sub30B0 (l016A)
sub30B0 (l016E)
subF38 (1.9#)
sub30B0 (l0172)
sub30B0 (l0176)
sub30B0 (l017A)
End Sub

Sub sub35C8 ()
Dim l017E As Variant
Dim l0182 As Variant
Dim l0186 As Variant
Dim l018A As Variant
Dim l018E As Variant
Dim l0192 As Variant
Dim l0196 As Variant
Dim l019A As Variant
l017E = "†   ___                         _" + Chr(13)
l0182 = "†  |   _| |Ø|              \Ø\   |\|" + Chr(13)
l0186 = "†   \  \   \  \ \Ø\ \Ø\/Ø/ \_\ |Ø| \Ø\   /Ø/\Ø\ \Ø/ /Ø/" + Chr(13)
l018A = "†   /  /_  |  |  |  | |  |__ØØ\ \  \ |  | '|  |  |  | \Ø/|  |" + Chr(13)
l018E = "†  |___| /_/ _\_\/  /\Ø\ \_''\ |_| \_\/_/  /_/:|   |\_\" + Chr(13)
l0192 = "†  |:::::| |::||:\_\/_/ |:\_\|::::||::|:'|:::::|:|'|::|:|   |:|::|" + Chr(13)
l0196 = "†  |:::::| |::| \'|:::::| ' \|::||::::||::|\'|:::::|/ '|::|/    \'|::|" + Chr(13)
l019A = "†  'ØØØ   Ø   'ØØØ     Ø  ØØ  Ø   ØØØ   Ø         Ø" + Chr(13)
sub30B0 (l017E)
sub30B0 (l0182)
sub30B0 (l0186)
subF38 (1.9#)
sub30B0 (l018A)
sub30B0 (l018E)
sub30B0 (l0192)
subF38 (1.9#)
sub30B0 (l0196)
sub30B0 (l019A)
End Sub

Sub sub3600 (p019E As String, p01A0 As String)
Dim l01A2 As Variant
Dim l01A6() As Variant
Dim l01AC As Variant
Dim l01B0 As Variant
p01A0 = ""
l01A2 = Len(p019E)
ReDim l01A6(1 To l01A2)
For l01AC = 1 To l01A2
69 l01B0 = Int(Rnd * l01A2) + 1
If  l01A6(l01B0) = 1 Then GoTo 69
p01A0 = p01A0 + Mid(p019E, l01B0, 1)
l01A6(l01B0) = 1
Next l01AC
End Sub

Sub sub3638 (p01B4 As String, p01B6 As String)
Dim l01B8 As Variant
Dim l01BC As Variant
Dim l01C0 As String
Dim l01C2 As String
Dim l01C4 As Variant
p01B4 = p01B4 + " x"
l01B8 = ""
Do Until InStr(p01B4, " ") = 0
    l01BC = InStr(1, p01B4, " ")
    l01C0$ = Left$(p01B4, l01BC - 1)
    p01B4 = Right(p01B4, Len(p01B4) - l01BC)
    Call sub3600(l01C0$, l01C2$)
    l01C4 = l01C4 + l01C2$ + " "
Loop
p01B6 = l01C4
End Sub

Sub sub3670 (p01C8 As Variant, p01CC As Variant)
Dim l01D0 As Variant
Dim l01D4 As Variant
Dim l01D8 As Variant
Dim l01DC As Variant
Dim l01E0 As Integer
Dim l01E2 As Integer
Dim l01E4 As Variant
Dim l01E8 As Variant
l01D0 = Len(p01C8)
For l01D4 = 0 To l01D0
    l01D8 = Left$(p01C8, 1)
    p01C8 = Right(p01C8, -l01DC)
    l01DC = Len(p01C8)
    Randomize
    l01E0% = (l01DC * Rnd + 1)
    l01E2% = (2 * Rnd + 1)
    If  l01E2% = 1 Then
        l01E4 = Left$(p01C8, l01E0% - 1)
        Else
        l01E4 = Left$(p01C8, l01E0% + 1)
        End If
    l01E8 = Right$(p01C8, l01E0%)
    p01CC = l01E4 + l01D8 + l01E8
    MsgBox l01E4
    MsgBox l01D8
    MsgBox l01E8
Next l01D4
End Sub

Sub sub36A8 (p01EC As Variant)
Do
sub30B0 (p01EC)
subF38 (1)
Loop
End Sub

Sub sub36E0 (p01F2 As TextBox)
sub30B0 ("†ïñï  £ˇßÔ¸mô v2.0 ïñï —ÎÜhÎrw¯rld TÿäsÎr v2.0ïñï")
sub30B0 ("†               Black, Red, Yellow, White")
sub30B0 ("†               Race, Creed, Color, Rights")
subF38 (2.5#)
sub30B0 ("†                 Love, Joy, Sorrow, Hate")
sub30B0 ("†                 Life, Death, Time, Waste")
sub30B0 ("Don't you see what I mean, Live your life so foolishly")
subF38 (2.5#)
sub30B0 ("Keep your mind clear at night - T-Bone's a way of life!")
sub30B0 ("†               Peace, War, Guns, Waste")
sub30B0 ("†                Dirt, Water, Mud, Paste")
subF38 (2.5#)
sub30B0 ("†                Flesh, Bone, Tear, Rip")
sub30B0 ("†               Skin, Color, Death, Trip")
sub30B0 ("†")
subF38 (2.5#)
sub30B0 ("†     Don't you see what I mean when I mean")
sub30B0 ("Take your time, take your time and live and breathe")
sub30B0 ("† Open your eyes, scrape away the hate and greed")
subF38 (2.5#)
sub30B0 ("†  Wash your hands, wash away race and creed")
sub30B0 ("† Free your mind, free your mind and you will see")
sub30B0 ("†  Life's too short, way too short to hate you")
subF38 (2.5#)
sub30B0 ("†Breed, strip you fear, rip away what you feel")
sub30B0 (p01F2 + ", contrary to what I just scrolled...")
sub30B0 ("ïñõNetherworld is now TOSsingóõ " + p01F2 + " ïóß¯rrˇóï")
End Sub

Sub sub3718 (p01F6 As TextBox)
sub30B0 ("†ïñï  £ˇßÔ¸mô v2.0 ïñï —ÎÜhÎrw¯rld TÿäsÎr v2.0ïñï")
sub30B0 ("†††††††††ïñ It was the day AOL stood still... ñï")
sub30B0 ("††ïñ The day when machines became self aware ñï")
subF38 (2.5#)
sub30B0 ("†††ïñ It was the day we all feared would come. ñï")
sub30B0 ("†ïñ And from this day forth, history will record this ñï")
sub30B0 ("ïñ as the War between " + fn7C8() + " and " + p01F6 + ". ñï")
End Sub
