' FAC.FRM
Option Explicit

Sub sub9748 ()
End Sub

Sub Command3D1_Click ()
Dim l0020 As Variant
Dim l0066 As Variant
Dim l006E As Variant
Dim l0074 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0092 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C6 As Variant
Dim l00CA As Variant
Dim l00CE As Variant
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0108 As Variant
Dim l010C As Variant
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
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
Dim l0162 As Variant
Dim l0166 As Variant
Dim l016A As Variant
Dim l016E As Variant
Dim l0172 As Variant
Dim l0176 As Variant
Dim l017A As Variant
Dim l017E As Variant
Dim l0182 As Variant
Dim l0186 As Variant
Dim l018A As Variant
Dim l018E As Variant
Dim l0192 As Variant
Dim l0196 As Variant
Dim l019A As Variant
Dim l019E As Variant
Dim l01A2 As Variant
Dim l01A6 As Variant
Dim l01AA As Variant
Dim l01AE As Variant
Dim l01B2 As Variant
Dim l01B6 As Variant
Dim l01BA As Variant
Dim l01BE As Variant
l0020 = InStr(Prefix.Text, "-")
If  l0020 Then
MsgBox "Please enter in the prefix without the "" - """, 16
Exit Sub
End If
If  fn.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  fn.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  address.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  city.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  state.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  zipcode.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  phone.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  cn.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  cp.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  ccn.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  bank.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
If  expiration.Text = "" Then
MsgBox "Please fill out all the information!", 16
Exit Sub
End If
l0066 = Mid(Prefix.Text, 1, 1)
If  l0066 = "5" Then GoTo 24
If  l0066 = "4" Then
GoTo 24
Else
MsgBox "This fake account generator only supports Visa and Mastercard!", 16
Exit Sub
End If
24 :
DoEvents
l0020 = extfn1018("_AOL_MODAL", 0&)
l006E = extfn10F8(l0020, "Registration")
If  l006E = 0 Then
MsgBox "You Must Sign On With New User!", 16
Exit Sub
End If
l0074 = extfn11A0(l0020, 12, 0, "AOL1")
l006E = extfn1168(l006E, 2)
l007A = extfn11A0(l006E, 12, 0, cn.Text)
l007E = extfn10F8(l0020, "Password")
l007E = extfn1168(l007E, 2)
l0082 = extfn11A0(l007E, 12, 0, cp.Text)
l0088 = fn13A0(.1#)
l008C = extfn10F8(l0020, "Continue")
l0092 = fn1480(l008C)
Do Until 2 > 3
DoEvents
If  fn7C8() <> "AOL1" Then GoTo 10
Loop
10 :
DoEvents
l0098 = extfn1018("_AOL_MODAL", 0&)
If  l0098 = 0 Then GoTo 10
l0066 = fn13A0(.7#)
l0074 = extfn11A0(l0098, 12, 0, "AOL2")
l009C = extfn10F8(l0098, "Continue")
l00A0 = fn1480(l009C)
DoEvents
Do Until 2 > 3
DoEvents
If  fn7C8() <> "AOL2" Then GoTo 20
Loop
20 :
l0066 = fn13A0(.7#)
202 :
DoEvents
l00A4 = extfn1018("_AOL_MODAL", 0&)
If  l00A4 = 0 Then GoTo 202
l00A8 = extfn11A0(l00A4, 12, 0, "AOL3")
l00AC = extfn10F8(l00A4, "First Name:")
l00B0 = extfn1168(l00AC, 2)
l00B4 = extfn11A0(l00B0, 12, 0, fn.Text)
l00B8 = extfn10F8(l00A4, "Last Name:")
l00BC = extfn1168(l00B8, 2)
l00C6 = extfn11A0(l00BC, 12, 0, ln.Text)
l00CA = extfn10F8(l00A4, "Address:")
l00CE = extfn1168(l00CA, 2)
l00D2 = extfn11A0(l00CE, 12, 0, address.Text)
l00D6 = extfn10F8(l00A4, "State:")
l00DA = extfn1168(l00D6, 2)
l00DE = extfn11A0(l00DA, 12, 0, state.Text)
l00D6 = extfn10F8(l00A4, "City:")
l00DA = extfn1168(l00D6, 2)
l00DE = extfn11A0(l00DA, 12, 0, city.Text)
l00D6 = extfn10F8(l00A4, "Zip Code:")
l00DA = extfn1168(l00D6, 2)
l00DE = extfn11A0(l00DA, 12, 0, zipcode.Text)
l00D6 = extfn10F8(l00A4, "Daytime Phone:")
l00DA = extfn1168(l00D6, 2)
l00DE = extfn11A0(l00DA, 12, 0, phone.Text)
l00D6 = extfn10F8(l00A4, "Evening Phone:")
l00DA = extfn1168(l00D6, 2)
l00DE = extfn11A0(l00DA, 12, 0, phone.Text)
l00E2 = extfn10F8(l00A4, "Continue")
l0066 = fn1480(l00E2)
l00E6 = fn13A0(.1#)
30 :
DoEvents
If  fn7C8() = "AOL3" Then GoTo 30
DoEvents
303 :
DoEvents
l00EA = extfn1018("_AOL_MODAL", 0&)
If  l00EA = 0 Then GoTo 303
l0020 = fn13A0(.7#)
l0074 = extfn11A0(l00EA, 12, 0, "AOL4")
l0088 = extfn10F8(l00EA, "Continue")
l00EE = fn1480(l0088)
80 :
DoEvents
If  fn7C8() = "AOL4" Then GoTo 80
DoEvents
808 :
DoEvents
l00F2 = extfn1018("_AOL_MODAL", 0&)
If  l00F2 = 0 Then GoTo 808
l0020 = fn13A0(.7#)
l0074 = extfn11A0(l00F2, 12, 0, "AOL5")
l00F6 = extfn10F8(l00F2, "Continue")
l00FA = fn1480(l00F6)
60 :
DoEvents
If  fn7C8() = "AOL5" Then GoTo 60

61 :
DoEvents
l00FE = extfn1018("AOL FRAME25", 0&)
l0102 = extfn1018("_AOL_MODAL", "Billing Options")
If  l0102 = 0 Then GoTo 61
l0066 = fn13A0(.7#)
l0108 = extfn10C0(l0102, "_AOL_LISTBOX")
l010C = Mid(ccn.Text, 1, 1)
If  l010C = "4" Then
l0112 = extfnFE0(l0108, 1031, 1, 0)
End If
If  l010C = "5" Then
l0112 = extfnFE0(l0108, 1031, 2, 0)
End If
l0116 = extfn10F8(l0102, "Select")
l011A = fn1480(l0116)

40 :
DoEvents
l011E = extfn1018("_AOL_MODAL", "Billing Options")
If  l011E > 0 Then GoTo 40
123 :
DoEvents
l011E = extfn1018("_AOL_MODAL", 0&)
If  l011E = 0 Then GoTo 123
l0066 = fn13A0(.7#)
l0074 = extfn11A0(l011E, 12, 0, "AOL6")
l0122 = extfn10F8(l011E, "Card Number")
l0126 = extfn1168(l0122, 2)
l012A = extfn11A0(l0126, 12, 0, ccn.Text)
l012E = extfn10F8(l011E, "Expiration")
l0132 = extfn1168(l012E, 2)
l0136 = extfn11A0(l0132, 12, 0, expiration.Text)
l013A = extfn10F8(l011E, "Bank Name")
l013E = extfn1168(l013A, 2)
l0142 = extfn11A0(l013E, 12, 0, bank.Text)
l0146 = extfn10F8(l011E, "First Name")
l014A = extfn1168(l0146, 2)
l014E = extfn11A0(l014A, 12, 0, fn.Text)
l0152 = extfn10F8(l011E, "Last Name")
l0156 = extfn1168(l0152, 2)
l015A = extfn11A0(l0156, 12, 0, ln.Text)
l015E = extfn10F8(l011E, "Continue")
l0162 = fn1480(l015E)
70 :
DoEvents
l0166 = extfn10F8(l011E, "Bank Name")
l016A = extfn1018("#32770", "America Online")
If  l016A > 0 Then
l016E = extfn10F8(l016A, "OK")
l0172 = fn1480(l016E)
sub9748
l0176 = extfn11A0(l0126, 12, 0, ccn.Text)
l017A = fn1480(l015E)
End If
If  l0166 > 0 Then GoTo 70

56 :
DoEvents
If  fn7C8() = "AOL6" Then GoTo 56
560 :
DoEvents
l017E = extfn1018("_AOL_MODAL", 0&)
If  l017E = 0 Then GoTo 560
l0066 = fn13A0(.7#)
l0074 = extfn11A0(l017E, 12, 0, "AOL7")
l0182 = extfn10F8(l017E, "Continue")
l0186 = fn1480(l0182)

91 :
DoEvents
If  fn7C8() = "AOL7" Then GoTo 91
910 :
DoEvents
l018A = extfn1018("_AOL_MODAL", 0&)
If  l018A = 0 Then GoTo 910
l0066 = fn13A0(.7#)
l018E = extfn10F8(l018A, "I Agree")
l0074 = extfn11A0(l018A, 12, 0, "AOL8")
l0192 = fn1480(l018E)
l0066 = fn13A0(.7#)
l0196 = extfn1018("AOL FRAME25", 0&)
l019A = extfn10F8(l0196, "Goodbye")
If  l019A > 0 Then
GoTo L331CA
End If

100 :
DoEvents
l019E = extfn1018("_AOL_MODAL", 0&)
l01A2 = extfn10F8(l019E, "Continue")
l0196 = extfn1018("AOL FRAME25", 0&)
l019A = extfn10F8(l0196, "Goodbye")
If  l019A > 0 Then
GoTo L331CA
End If
If  l01A2 = 0 Then GoTo 100
l0066 = fn13A0(.7#)
l0196 = extfn1018("AOL FRAME25", 0&)
l019A = extfn10F8(l0196, "Goodbye")
If  l019A > 0 Then
GoTo L331CA
End If
l01A6 = extfn11A0(l019E, 12, 0, "AOL9")
l01AA = fn1480(l01A2)

110 :
DoEvents
l0196 = extfn1018("AOL FRAME25", 0&)
l019A = extfn10F8(l0196, "Goodbye")
If  l019A > 0 Then
GoTo L331CA
End If
If  fn7C8() = "AOL9" Then GoTo 110

134 :
DoEvents
l01AE = extfn1018("_AOL_MODAL", 0&)
l01B2 = extfn10F8(l01AE, "Continue")
If  l01B2 = 0 Then GoTo 134
l0066 = fn13A0(.7#)
l01B6 = fn1480(l01B2)

120 :
DoEvents
l01BA = extfn1018("_AOL_MODAL", 0&)
l01BE = extfn10C0(l01BA, "_AOL_EDIT")
l0196 = extfn1018("AOL FRAME25", 0&)
l019A = extfn10F8(l0196, "Goodbye")
If  l019A > 0 Then
GoTo L331CA
End If
If  l01BE = 0 Then GoTo 120

L331CA:
End Sub

Sub Command3D2_Click ()
fac.WindowState = 1
End Sub

Sub Command3D3_Click ()
Dim l01C6 As Variant
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
Static m01FA(16) As Variant
Dim l0210 As Variant
Dim l0214 As Variant
Dim l0218 As Variant
Dim l021C As Variant
Dim l0220 As Variant
Dim l0224 As Variant
Dim l0228 As Variant
Dim l022C As Variant
Dim l0230 As Variant
Dim l0234 As Variant
Dim l0238 As Variant
Dim l023C As Variant
Dim l0240 As Variant
Dim l0244 As Variant
Dim l0248 As Variant
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
Dim l0280 As Variant
Dim l0284 As Variant
Dim l0288 As Variant
Dim l028C As Variant
Dim l0290 As Variant
Dim l0294 As Variant
Dim l0298 As Variant
Dim l029C As Variant
Dim l02A0 As Variant
Dim l02A4 As Variant
Dim l02A8 As Variant
Dim l02AC As Variant
Dim l02B0 As Variant
Dim l02B4 As Variant
Dim l02B8 As Variant
Dim l02BC As Variant
Dim l02C0 As Variant
Dim l02C4 As Variant
Dim l02C8 As Variant
Dim l02CC As Variant
Dim l02D0 As Variant
Dim l02D4 As Variant
Dim l02D8 As Variant
Dim l02DC As Variant
Dim l02E0 As Variant
Dim l02E4 As Variant
Dim l02E8 As Variant
Dim l02EC As Variant
Dim l02F0 As Variant
Dim l02F4 As Variant
Dim l02F8 As Variant
l01C6 = InStr(Prefix.Text, "-")
If  l01C6 Then
MsgBox "Please enter in the prefix without the "" - """, 16
Exit Sub
End If
l01C6 = InStr(Prefix.Text, " ")
If  l01C6 Then
MsgBox "Please enter in the prefix without the space!", 16
Exit Sub
End If
If  Len(Prefix.Text) > 8 Then
MsgBox "This only supports up to 8 digit prefixes!", 16
Exit Sub
End If
If  Not IsNumeric(Prefix.Text) Then
MsgBox "That is an invalid prefix!", 16
Exit Sub
End If
Randomize
l01C6 = Int(3 * Rnd)
Select Case l01C6
Case 0
l01CA = "7"
Case 1
l01CA = "8"
Case 2
l01CA = "9"
End Select
Randomize
l01CE = Int(12 * Rnd) + 1
expiration.Text = l01CE + "/" + "9" + l01CA
Randomize
l01D2 = Int(224 * Rnd)
Select Case l01D2
Case 0
l01D6 = "Amy"
Case 1
l01D6 = "Dave"
Case 2
l01D6 = "Dan"
Case 3
l01D6 = "Bob"
Case 4
l01D6 = "Kevin"
Case 5
l01D6 = "Matt"
Case 6
l01D6 = "Eric"
Case 7
l01D6 = "Tiffany"
Case 8
l01D6 = "Brian"
Case 9
l01D6 = "Tot"
Case 10
l01D6 = "Stevie"
Case 11
l01D6 = "John"
Case 12
l01D6 = "Michael"
Case 13
l01D6 = "Mark"
Case 14
l01D6 = "Ken"
Case 15
l01D6 = "Charles"
Case 16
l01D6 = "Paul"
Case 17
l01D6 = "Carl"
Case 18
l01D6 = "Alan"
Case 19
l01D6 = "Steve"
Case 20
l01D6 = "Bill"
Case 21
l01D6 = "Fred"
Case 22
l01D6 = "Wilma"
Case 23
l01D6 = "Victor"
Case 24
l01D6 = "Hector"
Case 25
l01D6 = "Alex"
Case 26
l01D6 = "Alexander"
Case 27
l01D6 = "Gabby"
Case 28
l01D6 = "Jamal"
Case 29
l01D6 = "Scott"
Case 30
l01D6 = "Jake"
Case 31
l01D6 = "Aldo"
Case 32
l01D6 = "Jack"
Case 33
l01D6 = "Ferdick"
Case 34
l01D6 = "Arash"
Case 35
l01D6 = "Brandon"
Case 36
l01D6 = "Bruce"
Case 37
l01D6 = "Janice"
Case 38
l01D6 = "Allen"
Case 39
l01D6 = "Fallon"
Case 40
l01D6 = "Saad"
Case 41
l01D6 = "Johnny"
Case 42
l01D6 = "Danny"
Case 43
l01D6 = "Jason"
Case 44
l01D6 = "Jonathan"
Case 45
l01D6 = "Chris"
Case 46
l01D6 = "Richard"
Case 47
l01D6 = "Andrea"
Case 48
l01D6 = "Amir"
Case 49
l01D6 = "Bobby"
Case 50
l01D6 = "George"
Case 51
l01D6 = "Jacob"
Case 52
l01D6 = "Jay"
Case 53
l01D6 = "Donald"
Case 54
l01D6 = "Robert"
Case 55
l01D6 = "Stephen"
Case 56
l01D6 = "Ronald"
Case 57
l01D6 = "Thomas"
Case 58
l01D6 = "Kirby"
Case 59
l01D6 = "Patrick"
Case 60
l01D6 = "Ryan"
Case 61
l01D6 = "Walter"
Case 62
l01D6 = "Kalvin"
Case 63
l01D6 = "Zhen"
Case 64
l01D6 = "Peter"
Case 65
l01D6 = "Tim"
Case 66
l01D6 = "Timothy"
Case 67
l01D6 = "Rodney"
Case 68
l01D6 = "Ray"
Case 69
l01D6 = "Gary"
Case 70
l01D6 = "Gilbert"
Case 71
l01D6 = "Earl"
Case 72
l01D6 = "Lubert"
Case 73
l01D6 = "Virginia"
Case 74
l01D6 = "Mark"
Case 75
l01D6 = "Jane"
Case 76
l01D6 = "Jordan"
Case 77
l01D6 = "Howard"
Case 78
l01D6 = "Phuong"
Case 79
l01D6 = "Tot"
Case 80
l01D6 = "Paul"
Case 81
l01D6 = "Jeffrey"
Case 82
l01D6 = "Lucey"
Case 83
l01D6 = "William"
Case 84
l01D6 = "Chase"
Case 85
l01D6 = "Matthew"
Case 86
l01D6 = "Phillip"
Case 87
l01D6 = "Nancy"
Case 88
l01D6 = "Andy"
Case 89
l01D6 = "Andrew"
Case 90
l01D6 = "Luik"
Case 91
l01D6 = "Kimberly"
Case 92
l01D6 = "Billy"
Case 93
l01D6 = "Micky"
Case 94
l01D6 = "Nadine"
Case 95
l01D6 = "Gene"
Case 96
l01D6 = "Logan"
Case 97
l01D6 = "Glen"
Case 98
l01D6 = "Gloria"
Case 99
l01D6 = "Destiny"
Case 100
l01D6 = "Axel"
Case 101
l01D6 = "Khiem"
Case 102
l01D6 = "Gregory"
Case 103
l01D6 = "Margaret"
Case 104
l01D6 = "Susan"
Case 105
l01D6 = "Ralph"
Case 106
l01D6 = "Willy"
Case 107
l01D6 = "Donatello"
Case 108
l01D6 = "Mike"
Case 109
l01D6 = "Kim"
Case 110
l01D6 = "Rob"
Case 111
l01D6 = "Dawn"
Case 112
l01D6 = "Liem"
Case 113
l01D6 = "Hipolito"
Case 114
l01D6 = "Natasha"
Case 115
l01D6 = "Edward"
Case 116
l01D6 = "Mark"
Case 117
l01D6 = "Lynn"
Case 118
l01D6 = "Joel"
Case 119
l01D6 = "Noelle"
Case 120
l01D6 = "Bunty"
Case 121
l01D6 = "Akira"
Case 122
l01D6 = "Ashim"
Case 123
l01D6 = "Clint"
Case 124
l01D6 = "Silvia"
Case 125
l01D6 = "Alicia"
Case 126
l01D6 = "Bart"
Case 127
l01D6 = "Homer"
Case 128
l01D6 = "Marge"
Case 129
l01D6 = "Lisa"
Case 130
l01D6 = "Humberto"
Case 131
l01D6 = "Macki"
Case 132
l01D6 = "Mary"
Case 133
l01D6 = "Carlos"
Case 134
l01D6 = "Tommy"
Case 135
l01D6 = "Kent"
Case 136
l01D6 = "Clark"
Case 137
l01D6 = "Joseph"
Case 138
l01D6 = "Patti"
Case 139
l01D6 = "Pat"
Case 140
l01D6 = "Randolph"
Case 141
l01D6 = "Ida"
Case 142
l01D6 = "Harry"
Case 143
l01D6 = "Larry"
Case 144
l01D6 = "Moe"
Case 145
l01D6 = "Warren"
Case 146
l01D6 = "Neil"
Case 147
l01D6 = "Aneil"
Case 148
l01D6 = "Renee"
Case 149
l01D6 = "Rita"
Case 150
l01D6 = "Douglas"
Case 151
l01D6 = "Charles"
Case 152
l01D6 = "Daniel"
Case 153
l01D6 = "Laura"
Case 154
l01D6 = "Laureen"
Case 155
l01D6 = "Martin"
Case 156
l01D6 = "Marty"
Case 157
l01D6 = "Jose"
Case 158
l01D6 = "Josephine"
Case 159
l01D6 = "Marwan"
Case 160
l01D6 = "Narwin"
Case 161
l01D6 = "Ashley"
Case 162
l01D6 = "Jerry"
Case 163
l01D6 = "Patricia"
Case 164
l01D6 = "Melissa"
Case 165
l01D6 = "Anthony"
Case 166
l01D6 = "Tony"
Case 167
l01D6 = "Kathleen"
Case 168
l01D6 = "Kathy"
Case 169
l01D6 = "Kristin"
Case 170
l01D6 = "Kelly"
Case 171
l01D6 = "Bud"
Case 172
l01D6 = "Frank"
Case 173
l01D6 = "Debbie"
Case 174
l01D6 = "Ellen"
Case 175
l01D6 = "Eugene"
Case 176
l01D6 = "Sue"
Case 177
l01D6 = "Sully"
Case 178
l01D6 = "Ricky"
Case 179
l01D6 = "Rick"
Case 180
l01D6 = "Peggy"
Case 181
l01D6 = "Keith"
Case 182
l01D6 = "Otis"
Case 183
l01D6 = "Weis"
Case 184
l01D6 = "Dick"
Case 185
l01D6 = "David"
Case 186
l01D6 = "Daniel"
Case 187
l01D6 = "Carol"
Case 188
l01D6 = "Kenneth"
Case 189
l01D6 = "Curtis"
Case 190
l01D6 = "Erica"
Case 191
l01D6 = "Craig"
Case 192
l01D6 = "Doug"
Case 193
l01D6 = "Ian"
Case 194
l01D6 = "James"
Case 195
l01D6 = "Jim"
Case 196
l01D6 = "Jimbo"
Case 197
l01D6 = "Margery"
Case 198
l01D6 = "Glenn"
Case 199
l01D6 = "Emlin"
Case 200
l01D6 = "Erin"
Case 201
l01D6 = "Shelly"
Case 202
l01D6 = "Stewart"
Case 203
l01D6 = "Milhouse"
Case 204
l01D6 = "Shirley"
Case 205
l01D6 = "Clarence"
Case 206
l01D6 = "Dennis"
Case 207
l01D6 = "Debra"
Case 208
l01D6 = "Diane"
Case 209
l01D6 = "Christopher"
Case 210
l01D6 = "Wilburt"
Case 211
l01D6 = "Garret"
Case 212
l01D6 = "Derek"
Case 213
l01D6 = "Casie"
Case 214
l01D6 = "Theresa"
Case 215
l01D6 = "Ronald"
Case 216
l01D6 = "Ron"
Case 217
l01D6 = "Nicole"
Case 218
l01D6 = "Russ"
Case 219
l01D6 = "Brian"
Case 220
l01D6 = "Malcom"
Case 221
l01D6 = "Jin"
Case 222
l01D6 = "Benjamin"
Case 223
l01D6 = "Carlina"
End Select
fn.Text = l01D6
Randomize
l01DA = Int(100 * Rnd)
Select Case l01DA
Case 0
l01DE = "Anderson"
Case 1
l01DE = "Banya"
Case 2
l01DE = "Clint"
Case 3
l01DE = "Nguyen"
Case 4
l01DE = "Russian"
Case 5
l01DE = "German"
Case 6
l01DE = "Mulhair"
Case 7
l01DE = "Lincoln"
Case 8
l01DE = "Abraham"
Case 9
l01DE = "Ramdeo"
Case 10
l01DE = "Barker"
Case 11
l01DE = "Eck"
Case 12
l01DE = "Eckerman"
Case 13
l01DE = "Baig"
Case 14
l01DE = "MCGinty"
Case 15
l01DE = "Clark"
Case 16
l01DE = "Morgan"
Case 17
l01DE = "Rieder"
Case 18
l01DE = "Hicks"
Case 19
l01DE = "Farrow"
Case 20
l01DE = "Jennings"
Case 21
l01DE = "Bibbee"
Case 22
l01DE = "Piddington"
Case 23
l01DE = "Lott"
Case 24
l01DE = "Bailey"
Case 25
l01DE = "Hershey"
Case 26
l01DE = "Long"
Case 27
l01DE = "Luong"
Case 28
l01DE = "Vasquez"
Case 29
l01DE = "Ercillo"
Case 30
l01DE = "Trittle"
Case 31
l01DE = "Boxill"
Case 32
l01DE = "Boy"
Case 33
l01DE = "Smith"
Case 34
l01DE = "Simpson"
Case 35
l01DE = "Gates"
Case 36
l01DE = "Clinton"
Case 37
l01DE = "Bush"
Case 38
l01DE = "Kerrigan"
Case 39
l01DE = "McMillan"
Case 40
l01DE = "Hunolt"
Case 41
l01DE = "Sellers"
Case 42
l01DE = "MacGuire"
Case 43
l01DE = "Schroeder"
Case 44
l01DE = "Dalmut"
Case 45
l01DE = "Kamsivone"
Case 46
l01DE = "Phem"
Case 47
l01DE = "Pham"
Case 48
l01DE = "Berry"
Case 49
l01DE = "Byrd"
Case 50
l01DE = "Ghadir"
Case 51
l01DE = "Lee"
Case 52
l01DE = "Warner"
Case 53
l01DE = "Leeds"
Case 54
l01DE = "Walton"
Case 55
l01DE = "Nichols"
Case 56
l01DE = "Waddell"
Case 57
l01DE = "Spellman"
Case 58
l01DE = "Ahoe"
Case 59
l01DE = "Leesburg"
Case 60
l01DE = "Bryan"
Case 61
l01DE = "Bryant"
Case 62
l01DE = "Slager"
Case 63
l01DE = "Le"
Case 64
l01DE = "Decamp"
Case 65
l01DE = "Haith"
Case 66
l01DE = "Tran"
Case 67
l01DE = "Machi"
Case 68
l01DE = "Davis"
Case 69
l01DE = "Alam"
Case 70
l01DE = "Mary"
Case 71
l01DE = "Irish"
Case 72
l01DE = "Amerillo"
Case 73
l01DE = "Ann"
Case 74
l01DE = "Roland"
Case 75
l01DE = "Martin"
Case 76
l01DE = "Grey"
Case 77
l01DE = "Moore"
Case 78
l01DE = "Hong"
Case 79
l01DE = "Ha"
Case 80
l01DE = "Jonna"
Case 81
l01DE = "Cumberledge"
Case 82
l01DE = "Freemen"
Case 83
l01DE = "Jill"
Case 84
l01DE = "Jon"
Case 85
l01DE = "Williams"
Case 86
l01DE = "Cummings"
Case 87
l01DE = "Jane"
Case 88
l01DE = "Jamison"
Case 89
l01DE = "Raymond"
Case 90
l01DE = "Lashley"
Case 91
l01DE = "Reed"
Case 92
l01DE = "Sadler"
Case 93
l01DE = "Brown"
Case 94
l01DE = "White"
Case 95
l01DE = "Taylor"
Case 96
l01DE = "Doyle"
Case 97
l01DE = "Black"
Case 98
l01DE = "Tyler"
Case 99
l01DE = "Jones"
End Select
ln.Text = l01DE
Randomize
l01E2 = Int(81 * Rnd)
Select Case l01E2
Case 0
l01E6 = "Hibbling Road"
Case 1
l01E6 = "Hanover Lane"
Case 2
l01E6 = "Cherry Drive"
Case 3
l01E6 = "Essex Avenue"
Case 4
l01E6 = "Monticello Blvd."
Case 5
l01E6 = "Dickerson Lane"
Case 6
l01E6 = "Commerce Avenue"
Case 7
l01E6 = "Dinwiddie Road"
Case 8
l01E6 = "Sherwood Park"
Case 9
l01E6 = "Hillsdale Park"
Case 10
l01E6 = "Walnut Street"
Case 11
l01E6 = "Wall Street"
Case 12
l01E6 = "Troy Avenue"
Case 13
l01E6 = "Graystone Road"
Case 14
l01E6 = "Bath Street"
Case 15
l01E6 = "Backlick Road"
Case 16
l01E6 = "Braddock Road"
Case 17
l01E6 = "Skyline Lane"
Case 18
l01E6 = "Quaker Pl."
Case 19
l01E6 = "Townsquare Lane"
Case 20
l01E6 = "Accotink Street"
Case 21
l01E6 = "Brookfield Street"
Case 22
l01E6 = "Brandon Lane"
Case 23
l01E6 = "Dana Avenue"
Case 24
l01E6 = "Rose Lane"
Case 25
l01E6 = "Telegraph Road"
Case 26
l01E6 = "Hillsdale Court"
Case 27
l01E6 = "Floyd Avenue"
Case 28
l01E6 = "Attentee Road"
Case 29
l01E6 = "Olive Road"
Case 30
l01E6 = "Pine Street"
Case 31
l01E6 = "Bluebird Lane"
Case 32
l01E6 = "Connecticut Street"
Case 33
l01E6 = "Leaf Lane"
Case 34
l01E6 = "River Road"
Case 35
l01E6 = "Pinewood Lane"
Case 36
l01E6 = "Church Road"
Case 37
l01E6 = "Gunstun Lane"
Case 38
l01E6 = "Martin Luther Drive"
Case 39
l01E6 = "Mendota Lane"
Case 40
l01E6 = "Peach Street"
Case 41
l01E6 = "Orchard Road"
Case 42
l01E6 = "Main Street"
Case 43
l01E6 = "Sherry Lane"
Case 44
l01E6 = "Keen Road"
Case 45
l01E6 = "Washington Street"
Case 46
l01E6 = "Woodbridge Lane"
Case 47
l01E6 = "Old Keen Mill Drive"
Case 48
l01E6 = "Rolling Road"
Case 49
l01E6 = "Steve Road"
Case 50
l01E6 = "Oakland Drive"
Case 51
l01E6 = "Dunsmore Road"
Case 52
l01E6 = "Cardinal Forest Lane"
Case 53
l01E6 = "Oriole Lane"
Case 54
l01E6 = "Jerome Street"
Case 55
l01E6 = "Raindrop Way"
Case 56
l01E6 = "Shropshire Court"
Case 57
l01E6 = "Craft Road"
Case 58
l01E6 = "Burdon Court"
Case 59
l01E6 = "Parklane Court"
Case 60
l01E6 = "Franconia Road"
Case 61
l01E6 = "Achilles Road"
Case 62
l01E6 = "Jansen Court"
Case 63
l01E6 = "Alban Sta Blvd."
Case 64
l01E6 = "Oak Grove Street"
Case 65
l01E6 = "Spelman Dr."
Case 66
l01E6 = "Julian Street"
Case 67
l01E6 = "Lexton Pl."
Case 68
l01E6 = "Ellingham Road"
Case 69
l01E6 = "Denton Court"
Case 70
l01E6 = "Huntington Road"
Case 71
l01E6 = "Stana Street"
Case 72
l01E6 = "Queensberry Street"
Case 73
l01E6 = "Euphrates Road"
Case 74
l01E6 = "Euclid Way"
Case 75
l01E6 = "Tremont Road"
Case 76
l01E6 = "Mount Eagle Drive"
Case 77
l01E6 = "Winter Blue Court"
Case 78
l01E6 = "Francis Court"
Case 79
l01E6 = "Lee Street"
Case 80
l01E6 = "Jefferson Road"
End Select
Randomize
l01EA = Int(8100 * Rnd) + 1000
address.Text = l01EA + " " + l01E6
Randomize
l01EE = Int(800 * Rnd) + 100
l01EE = Trim("-" + Trim(Str(l01EE) + "-"))
Randomize
l01F2 = Int(8000 * Rnd) + 1000
l01F2 = Trim(Str(l01F2))
Randomize
l01F6 = Int(32 * Rnd)
Select Case l01F6
Case 0
state.Text = "VA"
city.Text = "Springfield"
zipcode.Text = "22150"
phone = "703" + l01EE + l01F2
Case 1
state.Text = "MD"
city.Text = "Upper Marlboro"
zipcode.Text = "20773"
phone.Text = "301" + l01EE + l01F2
Case 2
state.Text = "MD"
city.Text = "Oxon Hill"
zipcode.Text = "20750"
phone.Text = "301" + l01EE + l01F2
Case 3
state.Text = "CA"
city.Text = "Sandiego"
zipcode.Text = "90064"
phone.Text = "619" + l01EE + l01F2
Case 4
state.Text = "IA"
city.Text = "Des Moines"
zipcode.Text = "51932"
phone.Text = "515" + l01EE + l01F2
Case 5
state.Text = "WA"
city.Text = "Olympia"
zipcode.Text = "98601"
phone.Text = "206" + l01EE + l01F2
Case 6
state.Text = "WI"
city.Text = "Green Bay"
zipcode.Text = "53082"
phone.Text = "414" + l01EE + l01F2
Case 7
state.Text = "NY"
city.Text = "Buffalo"
zipcode.Text = "14216"
phone.Text = "716" + l01EE + l01F2
Case 8
state.Text = "ID"
city.Text = "Idaho Falls"
zipcode.Text = "83235"
phone.Text = "208" + l01EE + l01F2
Case 9
state.Text = "NV"
city.Text = "Reno"
zipcode.Text = "89727"
phone.Text = "702" + l01EE + l01F2
Case 10
state.Text = "ID"
city.Text = "Boise"
zipcode.Text = "83706"
phone.Text = "208" + l01EE + l01F2
Case 11
state.Text = "ND"
city.Text = "Bismarck"
zipcode.Text = "58501"
phone.Text = "701" + l01EE + l01F2
Case 12
state.Text = "UT"
city.Text = "Salt Lake City"
zipcode.Text = "84105"
phone.Text = "801" + l01EE + l01F2
Case 13
state.Text = "NJ"
city.Text = "Atlantic City"
zipcode.Text = "08401"
phone.Text = "609" + l01EE + l01F2
Case 14
state.Text = "TX"
city.Text = "Dallas"
zipcode.Text = "75216"
phone.Text = "214" + l01EE + l01F2
Case 15
state.Text = "NE"
city.Text = "Lincoln"
zipcode.Text = "68512"
phone.Text = "402" + l01EE + l01F2
Case 16
state.Text = "RI"
city.Text = "Providence"
zipcode.Text = "01130"
phone.Text = "401" + l01EE + l01F2
Case 17
state.Text = "CA"
city.Text = "San Francisco"
zipcode.Text = "94133"
phone.Text = "415" + l01EE + l01F2
Case 18
state.Text = "DC"
city.Text = "Washington"
zipcode.Text = "20570"
phone.Text = "202" + l01EE + l01F2
Case 19
state.Text = "IA"
city.Text = "Sioux City"
zipcode.Text = "51102"
phone.Text = "712" + l01EE + l01F2
Case 20
state.Text = "AK"
city.Text = "Juneau"
zipcode.Text = "99823"
phone.Text = "907" + l01EE + l01F2
Case 21
state.Text = "AL"
city.Text = "Birmingham"
zipcode.Text = "35241"
phone.Text = "205" + l01EE + l01F2
Case 22
state.Text = "Ok"
city.Text = "Oklahoma City"
zipcode.Text = "73104"
phone.Text = "405" + l01EE + l01F2
Case 23
state.Text = "DE"
city.Text = "Dover"
zipcode.Text = "19901"
phone.Text = "302" + l01EE + l01F2
Case 24
state.Text = "MS"
city.Text = "Jackson"
zipcode.Text = "39214"
phone.Text = "601" + l01EE + l01F2
Case 25
state.Text = "CT"
city.Text = "Hartford"
zipcode.Text = "06101"
phone.Text = "203" + l01EE + l01F2
Case 26
state.Text = "IN"
city.Text = "South Bend"
zipcode.Text = "46631"
phone.Text = "219" + l01EE + l01F2
Case 27
state.Text = "Ok"
city.Text = "Tulsa"
zipcode.Text = "74110"
phone.Text = "918" + l01EE + l01F2
Case 28
state.Text = "Nh"
city.Text = "Concord"
zipcode.Text = "03300"
phone.Text = "603" + l01EE + l01F2
Case 29
state.Text = "VA"
city.Text = "Richmond"
zipcode.Text = "23223"
phone.Text = "804" + l01EE + l01F2
Case 30
state.Text = "NJ"
city.Text = "Trenton"
zipcode.Text = "08600"
phone.Text = "609" + l01EE + l01F2
Case 31
state.Text = "WV"
city.Text = "Charleston"
zipcode.Text = "25304"
phone.Text = "304" + l01EE + l01F2
End Select
 l0210 = 1
Randomize
l0214 = Mid(Prefix.Text, 1, 4)
If  Prefix.Text = "" Then
MsgBox "You must enter in a prefix"
Exit Sub
End If

L38D74:
l0218 = 0
For l021C = 1 To 16
m01FA(l021C) = 0
Next l021C
Randomize
l0220 = Rnd * 10000
If  Len(Prefix.Text) = 5 Then
l0220 = Rnd * 1000
l0224 = l0220
If  Int(Mid(Prefix.Text, 5, 1)) <= 0 Then
If  Len(Int(l0220)) < 3 Then
GoTo L38D74
End If
l0218 = 1
l0220 = Trim(Str(Mid(Prefix.Text, 5, 1))) + Trim(Str(l0220))
Else
l0218 = 0
l0220 = Trim(Str(Mid(Prefix.Text, 5, 1))) + Trim(Str(l0220))
End If
End If
If  Len(Prefix.Text) = 6 Then
l0220 = Rnd * 100
l0224 = l0220
If  Len(Int(Mid(Prefix.Text, 5, 2))) <= 1 Then
If  Len(Int(l0220)) < 2 Then
GoTo L38D74
End If
l0218 = 1
l0220 = Trim(Str(Mid(Prefix.Text, 5, 2))) + Trim(Str(l0220))
Else
l0218 = 0
l0220 = Trim(Str(Mid(Prefix.Text, 5, 2))) + Trim(Str(l0220))
End If
End If
If  Len(Prefix.Text) = 7 Then
l0220 = Rnd * 10
l0224 = l0220
If  Int(Mid(Prefix.Text, 5, 3)) <= 2 Then
If  Int(l0220) < 1 Then
GoTo L38D74
End If
l0218 = 1
l0220 = Trim(Str(Mid(Prefix.Text, 5, 3))) + Trim(Str(l0220))
Else
l0218 = 0
l0220 = Trim(Str(Mid(Prefix.Text, 5, 3))) + Trim(Str(l0220))
End If
End If
If  Len(Prefix.Text) = 8 Then
l0220 = Mid(Prefix.Text, 5, 4)
End If
If  Len(Prefix.Text) = 5 Then
If  l0218 = 1 Then
l0218 = 0
Else
If  Len(Int(l0224)) < 3 Then GoTo L38D74
End If
End If
If  Len(Prefix.Text) = 6 Then
If  l0218 = 1 Then
l0218 = 0
Else
If  Len(Int(l0224)) < 2 Then GoTo L38D74
End If
End If
If  Len(Prefix.Text) = 7 Then
If  l0218 = 1 Then
l0218 = 0
Else
If  Int(l0224) <= 0 Then GoTo L38D74
End If
End If
l0228 = Rnd * 10000
l022C = Rnd * 10000
l0230 = Int(l0220)
l0234 = Int(l0228)
l0238 = Int(l022C)
l023C = l0230 \ 1000
l0240 = l023C * 1000
m01FA(1) = l023C
l0244 = l0230 - l0240
l0248 = l0244 \ 100
l024C = l0248 * 100
m01FA(2) = l0248
l0250 = l0244 - l024C
l0254 = l0250 \ 10
l0258 = l0254 * 10
m01FA(3) = l0254
l025C = l0250 - l0258
m01FA(4) = l025C
l0260 = l0234 \ 1000
l0264 = l0260 * 1000
m01FA(5) = l0260
l0268 = l0234 - l0264
l026C = l0268 \ 100
l0270 = l026C * 100
m01FA(6) = l026C
l0274 = l0268 - l0270
l0278 = l0274 \ 10
l027C = l0278 * 10
m01FA(7) = l0278
l0280 = l0274 - l027C
m01FA(8) = l0280
l0284 = l0238 \ 1000
l0288 = l0284 * 1000
m01FA(9) = l0284
l028C = l0238 - l0288
l0290 = l028C \ 100
l0294 = l0290 * 100
m01FA(10) = l0290
l0298 = l028C - l0294
l029C = l0298 \ 10
l02A0 = l029C * 10
m01FA(11) = l029C
l02A4 = l0298 - l02A0
m01FA(12) = l02A4
l02A8 = l0214 \ 1000
l02AC = l02A8 * 1000
m01FA(13) = l02A8
l02B0 = l0214 - l02AC
l02B4 = l02B0 \ 100
l02B8 = l02B4 * 100
m01FA(14) = l02B4
l02BC = l02B0 - l02B8
l02C0 = l02BC \ 10
l02C4 = l02C0 * 10
m01FA(15) = l02C0
l02C8 = l02BC - l02C4
m01FA(16) = l02C8
For l02CC = 1 To 15 Step 2
m01FA(l02CC) = m01FA(l02CC) * 2
If  m01FA(l02CC) > 9 Then
m01FA(l02CC) = m01FA(l02CC) - 9
End If
Next l02CC
l02D0 = 0
For l02D4 = 1 To 16
l02D0 = l02D0 + m01FA(l02D4)
Next l02D4
l02D8 = l02D0
l02DC = l02D8 \ 10
l02E0 = l02DC * 10
l02E4 = l02D8 - l02E0
If  l02E4 = 0 Then
l02E8 = Str$(l0214)
l02EC = Str$(l0230)
l02F0 = Str$(l0234)
l02F4 = Str$(l0238)
l02F8 = Str$(l0210)
ccn.Text = Trim(l02E8) & "-" & Format$(l02EC, "0###") & "-" & Format$(l02F0, "0###") & "-" & Format$(l02F4, "0###")
Else
GoTo L38D74
End If
GoTo L397E6
L397E0:
L397E6:
End Sub

Sub Form_Load ()
Dim l02FC As Variant
Dim l0300 As String
Dim l0302 As Variant
Dim l0306 As Variant
Dim l030A As Variant
Dim l030E As Variant
Dim l0312 As Variant
extsub1050 fac.hWnd, -1, 0, 0, 0, 0, 3
Prefix.AddItem "45309810"
Prefix.AddItem "444296"
Prefix.AddItem "445501"
Prefix.AddItem "401901"
Prefix.AddItem "403719"
Prefix.AddItem "4719"
Prefix.AddItem "4783"
Prefix.AddItem "5398"
Prefix.AddItem "540021"
Prefix.AddItem "549990"
Prefix.AddItem "548998"
Prefix.AddItem "544625"
Prefix.AddItem "443101"
Prefix.AddItem "565409"
On Error Resume Next
Open "fate.ini" For Binary As #1
l02FC = LOF(1)
If  l02FC = 0 Then
Close #1
Exit Sub
End If
l0300$ = String(l02FC, 0)
Get #1, 1, l0300$
Close #1
l0302 = l0300$
l0306 = InStr(1, l0302, "CC NUMBER=", 1)
l030A = Mid(l0302, l0306 + 10, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
ccn.Text = l0312
l0306 = InStr(1, l0302, "FIRST NAME=", 1)
l030A = Mid(l0302, l0306 + 11, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
fn.Text = l0312
l0306 = InStr(1, l0302, "LAST NAME=", 1)
l030A = Mid(l0302, l0306 + 10, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
ln.Text = l0312
l0306 = InStr(1, l0302, "ADDRESS=", 1)
l030A = Mid(l0302, l0306 + 8, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
address.Text = l0312
l0306 = InStr(1, l0302, "CITY=", 1)
l030A = Mid(l0302, l0306 + 5, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
city.Text = l0312
l0306 = InStr(1, l0302, "STATE=", 1)
l030A = Mid(l0302, l0306 + 6, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
state.Text = l0312
l0306 = InStr(1, l0302, "ZIPCODE=", 1)
l030A = Mid(l0302, l0306 + 8, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
zipcode.Text = l0312
l0306 = InStr(1, l0302, "PHONE NUMBER=", 1)
l030A = Mid(l0302, l0306 + 13, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
phone.Text = l0312
l0306 = InStr(1, l0302, "CERTIFICATE NUMBER=", 1)
l030A = Mid(l0302, l0306 + 19, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
cn.Text = l0312
l0306 = InStr(1, l0302, "CERTIFICATE PASSWORD=", 1)
l030A = Mid(l0302, l0306 + 21, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
cp.Text = l0312
l0306 = InStr(1, l0302, "EXPIRATION=", 1)
l030A = Mid(l0302, l0306 + 11, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
expiration.Text = l0312
l0306 = InStr(1, l0302, "BANK NAME=", 1)
l030A = Mid(l0302, l0306 + 10, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
bank.Text = l0312
l0306 = InStr(1, l0302, "PREFIX=", 1)
l030A = Mid(l0302, l0306 + 7, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
Prefix.Text = l0312
l0306 = InStr(1, l0302, "SECONDARY PREFIX=", 1)
l030A = Mid(l0302, l0306 + 17, Len(l0302))
l030E = InStr(l030A, Chr(13))
l0312 = Mid(l030A, 1, l030E - 1)
Close #1
If  Err Then sub3A0
End Sub

Sub Form_Unload (p1 As Integer)
Dim l0318 As String
Dim l031A As Variant
Dim l031E As Variant
Dim l0322 As Variant
Dim l0326 As String
Dim l0328 As String
Dim l032A As String
On Error Resume Next
SetAttr "fate.ini", 0
Open "fate.ini" For Binary As #1
l0318$ = String(LOF(1), 0)
Get #1, 1, l0318$
Close #1
l031A = InStr(1, l0318$, "[FAC]", 1)
l031E = Mid(l0318$, 1, l031A + 6)
l0322 = Chr(13) + Chr(10)
Open "fate.ini" For Output As #1
l0326$ = "Certificate Password=" + cp.Text + l0322 + "Expiration=" + expiration.Text + l0322 + "Bank Name=" + bank.Text + l0322 + "Prefix=" + Prefix.Text
l0328$ = "Zipcode=" + zipcode.Text + l0322 + "Phone Number=" + phone.Text + l0322 + "Certificate Number=" + cn + l0322 + l0326$
l032A$ = "Address=" + address.Text + l0322 + "City=" + city.Text + l0322 + "State=" + state.Text + l0322 + l0328$
l0318$ = "CC Number=" + ccn.Text + l0322 + "First Name=" + fn.Text + l0322 + "Last Name=" + ln.Text + l0322 + l032A$
Print #1, l031E + l0318$
Close #1
If  Err Then sub3A0
End Sub

Sub Prefix_Click ()
If  Prefix.Text = "45309810" Then
bank.Text = "Canadian Imperial"
End If
If  Prefix.Text = "444296" Then
bank.Text = "Bank of America"
End If
If  Prefix.Text = "401901" Then
bank.Text = "American Bank"
End If
If  Prefix.Text = "403719" Then
bank.Text = "Wells Fargo"
End If
If  Prefix.Text = "4719" Then
bank.Text = "Rocky Mountain"
End If
If  Prefix.Text = "4783" Then
bank.Text = "AT&T's Universal Card"
End If
If  Prefix.Text = "540021" Then
bank.Text = "Bank of India"
End If
If  Prefix.Text = "549990" Then
bank.Text = "Rocky Mountain Bankcard"
End If
If  Prefix.Text = "548998" Then
bank.Text = "First American Bank"
End If
If  Prefix.Text = "5398" Then
bank.Text = "AT&T Universal Card"
End If
If  Prefix.Text = "445501" Then
bank.Text = "Bank of America"
End If
End Sub

Sub sub9980 ()
End Sub
