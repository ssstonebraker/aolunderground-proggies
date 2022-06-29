' Module2
Option Explicit
Declare Function extfn1870 Lib "User" Alias "setparent" () As Integer
Declare Function extfn1448 Lib "311.dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn16E8 Lib "User" Alias "getwindow" () As Integer
Declare Function extfn1720 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1598 Lib "User" Alias "GetClassName" () As Integer
Declare Function extfn1528 Lib "User" Alias "findwindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfn1480 Lib "VBWFind.dll" Alias "FindChild" () As Integer
Declare Function extfn14F0 Lib "VBWFind.dll" Alias "FindCHildByTitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn14B8 Lib "VBWFind.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn15D0 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn16B0 Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1608 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1790 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4%) As Long
Declare Function extfn1800 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn17C8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn1678 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1640 Lib "User" Alias "GetNextWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1838 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn1560 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn1758 Lib "User" Alias "IsWindow" () As Integer
Declare Function extfn18A8 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn18E0 Lib "User" Alias "getwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1918 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1950 Lib "User" Alias "findwindow" () As Integer
Declare Function extfn1988 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn19C0 Lib "311.dll" Alias "AOLGetCombo" () As Integer
Declare Function extfn19F8 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Const c0092 = 2 ' &H2%
Const c00A6 = 513 ' &H201%
Const c00A8 = 514 ' &H202%
Const c00DE = 1040 ' &H410%
Const c0122 = 12 ' &HC%
Const c0124 = 513 ' &H201%
Const c0128 = 514 ' &H202%
Const c0156 = 273 ' &H111%

Function fn1A30 (p004C As String) As Integer
fn1A30 = True
If  p004C$ = "" Or InStr(UCase(p004C$), "A") Or InStr(UCase(p004C$), "B") Or InStr(UCase(p004C$), "C") Or InStr(UCase(p004C$), "D") Or InStr(UCase(p004C$), "E") Or InStr(UCase(p004C$), "F") Or InStr(UCase(p004C$), "G") Or InStr(UCase(p004C$), "H") Or InStr(UCase(p004C$), "I") Or InStr(UCase(p004C$), "J") Or InStr(UCase(p004C$), "K") Or InStr(UCase(p004C$), "L") Or InStr(UCase(p004C$), "M") Or InStr(UCase(p004C$), "N") Or InStr(UCase(p004C$), "O") Or InStr(UCase(p004C$), "P") Or InStr(UCase(p004C$), "Q") Or InStr(UCase(p004C$), "R") Or InStr(UCase(p004C$), "S") Or InStr(UCase(p004C$), "T") Or InStr(UCase(p004C$), "U") Or InStr(UCase(p004C$), "V") Or InStr(UCase(p004C$), "W") Or InStr(UCase(p004C$), "X") Or InStr(UCase(p004C$), "Y") Or InStr(UCase(p004C$), "Z") Or InStr(UCase(p004C$), " ") Then
   fn1A30 = False
End If
If  InStr(p004C$, "~") Or InStr(p004C$, "`") Or InStr(p004C$, "!") Or InStr(p004C$, "@") Or InStr(p004C$, "#") Or InStr(p004C$, "$") Or InStr(p004C$, "%") Or InStr(p004C$, "^") Or InStr(p004C$, "&") Or InStr(p004C$, "*") Or InStr(p004C$, "(") Or InStr(p004C$, ")") Or InStr(p004C$, "-") Or InStr(p004C$, "_") Or InStr(p004C$, "+") Or InStr(p004C$, "=") Or InStr(p004C$, "{") Or InStr(p004C$, "}") Or InStr(p004C$, "[") Or InStr(p004C$, "]") Or InStr(p004C$, "|") Or InStr(p004C$, "\") Or InStr(p004C$, ":") Or InStr(p004C$, ";") Or InStr(p004C$, "'") Or InStr(p004C$, Chr(34)) Or InStr(p004C$, "<") Or InStr(p004C$, ">") Or InStr(p004C$, ",") Or InStr(p004C$, ".") Or InStr(p004C$, "?") Or InStr(p004C$, "/") Then
   fn1A30 = False
End If
End Function

Function fn1A68 () As Integer
Dim l0050 As Integer
Dim l0052 As Integer
Dim l0054 As Integer
l0050% = extfn1528("AOL Frame25", 0&)
l0052% = extfn14B8(l0050%, "_AOL_EDIT")
l0054% = extfn14B8(l0050%, "_AOL_VIEW")
If  l0054% And l0052% Then
   fn1A68 = True
Else
   fn1A68 = False
End If
End Function

Function fn1AA0 () As Integer
Dim l0058 As Integer
Dim l005A As Integer
l0058% = extfn1528("AOL Frame25", 0&)
l005A% = extfn14F0(l0058%, "Welcome, ")
If  l005A% Then
   fn1AA0 = True
Else
   fn1AA0 = False
End If
End Function

Function fn1AD8 (p005E As String) As String
Dim l0060 As String
Dim l0062 As Integer
For l0062 = 1 To Len(p005E$)
    If  (Mid$(p005E$, l0062, 1) <> Chr$(0)) Then
        l0060$ = l0060$ + Mid$(p005E$, l0062, 1)
    Else
        GoTo L5C74
    End If
Next
L5C74:
fn1AD8 = l0060$
End Function

Function fn1B10 () As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0070 As Integer
Dim l0074 As String * 255
Dim l0076 As Integer
l0070 = extfn18E0(fn1B48(), l0068)
Do
  l0076% = extfn1918(l0070, l0074$, 254)
  If  InStr(l0074$, "AOL_Combobox") Then
     fn1B10 = l0070
     Exit Function
  End If
  l0070 = extfn18E0(l0070, l006C)
  DoEvents
Loop While l0070 <> 0
fn1B10 = 0
End Function

Function fn1B48 () As Variant
Dim l007C As Integer
Dim l007E As Integer
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Integer
Dim l0086 As Integer
Dim l0088 As Integer
Dim l008A As Integer
Dim l008E As String * 255
Dim l0090 As Integer
' Const c0092 = 2 ' &H2%
l0084 = 0
l0086 = 0
l0088 = 0
l008A = 0
Do
  l0090% = extfn1918(l007C%, l008E$, 254)
    If  InStr(l008E$, "MDIClient") Then
      l007E% = l007C%
    End If
  DoEvents
Loop While l007C% <> 0
Do
  Do
    l0090% = extfn1918(l0082, l008E$, 254)
    If  InStr(l008E$, "AOL_Edit") Then
       l0084 = l0084 + 1
    ElseIf InStr(l008E$, "AOL_Combobox") Then
       l0086 = l0086 + 1
    ElseIf InStr(l008E$, "AOL_Glyph") Then
       l0088 = l0088 + 1
    ElseIf InStr(l008E$, "AOL_Icon") Then
       l008A = l008A + 1
    End If
    l0082 = extfn18E0(l0082, c0092)
    DoEvents
  Loop While l0082 <> 0
  If  l0084 = 1 And l008A = 3 And l0088 = 2 And l0086 = 1 Then
     fn1B48 = l0080
     Exit Function
  End If
Loop While l0080 <> 0
fn1B48 = 0
End Function

Function fn1B80 (p0096 As String) As String
Dim l0098 As Integer
Dim l009A As String
For l0098 = 1 To Len(p0096$)
    If  (Mid$(p0096$, l0098, 1) <> Chr$(0)) Then
        l009A$ = l009A$ + Mid$(p0096$, l0098, 1)
    Else
        GoTo L6008
    End If
Next
L6008:
fn1B80 = l009A$
End Function

Sub sub1BB8 (p009C As Integer)
Dim l00A2 As Integer
Dim l00A4 As Integer
' Const c00A6 = 513 ' &H201%
' Const c00A8 = 514 ' &H202%
l00A2% = extfn1528("AOL Frame25", 0&)
l00A4 = extfn17C8&(p009C%, c00A6, 0, 0&)
l00A4 = extfn17C8&(p009C%, c00A8, 0, 0&)
End Sub

Sub sub1BF0 (p00AC As Form)
Dim l00AE As String
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00BC As String
Dim l00BE As Variant
Dim l00C8 As Integer
Dim l00CC As Integer
Dim l00CE As Variant
Dim l00D4 As Variant
Dim l00DA As Integer
Dim l00DC As String
' Const c00DE = 1040 ' &H410%
Call sub11D8(l00AE$)

L60BA:
DoEvents
l00C8% = extfn288("AOL Frame25", 0&)
l00CC% = extfn250(l00C8%, "Who's Here")
l00CE = 23
If  l00CC% = 0 Then l00CC% = extfn250(l00C8%, "Rooms") And l00CE = 48
l00D4 = extfn218(l00CC%, "_AOL_Listbox")
For l00DA% = 0 To l00CE
 l00DC$ = String$(256, " ")
 l00BE = extfn1448(l00DA%, l00DC$) & l00BC$
 If  Len(Trim$(l00DC$)) <= 1 Then GoTo L640E
 l00DC$ = Left$(Trim$(l00DC$), Len(Trim(l00DC$)) - 1)
l00B0 = extfn1800(p00AC.List1.hWnd, c00DE, 0, l00DC$)
 l00B4 = extfn1800(p00AC.List1.hWnd, c00DE, l00B0, l00DC$)
  If  l00B4 >= 0 Then GoTo L63F4
   If  Left(UCase(l00DC$), 5) = "GUIDE" Or Left(UCase(l00DC$), 3) = "TOS" Or Left(UCase(l00DC$), 8) = "CATWATCH" Or Right(UCase(l00DC$), 4) = "HOST" Then
      p00AC.List1.AddItem l00DC$
GoTo L63F4
   Else
      p00AC.List1.AddItem l00DC$
   End If
      GoTo L63FA
L63F4:
L63FA:
Next l00DA%

L640E:
End Sub

Sub sub11D8 (p00E0 As String)
Dim l00E2 As String
Dim l00E4 As String
Dim l00E6 As Integer
Dim l00E8 As Integer
Dim l00EA As String
Dim l00EC As Integer
Dim l00EE As String
Dim l00F0 As String
On Error GoTo L65C4
l00E6% = extfn1528("AOL Frame25", 0&) Or extfn1528("America  Online", 0&)
If  l00E6% = 0 Then MsgBox "AOL Dissapeared", 64, "¡Queso! 2.0": Exit Sub
l00E8% = extfn14F0(l00E6%, "Welcome, ")
If  l00E8% = 0 Then Exit Sub
l00EA$ = String$(255, 0)
l00EC% = extfn1720(l00E8%, l00EA$, 255)
l00EE$ = fn1AD8(l00EA$)
l00F0 = Right$(l00EE$, Len(l00EE) - 9)
l00E4$ = Left$(l00F0, Len(l00F0) - 1)
l00E2$ = l00E4$
p00E0$ = l00E2$
Exit Sub

L65C4:
p00E0$ = ""
Exit Sub
End Sub

Sub sub1C28 (p00F2 As Variant)
Dim l00F6 As Integer
Dim l00F8 As Integer
Dim l00FA As Variant
Dim l00FE As Integer
Dim l0100 As Integer
Dim l0104 As Variant
Dim l0108 As Variant
Dim l010C As Variant
Dim l0110 As Variant
Dim l0114 As Variant
Dim l0118 As Variant
Dim l011C As Variant
Dim l0120 As Integer
' Const c0122 = 12 ' &HC%
' Const c0124 = 513 ' &H201%
' Const c0128 = 514 ' &H202%
l00F6% = extfn1528("AOL Frame25", 0&)
l00F8% = extfn14F0(l00F6%, "Rooms")
If  l00F8% = 0 Then
l0104 = extfn288("AOL FRAME25", 0&)
l0108 = extfn218(l0104, "_aol_glyph")
l010C = extfn1678(l0108)
l0110 = extfn218(l010C, "_aol_listbox")
l0114 = extfn1678(l0110)
l0118 = extfn218(l0114, "_aol_icon")
l011C = extfn1678(l0118)
l00F8% = extfn218(l011C, "_AOL_edit")
End If
If  l00F8% <> 0 Or l0120% <> 0 Then
   If  l00F8% Then l00F8% = extfn1678(l00F8%)
   l00FE% = extfn14B8(l00F8%, "_AOL_EDIT")
   l00FA = extfn1800(l00FE%, c0122, 0, p00F2)
   l0100% = extfn1640(l00FE%, c0092)
   DoEvents
   l0104 = extfn6E8(l0100%, c0124, 0, 0)
   l0104 = extfn6E8(l0100%, c0128, 0, 0)
   DoEvents
End If
End Sub

Sub sub1C60 (p012A As Variant)
Dim l012E As Integer
Dim l0130 As Integer
Dim l0132 As Variant
Dim l0136 As Integer
Dim l0138 As Integer
Dim l013A As String
Dim l013C As Variant
Dim l0140 As Variant
l012E% = extfn1528("AOL Frame25", 0&)
l0130% = extfn14F0(l012E%, "Rooms")
If  l0130% = 0 Then
l0140 = l0140 + "CHATROOM NOT TO BE FOUND"
Exit Sub
Else
   l0138% = extfn1678(l0130%)
   l0136% = extfn14B8(l0138%, "_AOL_VIEW")
   l013C = extfn17C8(l0136%, &HE, 0, 0&)
   l013A$ = String$(l013C, 0)
   l0132 = extfn1800(l0136%, &HD, l013C, l013A$)
   p012A = Left$(l013A$, l0132)
End If
End Sub

Sub sub1C98 (p0144 As Integer, p0146 As Integer)
Dim l0148 As Integer
Dim l014A As Integer
Dim l014C As Integer
Dim l014E As Integer
Dim l0150 As Integer
Dim l0152 As Long
' Const c0156 = 273 ' &H111%
    l0150% = extfn1528("AOL Frame25", 0&)
    l014A = extfn15D0%(l0150%)
    l014C = extfn16B0%(l014A, p0144)
    l014E% = extfn1608%(l014C, p0146)
    l0152 = CLng(0) * &H10000& Or gv00CC
    l0148% = extfn17C8(l0150%, c0156, l014E, 0&)
End Sub

Sub sub1CD0 (p0158 As String, p015A As String)
Dim l015C As Integer
Dim l015E As Integer
Dim l0160 As Integer
Dim l0162 As Integer
Dim l0164 As Integer
Dim l0166 As Integer
Dim l0168 As Integer
Dim l016A As Variant
' Const c016E = 16 ' &H10%
l015C% = extfn1528("AOL Frame25", 0&)
Call sub1C98(4, 3)
Do
DoEvents
l015E% = extfn14F0(l015C%, "Send Instant Message")
Loop Until l015E%
l0160% = extfn14B8(l015E%, "_AOL_EDIT")
l0162% = extfn1640(l0160%, c0092)
l0164% = extfn14F0(l015E%, "Send")
l016A = extfn1800(l0160%, c0122, 0, p0158)
l016A = extfn1800(l0162%, c0122, 0, p015A)
sub1BB8 (l0164%)
Do
DoEvents
l0166% = extfn1528("#32770", "America Online")
l015E% = extfn14F0(l015C%, "Send Instant Message")
Loop Until l0166% Or l015E% = False
If  l015E% = False Then Exit Sub
If  l0166% Then
   l0168% = extfn14F0(l0166%, "OK")
   sub1BB8 (l0168%)
   l016A = extfn1790(l015E%, c016E, 0, 0)
End If
End Sub
