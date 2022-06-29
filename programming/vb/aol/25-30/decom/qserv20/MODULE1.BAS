' Module1
Option Explicit
Declare Function extfn870 Lib "GDI" Alias "SetTextColor" () As Long
Declare Function extfn170 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn138 Lib "im30.dll" Alias "GetAPIText" (ByVal p1%) As String
Declare Function extfn838 Lib "User" Alias "setparent" () As Integer
Declare Function extfn410 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn3A0 Lib "User" Alias "GetMenuItemCount" (ByVal p1%) As Integer
Declare Function extfn330 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn3D8 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn4B8 Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn368 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn758 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn7C8 Lib "User" Alias "SetCursor" () As Integer
Declare Function extfn918 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn8E0 Lib "User" Alias "SetWindowText" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn800 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn480 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn6B0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn6E8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn720 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn288 Lib "User" Alias "findwindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn2F8 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn560 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn2C0 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn598 Lib "User" Alias "GetWindowTextLength" (ByVal p1%) As Integer
Declare Function extfn5D0 Lib "User" Alias "GetWindowWord" () As Integer
Declare Function extfn608 Lib "Kernel" Alias "lstrlen" (ByVal p1&) As Integer
Declare Function extfn678 Lib "VBMSG.VBX" Alias "PTGETSTRINGFROMADDRESS" (ByVal p1&, ByVal p2%) As String
Declare Function extfn218 Lib "VBWFind.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1A8 Lib "User" Alias "enumchildwindows" () As Integer
Declare Function extfn250 Lib "VBWFind.dll" Alias "FindCHildByTitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn448 Lib "User" Alias "GetNextWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn790 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn1E0 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfnC8 Lib "311.dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn100 Lib "311.dll" Alias "AOLGetCombo" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn4F0 Lib "User" Alias "GetTopWindow" (ByVal p1%) As Integer
Declare Function extfn528 Lib "User" Alias "getwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Sub extsub8A8 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Sub extsub640 Lib "User" Alias "MoveWindow" ()
Declare Function extfn950 Lib "VBMSG.VBX" Alias "PTGETSTRINGFROMADDRESS" () As String
Const c009C = 1024 ' &H400%
Const c00C6 = 273 ' &H111%
Const c00DC = 5 ' &H5%
Const c00DE = 2 ' &H2%
Const c00E0 = 12 ' &HC%
Const c00E2 = 258 ' &H102%
Const c0232 = 1036 ' &H40C%
Const c038A = 1040 ' &H410%
Const c0396 = 513 ' &H201%
Const c0398 = 514 ' &H202%
Const c03A6 = 16 ' &H10%
Const c044E =-1 ' &HFFFF%
Const c0450 = 16 ' &H10%
Const c0452 = 64 ' &H40%
Const c04C0 = 513 ' &H201%
Const c04C2 = 514 ' &H202%

Function fn988 (p0092 As Integer, p0094 As Integer) As String
Dim l0098 As String * 255
Dim l009A As Integer
' Const c009C = 1024 ' &H400%
    l009A% = extfn410(p0092%, p0094%, l0098$, 128, c009C)
    fn988 = Left$(l0098$, l009A%)
End Function

Function fn9C0 (ByVal p00A2 As String, p00A4 As String) As Variant
Dim l00A6 As Integer
Dim l00A8 As Integer
Dim l00AA As Integer
Dim l00AC As Integer
Dim l00AE As Long
Dim l00B0 As Integer
Dim l00B2 As Variant
Dim l00B6 As Integer
Dim l00B8 As String
Dim l00BA As Integer
Dim l00BC As Integer
Dim l00BE As String
Dim l00C0 As Variant
' Const c00C6 = 273 ' &H111%
    l00B0% = extfn288("AOL Frame25", 0&)
    l00A6% = extfn368(l00B0%)
    If  l00A6% < 0 Then
        l00B2 = MsgBox("Menu Not Found", 16, "Menu Error!")
    Else
        l00B6% = extfn3A0(l00A6%)
        For l00A8 = 0 To l00B6% - 1
            l00B8$ = fn988(l00A6%, l00A8)
            If  l00B8$ = p00A2 Then
                l00BA% = extfn4B8(l00A6%, l00A8)
                l00BC% = extfn3A0(l00BA%)
                For l00AA = 0 To l00BC% - 1
                    l00BE$ = fn988(l00BA%, l00AA)
                    If  l00BE$ = p00A4 Then
                        l00AC% = extfn3D8(l00BA%, l00AA)
                        l00AE = CLng(0) * &H10000& Or gv0014
                        l00C0 = extfn6E8(extfn288("AOL Frame25", 0&), c00C6, l00AC, 0&)
                        Exit Function
                    End If
                Next l00AA
            End If
        Next l00A8
    End If
End Function

Function fn9F8 (ByVal pv00CA As Integer, p00CC As String) As Integer
Dim l00CE As Integer
Dim l00D0 As Integer
Dim l00D4 As String * 255
Dim l00D6 As Integer
Dim l00D8 As Integer
Dim l00DA As Integer
' Const c00DC = 5 ' &H5%
' Const c00DE = 2 ' &H2%
' Const c00E0 = 12 ' &HC%
' Const c00E2 = 258 ' &H102%
If  pv00CA <> 0 Then
  l00CE = extfn528(pv00CA, c00DC)
  Do
    l00D0% = extfn2F8(l00CE, l00D4$, 254)
    If  InStr(l00D4$, "AOL_Edit") Then
        l00D6 = l00CE
    ElseIf InStr(l00D4$, "AOL_View") Then
        l00D8 = l00CE
    ElseIf InStr(l00D4$, "AOL_Button") Then
        l00DA = l00CE
    End If
    l00CE = extfn528(l00CE, c00DE)
    DoEvents
  Loop While l00CE <> 0
  If  l00D6 And l00D8 And l00DA Then
    l00D0 = extfn720(l00D6%, c00E0, 0, p00CC)
    l00D0 = extfn6E8(l00D6%, c00E2, 13, 0)
    fn9F8 = True
  End If
End If
End Function

Function fnA30 (p00E8 As ListBox) As String
Dim l00EA As Variant
Dim l00EE As Variant
On Error GoTo LD76
For l00EA = 1 To p00E8.ListCount
    l00EE = l00EE & p00E8.List(l00EA - 1) & ","
Next l00EA
fnA30 = l00EE
Exit Function
LD76:
fnA30 = ""
End Function

Function fnA68 (p00F4 As String) As String
Dim l00F6 As Integer
Dim l00F8 As String
For l00F6 = 1 To Len(p00F4$)
    If  (Mid$(p00F4$, l00F6, 1) <> Chr$(0)) Then
        l00F8 = l00F8 + Mid$(p00F4$, l00F6, 1)
    Else
        GoTo LE02
    End If
Next
LE02:
fnA68 = l00F8$
End Function

Function fnAA0 (p00FC As Variant) As String
Dim l0102 As String * 255
Dim l0104 As Integer
l0104% = extfn560(p00FC, l0102, 254)
fnAA0 = l0102
End Function

Function fnAD8 (p010A As String) As Variant
Dim l010C As Variant
Dim l0110 As Variant
l010C = fnD08()
l0110 = extfn8E0(l010C, p010A)
fnAD8 = l0110
End Function

Function fnB10 () As Variant
fnB10 = fnAA0(fnD08())
End Function

Function fnB48 () As Variant
Dim l011C As Variant
Dim l0120 As Integer
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126 As Integer
Dim l0128 As Integer
Dim l012C As String * 255
Dim l012E As Integer
Dim l0130 As Variant
Dim l0134 As Variant
Dim l0138 As Variant
Dim l013C As Variant
  l011C = extfn288("AOL Frame25", "America  Online")
Do
  l0124 = extfn528(l0122, c00DC)
  Do
    l012E% = extfn2F8(l0124, l012C$, 254)
    If  l0130 = False And InStr(l012C$, "AOL_Edit") Then
        l0128 = l0128 + 1
        l0130 = True
    End If
    If  l0134 = False And InStr(l012C$, "AOL_ListBox") Then
        l0128 = l0128 + 1
        l0134 = True
    End If
    If  l0138 = False And InStr(l012C$, "AOL_View") Then
        l0128 = l0128 + 1
        l0138 = True
    End If
    If  l013C < 5 And InStr(l012C$, "AOL_Icon") Then
        l0128 = l0128 + 1
        l013C = l013C + 1
    End If
    If  extfn250(l0120%, "People in Room:") Then l0128 = l0128 + 1
    If  extfn250(l0120%, "Who's Here") Then l0128 = l0128 + 1
    If  extfn250(l0120%, "Center Stage") Then l0128 = l0128 + 1
    If  extfn250(l0120%, "PC Studio") Then l0128 = l0128 + 1
    If  extfn250(l0120%, "Chat Preferences") Then l0128 = l0128 + 1
    If  extfn250(l0120%, "Parental Control") Then l0128 = l0128 + 1
    If  extfn250(l0120%, "Send") Then l0128 = l0128 + 1
    l0124 = extfn528(l0124, c00DE)
    DoEvents
  Loop While l0124 <> 0
  If  l0128 = 15 Then l0126 = l0122: Exit Do
  MsgBox "You must be in a chat room!!! geez!"
  l0122 = extfn528(l0122, c00DE)
Loop While l0122 <> 0
fnB48 = l0126
End Function

Function fnB80 (p0144 As Integer, p0146 As String) As Variant
Dim l0148 As Integer
Dim l014A As Integer
Dim l014C As Integer
Dim l014E As Variant
l014A% = p0144
If  l014A% = 0 Then
    l014E = 0
    Exit Function
End If
l0148 = extfn528(l014C%, c00DC)
Do
  If  InStr(fnAA0(l0148), p0146) <> 0 Then
      fnB80 = l0148
      Exit Do
  End If
  l0148 = extfn528(l0148, c00DE)
Loop While l0148 <> 0
End Function

Function fnBB8 (p0156 As String, p0158 As String) As Variant
Dim l015A As Variant
Dim l015E As Variant
Dim l0162 As Variant
l015A = Len(p0158$)
For l015E = 1 To Len(p0156$)
 l0162 = Asc(Mid$(p0158$, (l015E Mod l015A) - l015A * ((l015E Mod l015A) = 0), 1)) + 81
 Mid$(p0156$, l015E, 1) = Chr$(Asc(Mid$(p0156$, l015E, 1)) Xor l0162)
Next l015E
fnBB8 = p0156$
End Function

Function fnBF0 (p0168 As String) As Integer
Dim l016A As Integer
Dim l016C As Integer
Dim l016E As Integer
Dim l0170 As Integer
Dim l0174 As String * 255
l016E% = extfn528(extfn288("AOL Frame25", 0&), c00DC)
Do
  l016A% = extfn2F8(l016E%, l0174$, 254)
  If  InStr(l0174$, "MDIClient") Then l0170% = l016E%
  l016E% = extfn528(l016E%, c00DE)
  DoEvents
Loop While l016E% <> 0
If  p0168 = "MDIClient" Then fnBF0 = l0170%
l016C = extfn528(l0170%, c00DC)
Do
  If  InStr(fnAA0(l016C), p0168) <> 0 Then
      fnBF0 = l016C
      Exit Do
  End If
  l016C = extfn528(l016C, c00DE)
Loop While l016C <> 0
End Function

Function fnC28 () As Integer
Dim l0178 As Integer
Dim l017A As Integer
Dim l017C As Integer
Dim l017E As Integer
Dim l0180 As Integer
Dim l0182 As Integer
Dim l0186 As String * 255
Dim l0188 As Integer
l0178% = extfn528(extfn288("AOL Frame25", 0&), c00DC)
Do
  l0188% = extfn2F8(l0178%, l0186$, 254)
    If  InStr(l0186$, "MDIClient") Then
      l017A% = l0178%
    End If
  l0178% = extfn528(l0178%, c00DE)
  DoEvents
Loop While l0178% <> 0
l017C = extfn528(l017A%, c00DC)
Do
  l017E = extfn528(l017C, c00DC)
  Do
    l0188% = extfn2F8(l017E, l0186$, 254)
    If  InStr(l0186$, "AOL_Edit") Then l0182 = l0182 + 1
    If  InStr(l0186$, "AOL_View") Then l0182 = l0182 + 1
    If  InStr(l0186$, "AOL_Listbox") Then l0182 = l0182 + 1
    If  InStr(l0186$, "AOL_Icon") Then l0182 = l0182 + 1
    If  extfn250(l017C, "Send") Then l0182 = l0182 + 1
    l017E = extfn528(l017E, c00DE)
    DoEvents
  Loop While l017E <> 0
  If  l0182 = 5 Then l0180 = l017C: Exit Do
  l017C = extfn528(l017C, c00DE)
Loop While l017C <> 0
fnC28 = l0180
End Function

Function fnC60 () As Variant
Dim l018E As Variant
Dim l0192 As Variant
l018E = extfn288("AOL Frame25", "America  Online")
l0192 = extfn250(l018E, "Who's Here")
If  l0192 = 0 Then fnC60 = 0: Exit Function
fnC60 = extfn480(l0192)
End Function

Function fnC98 () As Variant
fnC98 = extfn288("AOL Frame25", "America  Online")
End Function

Function fnCD0 () As Variant
Dim l019E As Integer
Dim l01A0 As Variant
Dim l01A4 As Variant
Dim l01AA As String * 100
Dim l01AC As Variant
l01A0 = extfn288("aol frame25", "america  online")
l01A4 = extfn4F0(l01A0)
l01A4 = extfn4F0(l01A4)
l019E = extfn598(l01A4)
l01AC = extfn560(l01A4, l01AA, l019E)
End Function

Function fnD08 () As Variant
Dim l01B4 As Variant
Dim l01B8 As Variant
Dim l01BC As Variant
Dim l01C0 As Variant
Dim l01C4 As Variant
Dim l01C8 As Variant
Dim l01CC As Variant
Dim l01D0 As Variant
l01B4 = extfn288("AOL Frame25", 0&)
l01B8 = extfn528(l01B4, c00DC)
l01BC = extfn250(l01B4, "List" & Chr(13) & "Rooms")
l01C0 = extfn250(l01B4, "Center" & Chr(13) & "Stage")
l01C4 = extfn250(l01B4, "PC" & Chr(13) & "Studio")
l01C8 = extfn250(l01B4, "Chat" & Chr(13) & "Preferences")
l01CC = extfn250(l01B4, "Parental" & Chr(13) & "Control")
l01D0 = extfn218(l01B4, "_AOL_LISTBOX")
If  l01BC <> 0 And extfn480(l01BC) = extfn480(l01C0) And extfn480(l01C0) = extfn480(l01C4) And extfn480(l01C4) = extfn480(l01C8) And extfn480(l01C8) = extfn480(l01CC) And extfn480(l01CC) = extfn480(l01D0) Then fnD08 = extfn480(l01BC): Exit Function
fnD08 = 0
End Function

Function fnD40 () As String
Dim l01D6 As Integer
Dim l01D8 As Integer
Dim l01DA As Integer
Dim l01DC As Integer
Dim l01DE As String
Dim l01E0 As Variant
Dim l01E4 As Variant
l01D6% = extfn288("AOL Frame25", "America  Online")
l01D8% = extfn528(l01D6%, 5)
l01DA% = extfn218(l01D8%, "MDIClient")
l01DC% = extfn528(l01DA%, 5)
On Error Resume Next
l01DE$ = String$(255, 0)
l01E0 = extfn528(l01DC%, 0)
l01E4 = extfn560(l01E0, l01DE$, 255)
If  Mid(l01DE$, 1, 23) = ">Instant Message From: " Then
fnD40 = l01DE$: Exit Function
End If
End Function

Function fnD78 () As Variant
Dim l01EC As Variant
Dim l01F0 As Variant
l01EC = extfn250(extfn288("AOL Frame25", "America  Online"), "Welcome, ")
If  l01EC = 0 Then Exit Function
l01F0 = fnAA0(l01EC)
fnD78 = Mid$(l01F0, 10, (InStr(l01F0, "!") - 10))
End Function

Function fnDB0 () As Integer
Dim l01F6 As Variant
l01F6 = extfn330()
fnDB0 = extfn480(l01F6)
End Function

Function fnDE8 () As Variant
Dim l01FE As Integer
Dim l0200 As String
Dim l0202 As Variant
Dim l0206 As Integer
Dim l0208 As Variant
Dim l020C As Variant
    l01FE = extfn6E8(l0202, &HE, 0, 0&)
    l0200 = Space$(l01FE)
    l0206 = extfn720(l0202, &HD, l01FE, l0200)
    If  l0208 = 1 Then
        l020C = Left$(l0200, l0206)
    Else
      l020C = Right$(Left$(l0200, l0206), 128)
    End If
End Function

Function fnE20 () As Integer
Dim l0212 As Variant
Dim l0216 As Variant
l0212 = extfn288("AOL Frame25", "America  Online")
l0216 = extfn250(l0212, "Welcome, ")
If  l0216 <> 0 Then fnE20 = True
If  l0216 = 0 Then
    fnE20 = False
    MsgBox "You must be signed on to AOL to do this.", 48, "You are not signed on!"
End If
End Function

Function fnE58 () As Variant
Dim l021E As Variant
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022A As Variant
Dim l022E As Variant
' Const c0232 = 1036 ' &H40C%
If  extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail") = False Then
Call subE90(3, 2)
Do
DoEvents
l021E = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
Loop Until l021E
l0222 = extfn218(l021E, "_AOL_Tree")
Do
DoEvents
l0222 = extfn218(l021E, "_AOL_Tree")
l0226 = extfn6B0(l0222, c0232, 0, 0&)
subEC8 (.5#)
l0222 = extfn218(l021E, "_AOL_Tree")
l022A = extfn6B0(l0222, c0232, 0, 0&)
l0222 = extfn218(l021E, "_AOL_Tree")
subEC8 (.5#)
l0222 = extfn218(l021E, "_AOL_Tree")
l022E = extfn6B0(l0222, c0232, 0, 0&)
Loop Until l0226 = l022A And l022A = l022E
Else
    l021E = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
    l0222 = extfn218(l021E, "_AOL_Tree")
    l0226 = extfn6B0(l0222, c0232, 0, 0&)
End If
fnE58 = l0226
MsgBox l0226
End Function

Function fnF00 () As String
Dim l0236 As Integer
Dim l0238 As Integer
Dim l023A As Integer
Dim l023E As String * 255
Dim l0240 As String
Dim l0242 As Integer
l0236% = extfn2C0()
l023A% = extfn528(l0236%, c00DC)
l0238% = extfn2F8(l0236%, l023E$, 254)
Do
  l0238% = extfn2F8(l023A%, l023E$, 254)
  If  InStr(UCase$(l023E$), "STATIC") Then
      l023A% = extfn528(l023A%, c00DE)
      l0242 = extfn598(l023A%)
      l0240 = String$(l0242, " ")
      l0238 = extfn560(l023A%, l0240, l0242)
      fnF00 = Trim$(l0240)
      Exit Do
  End If
  l023A% = extfn528(l023A%, c00DE)
  DoEvents
Loop While l023A% <> 0
End Function

Function fnF38 () As Variant
Dim l0248 As Variant
Dim l024C As Variant

L290C:
l0248 = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
If  l0248 Then
    fnF38 = True
    Exit Function
End If
l024C = MsgBox("Your mail box was not open.  Open it then press Retry, or to cancel press Cancel.", 5, "Open your New Mail box")
    If  l024C = 2 Then
        fnF38 = False
        Exit Function
    End If
GoTo L290C
End Function

Sub subF70 (p0250, p0254)
Dim l0258 As Integer
Dim l025A As String
Dim l025C As Integer
Dim l025E As Integer
Dim l0260 As Integer
Dim l0262 As Integer
Dim l0264 As Integer
Dim l0266 As Integer
Dim l0268 As Integer
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
Do
l0258% = extfn288("AOL Frame25", "America  Online")
l025A$ = "Send an Instant Message" + Chr$(9) + "Ctrl+I"
l025C% = fn9C0("Mem&bers", l025A$)
DoEvents
Loop Until extfn250(l0258%, "Available?") Or extfn250(l0258%, "Send Instant Message")
l025E% = extfn250(l0258%, "Send Instant Message")
l0260% = extfn218(l025E%, "_AOL_Edit")
l026C = extfn14B8(l025E%, "RICHCNTL")
If  l026C Then
        Call subFA8(l0260%, p0250)
        Call subFA8(l026C, p0254)
        l0270 = extfn218(l025E%, "_AOL_Icon")
    l0274 = extfn448(l0270, c00DE)
    l0278 = extfn448(l0274, c00DE)
    l027C = extfn448(l0278, c00DE)
    l0280 = extfn448(l027C, c00DE)
    l0284 = extfn448(l0280, c00DE)
    l0288 = extfn448(l0284, c00DE)
    l028C = extfn448(l0288, c00DE)
    l0290 = extfn448(l028C, c00DE)
    l0294 = extfn448(l0290, c00DE)
        subFE0 (l0294)
DoEvents
subEC8 (.2#)
Else
    l0262% = extfn448(l0260%, c00DE)
End If
Call subFA8(l0260%, p0250)
Call subFA8(l0262%, p0254)
l0264% = extfn250(l025E%, "Send")
subFE0 (l0264%)
DoEvents
subEC8 (.2#)
Do
l0258% = extfn288("AOL Frame25", "America  Online")
l025E% = extfn250(l0258%, "Send Instant Message")
subEC8 (.3#)
l0266% = extfn288("#32770", 0&)
l0268% = extfn250(l0266%, "OK")
subFE0 (l0268%)
Loop Until l0268% Or (l025E% = 0)
subFE0 (l0268%)
    sub1018 (l0266%)
        DoEvents
DoEvents
sub1018 (l025E%)
l0298 = "go"
End Sub

Sub sub1050 (p029C As Variant, p02A0)
Dim l02A4 As Variant
Dim l02A8 As Variant
l02A4 = extfn218(p029C, "_AOL_Edit")
Call subFA8(l02A4, p02A0)
DoEvents
l02A8 = extfn6B0(l02A4, c00E2, 13, 0)
End Sub

Sub sub1088 (p02AC, p02B0)
Dim l02B4 As Integer
Dim l02B6 As String
Dim l02B8 As Integer
Dim l02BA As Integer
Dim l02BC As Integer
Dim l02BE As Integer
Dim l02C0 As Integer
Dim l02C2 As Integer
Dim l02C4 As Integer
Dim l02C6 As Variant
Dim l02CA As Variant
Dim l02CE As Variant
Dim l02D2 As Variant
Dim l02D6 As Variant
Dim l02DA As Variant
Dim l02DE As Variant
Dim l02E2 As Variant
Dim l02E6 As Variant
Dim l02EA As Variant
Dim l02EE As Variant
l02B4% = extfn288("AOL Frame25", "America  Online")
l02B6$ = "Send an Instant Message" + Chr$(9) + "Ctrl+I"
l02B8% = fn9C0("Mem&bers", l02B6$)
Do
DoEvents
Loop Until extfn250(l02B4%, "Available?") Or extfn250(l02B4%, "Send Instant Message")
l02BA% = extfn250(l02B4%, "Send Instant Message")
l02BC% = extfn218(l02BA%, "_AOL_Edit")
l02C6 = extfn14B8(l02BA%, "RICHCNTL")
If  l02C6 Then
        Call subFA8(l02BC%, p02AC)
        Call subFA8(l02C6, p02B0)
        l02CA = extfn218(l02BA%, "_AOL_Icon")
    l02CE = extfn448(l02CA, c00DE)
    l02D2 = extfn448(l02CE, c00DE)
    l02D6 = extfn448(l02D2, c00DE)
    l02DA = extfn448(l02D6, c00DE)
    l02DE = extfn448(l02DA, c00DE)
    l02E2 = extfn448(l02DE, c00DE)
    l02E6 = extfn448(l02E2, c00DE)
    l02EA = extfn448(l02E6, c00DE)
    l02EE = extfn448(l02EA, c00DE)
        subFE0 (l02EE)
DoEvents
subEC8 (.9#)
Else
    l02BE% = extfn448(l02BC%, c00DE)
End If
Call subFA8(l02BC%, p02AC)
Call subFA8(l02BE%, p02B0)
l02C0% = extfn250(l02BA%, "Send")
subFE0 (l02C0%)
DoEvents
subEC8 (.9#)
Do
DoEvents
l02C2% = extfn288("#32770", 0&)
l02C4% = extfn250(l02C2%, "OK")
    DoEvents
Loop Until l02C2% <> 0 Or extfn250(l02B4%, "Available?") = 0
    sub1018 (l02C2%)
        DoEvents
DoEvents
sub1018 (l02BA%)
End Sub

Sub sub10C0 ()
Dim l02F2 As Integer
Dim l02F4 As Integer
Dim l02F6 As Variant
l02F2 = extfn330()
l02F4 = extfn330()
Do Until l02F4 <> l02F2
l02F6 = DoEvents()
l02F4 = extfn330()
Loop
End Sub

Sub sub10F8 ()
Dim l02FA As Variant
Dim l02FE As Variant
Dim l0302 As Variant
Dim l0306 As Variant
Dim l030A As Variant
Do
l02FA = extfn250(extfn288("AOL Frame25", "America  Online"), "New Mail")
subEC8 (.1#)
Loop Until l02FA <> 0
l02FE = extfn218(l02FA, "_AOL_Tree")
Do
DoEvents
l0302 = extfn6B0(l02FE, c0232, 0, 0&)
Call subEC8(.5#)
l0306 = extfn6B0(l02FE, c0232, 0, 0&)
Call subEC8(.5#)
l030A = extfn6B0(l02FE, c0232, 0, 0&)
Loop Until l0302 = l0306 And l0306 = l030A
    l0302 = extfn6B0(l02FE, c0232, 0, 0&)
End Sub

Sub sub1130 (p030E, p0312)
Dim l0316 As Integer
Dim l0318 As String
Dim l031A As Integer
Dim l031C As Integer
Dim l031E As Integer
Dim l0320 As Integer
Dim l0322 As Integer
Dim l0324 As Integer
Dim l0326 As Integer
Dim l0328 As Variant
Dim l032C As Variant
Dim l0330 As Variant
Dim l0334 As Variant
Dim l0338 As Variant
Dim l033C As Variant
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As Variant
Dim l0350 As Variant
l0316% = extfn288("AOL Frame25", "America  Online")
l0318$ = "Send an Instant Message" + Chr$(9) + "Ctrl+I"
l031A% = fn9C0("Mem&bers", l0318$)
Do
DoEvents
Loop Until extfn250(l0316%, "Available?") Or extfn250(l0316%, "Send Instant Message")
l031C% = extfn250(l0316%, "Send Instant Message")
l031E% = extfn218(l031C%, "_AOL_Edit")
l0328 = extfn14B8(l031C%, "RICHCNTL")
If  l0328 Then
        Call subFA8(l031E%, p030E)
        Call subFA8(l0328, p0312)
        l032C = extfn218(l031C%, "_AOL_Icon")
    l0330 = extfn448(l032C, c00DE)
    l0334 = extfn448(l0330, c00DE)
    l0338 = extfn448(l0334, c00DE)
    l033C = extfn448(l0338, c00DE)
    l0340 = extfn448(l033C, c00DE)
    l0344 = extfn448(l0340, c00DE)
    l0348 = extfn448(l0344, c00DE)
    l034C = extfn448(l0348, c00DE)
    l0350 = extfn448(l034C, c00DE)
        subFE0 (l0350)
DoEvents
subEC8 (.1#)
Else
    l0320% = extfn448(l031E%, c00DE)
End If
Call subFA8(l031E%, p030E)
Call subFA8(l0320%, p0312)
l0322% = extfn250(l031C%, "Send")
subFE0 (l0322%)
DoEvents
subEC8 (.1#)
Do
DoEvents
subEC8 (.1#)
l0324% = extfn288("#32770", 0&)
l0326% = extfn250(l0324%, "OK")
subFE0 (l0326%)
    DoEvents
Loop Until l0324% <> 0 Or extfn250(l0316%, "Available?") = 0
    sub1018 (l0324%)
        DoEvents
DoEvents
sub1018 (l031C%)
End Sub

Sub sub1168 (p0354 As String)
Dim l0356 As Integer
Dim l0358 As Variant
On Error Resume Next
For l0356 = Len(p0354) To 1 Step -1
    If  Mid$(p0354, l0356, 1) = Chr(13) Then
        If  l0356 <> Len(p0354) Then
            If  Mid$(p0354, l0356 + 1, 1) <> Chr(10) Then
                p0354 = Left$(p0354, l0356) & Chr(10) & Mid$(p0354, l0356 + 1)
            End If
        Else
            p0354 = p0354 & Chr(10)
        End If
    End If
Next l0356
l0358 = p0354
End Sub

Sub sub11A0 (p035E As Form)
Dim l0360 As String
Dim l0362 As Variant
Dim l0366 As Variant
Dim l036E As String
Dim l0370 As Variant
Dim l0378 As Integer
Dim l037A As Integer
Dim l037C As Variant
Dim l0380 As Variant
Dim l0386 As Integer
Dim l0388 As String
' Const c038A = 1040 ' &H410%
Call sub11D8(l0360$)
DoEvents
l0378% = extfn288("AOL Frame25", 0&)
l037A% = extfn250(l0378%, "TCP/IP")
l037C = 23
If  l037A% = 0 Then l037A% = extfn250(l0378%, "Rooms") And l037C = 48
l0380 = extfn218(l037A%, "_AOL_combobox")
For l0386% = 0 To 7
 l0388$ = String$(256, " ")
 l0370 = extfn100(l0386%, l0388$) & l036E$
 If  Len(Trim$(l0388$)) <= 1 Then GoTo L4296
 l0388$ = Left$(Trim$(l0388$), Len(Trim(l0388$)) - 1)
l0362 = extfn1800(p035E.List1.hWnd, c038A, 0, l0388$)
 l0366 = extfn1800(p035E.List1.hWnd, c038A, l0362, l0388$)
  If  l0366 >= 0 Then
GoTo L427E
   Else
   End If
      GoTo L4284
L427E:
L4284:
Next l0386%
L4296:
End Sub

Sub subFE0 (p038E As Variant)
Dim l0392 As Variant
' Const c0396 = 513 ' &H201%
' Const c0398 = 514 ' &H202%
    l0392 = extfn6E8(p038E, c0396, 0&, 0&)
    l0392 = extfn6E8(p038E, c0398, 0&, 0&)
End Sub

Sub sub1210 ()
Dim l039A As Variant
  l039A = extfn288("AOL Frame25", 0&)
  sub1018 (l039A)
  Call sub1248
End Sub

Sub sub1018 (p039E As Variant)
Dim l03A2 As Variant
' Const c03A6 = 16 ' &H10%
l03A2 = extfn6B0(p039E, c03A6, 0, 0)
End Sub

Sub sub1248 ()
Dim l03A8 As Integer
Dim l03AA As Variant
Dim l03AE As Variant
Do
l03A8% = extfn288("AOL Frame25", "America  Online")
l03AA = extfn288%("_AOL_Modal", "America Online")
l03AE = extfn6E8(l03AA, c03A6, 0, 0&)
    DoEvents
Loop Until l03AA = 0
End Sub

Sub subFA8 (p03B2 As Variant, p03B6 As Variant)
Dim l03BA As Variant
    l03BA = extfn720(p03B2, c00E0, 0, p03B6)
End Sub

Sub sub1280 ()
Dim l03BE As Integer
Dim l03C0 As Integer
Dim l03C2 As Integer
Dim l03C4 As Integer
Dim l03C6 As Variant
Dim l03CE As Variant
Dim l03D2 As Variant
Dim l03D6 As Integer
Dim l03D8 As Integer
l03C6 = extfn288(0&, "America  Online")
l03CE = fn9C0("Mem&bers", "Send an Instant Message" + Chr$(9) + "Ctrl+I")
Do
For l03D2 = 1 To 25
DoEvents
Next l03D2
l03C0 = extfn250(l03C6, "Send Instant Message")
Loop Until l03C0
DoEvents
l03BE = extfn218(l03C0, "_AOL_Edit")
Call subFA8(l03BE, "$IM_OFF")
l03C2 = extfn448(l03BE, c00DE)
Call subFA8(l03C2, "- -Instant Messages OFF- -")
l03C0 = extfn250(l03C6, "Send Instant Message")
l03C4 = extfn250(l03C0, "Send")
subFE0 (l03C4)
Do
    DoEvents
    l03D6% = extfn2C0()
    l03D8% = extfn250(l03D6%, "OK")
Loop Until l03D8%
Call subFE0(l03D8%)
sub1018 (l03C0)
End Sub

Sub sub12B8 ()
Dim l03DE As Integer
Dim l03E0 As Integer
Dim l03E2 As Integer
Dim l03E4 As Integer
Dim l03E6 As Variant
Dim l03EE As Variant
Dim l03F2 As Variant
Dim l03F6 As Integer
Dim l03F8 As Integer
l03E6 = extfn288(0&, "America  Online")
l03EE = fn9C0("Mem&bers", "Send an Instant Message" & Chr$(9) & "Ctrl+I")
Do
For l03F2 = 1 To 25
DoEvents
Next l03F2
l03E0 = extfn250(l03E6, "Send Instant Message")
Loop Until l03E0
DoEvents
l03DE = extfn218(l03E0, "_AOL_Edit")
Call subFA8(l03DE, "$IM_ON")
l03E2 = extfn448(l03DE, c00DE)
Call subFA8(l03E2, "- -Instant Messages ON- -")
l03E4 = extfn250(l03E0, "Send")
Call subFE0(l03E4)
Do
    DoEvents
    l03F6% = extfn2C0()
    l03F8% = extfn250(l03F6%, "OK")
Loop Until l03F8%
Call subFE0(l03F8%)
sub1018 (l03E0)
End Sub

Sub sub12F0 (p0400 As Label, p0402 As Integer, p0404 As Integer)
Dim l0406 As Variant
p0400.Visible = True
Do
For l0406 = 1 To p0404
DoEvents
Next l0406
p0400.FontSize = p0400.FontSize + 1
Loop Until p0400.FontSize >= p0402
End Sub

Sub sub1328 (p040C As Label, p040E As Integer, p0410 As Integer)
Dim l0412 As Variant
Do
For l0412 = 1 To p0410
DoEvents
Next l0412
p040C.FontSize = p040C.FontSize - 1
Loop Until p040C.FontSize <= p040E
p040C.Visible = False
End Sub

Sub sub1360 (p0418 As ListBox, p041A As String)
Dim l041C As Variant
Dim l0420 As Variant
Dim l0424 As Variant
Dim l0428 As Variant
Dim l042C As Variant
l041C = InStr(p041A, ",")
If  l041C Then
    l0420 = Mid$(p041A, 1, l041C - 1)
    p0418.AddItem l0420
Else
    p0418.AddItem p041A
    Exit Sub
End If
Do
l0424 = InStr(l041C + 1, p041A, ",")
If  l0424 Then
    l0428 = Len(p041A) - (l041C + 1) - (Len(p041A) - l0424)
    l0420 = Mid$(p041A, l041C + 1, l0428)
    p0418.AddItem l0420
ElseIf l0424 = False Then
    l042C = Mid$(p041A, l041C + 1)
    If  l042C = "" Then Exit Sub
    p0418.AddItem l042C
End If
l041C = l0424
Loop Until l041C = False
End Sub

Sub subE90 (p0430 As Integer, p0432 As Integer)
Dim l0434 As Integer
Dim l0436 As Integer
Dim l0438 As Integer
Dim l043A As Integer
Dim l043C As Long
l0436% = extfn368%(extfn288("AOL Frame25", 0&))
l0438% = extfn4B8%(l0436%, p0430)
l043A = extfn3D8%(l0438%, p0432)
l043C = CLng(0) * &H10000& Or gv0014
l0434 = extfn6E8(extfn288("AOL Frame25", 0&), c00C6, l043A, 0&)
End Sub

Sub subEC8 (p043E As Variant)
Dim l0442 As Variant
Dim l0446 As Variant
l0442 = Timer
Do While Timer - l0442 < p043E
l0446 = DoEvents()
Loop
End Sub

Sub sub1398 (p044C As Form)
' Const c044E =-1 ' &HFFFF%
' Const c0450 = 16 ' &H10%
' Const c0452 = 64 ' &H40%
 extsub8A8 p044C.hWnd, c044E, 0, 0, 0, 0, c0450 Or c0452
p044C.Left = (Screen.Width - p044C.Width) / 2
p044C.Top = (Screen.Height - p044C.Height) / 2
End Sub

Sub sub13D0 ()
Dim l0458 As Variant
Dim l045C As String
Dim l045E As Integer
Dim l0460 As Integer
Dim l0462 As Variant
Dim l046A As Variant
Dim l046E As Variant
Dim l0472 As Variant
Dim l0476 As Variant
Dim l047A As Variant
Dim l047E As Variant
Dim l0482 As Variant
Dim l0486 As Variant
Dim l048A As Variant
Dim l048E As Variant
Dim l0492 As Variant
Dim l0496 As Variant
Dim l049A As Variant
Dim l049E As Variant
Dim l04A2 As Variant
Dim l04A6 As Variant
Dim l04AA As Variant
Dim l04AE As Variant
Dim l04B2 As Variant
Dim l04B6 As Variant
Dim l04BA As Integer
Dim l04BC As Variant
' Const c04C0 = 513 ' &H201%
' Const c04C2 = 514 ' &H202%
' Const c04C4 = 1025 ' &H401%
l0458 = extfn288("AOL Frame25", "America  Online")
l045C$ = "Set Preferences" + Chr$(9) + "Ctrl+="
l045E% = fn9C0("Mem&bers", l045C$)
    DoEvents
l0460% = extfn288("AOL Frame25", "America  Online")
l0462 = extfn250(l0460%, "Preferences")
l046A = extfn218%(l0462, "_AOL_Icon")
l046E = extfn528(l046A, c00DE)
l0472 = extfn528(l046E, c00DE)
l0476 = extfn528(l0472, c00DE)
l047A = extfn528(l0476, c00DE)
l047E = extfn528(l047A, c00DE)
l0482 = extfn528(l047E, c00DE)
l0486 = extfn528(l0482, c00DE)
l048A = extfn528(l0486, c00DE)
l048E = extfn528(l048A, c00DE)
l0492 = extfn528(l048E, c00DE)
l0496 = extfn528(l0492, c00DE)
l049A = extfn6B0(l0496, c04C0, 0&, 0&)
l049E = extfn6B0(l0496, c04C2, 0&, 0&)
l04A2 = extfn330()
l04A6 = extfn330()
Do Until l04A6 <> l04A2
l04AA = DoEvents()
l04A6 = extfn330()
Loop
l04AE = extfn288("_AOL_Modal", "Mail Preferences")
l04B2 = extfn250(l04AE, "Confirm mail after it has been sent")
l04B6 = extfn250(l04AE, "Close mail after it has been sent")
l04BA% = extfn6B0(l04B2, c04C4, True, 0&)
l04BA% = extfn6B0(l04B6, c04C4, True, 0&)
l04BC = extfn250(l04AE, "OK")
l04BA% = extfn6E8&(l04BC, &H102, &HD, 1)
sub1018 (l0462)
End Sub
