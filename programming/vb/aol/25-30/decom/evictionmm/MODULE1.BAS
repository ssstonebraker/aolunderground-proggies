' Module1
Option Explicit
Declare Function extfn04A2 Lib "Apiguide.dll" Alias "agGetStringFromLPSTR" (ByVal p1&) As String
Declare Function extfn0466 Lib "311.dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn0405 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn0327 Lib "GDI" Alias "BitBlt" () As Integer
Declare Sub extsub0291 Lib "User" Alias "bringwindowtotop" ()
Declare Function extfn034F Lib "User" Alias "CopyRect" () As Integer
Declare Function extfn064F Lib "GDI" Alias "CreateDC" () As Integer
Declare Function extfn0426 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn0413 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn03E6 Lib "User" Alias "DeleteMenu" () As Integer
Declare Sub extsub03B7 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn05A6 Lib "User" Alias "EnableHardwareInput" () As Integer
Declare Function extfn05BD Lib "User" Alias "EnableWindow" () As Integer
Declare Function extfn065B Lib "User" Alias "SetWindowPos" () As Integer
Declare Function extfn0524 Lib "vbwfind.dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn0510 Lib "vbwfind.dll" Alias "FindChildByTitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn04BA Lib "User" Alias "findwindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn02A5 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn0481 Lib "User" Alias "GetClassName" () As Integer
Declare Sub extsub03A7 Lib "User" Alias "GetCursorPos" ()
Declare Function extfn031E Lib "User" Alias "GetDC" () As Integer
Declare Function extfn030A Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn04F4 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn0434 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn02CB Lib "User" Alias "GetMenuItemCount" (ByVal p1%) As Integer
Declare Function extfn03F4 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0500 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn05CD Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn055D Lib "User" Alias "GetNextWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0474 Lib "User" Alias "getparent" (ByVal p1%) As Integer
Declare Function extfn056E Lib "Kernel" Alias "GetPrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn05DE Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn037C Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn063B Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn04C8 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Sub extsub033E Lib "User" Alias "GetWindowRect" ()
Declare Function extfn0491 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn0546 Lib "User" Alias "GetWindowTextLength" (ByVal p1%) As Integer
Declare Function extfn03D8 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn044A Lib "User" Alias "isiconic" () As Integer
Declare Function extfn05F2 Lib "User" Alias "IsWindowVisible" () As Integer
Declare Function extfn0366 Lib "Kernel" Alias "lstrcpy" () As Long
Declare Function extfn035B Lib "Kernel" Alias "lstrlen" () As Integer
Declare Sub extsub061B Lib "User" Alias "MoveWindow" ()
Declare Sub extsub0629 Lib "User" Alias "ReleaseCapture" ()
Declare Function extfn0331 Lib "User" Alias "ReleaseDC" () As Integer
Declare Function extfn038A Lib "User" Alias "RemoveMenu" () As Integer
Declare Function extfn0398 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn02DF Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn02F3 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn02B8 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn04D5 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn0371 Lib "User" Alias "SetMenu" () As Integer
Declare Sub extsub04E4 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn0538 Lib "User" Alias "ShowWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0456 Lib "MMSYSTEM" Alias "sndPlaySound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn03C6 Lib "User" Alias "TrackPopupMenu" () As Integer
Declare Function extfn043F Lib "Kernel" Alias "WinExec" () As Integer
Declare Function extfn0589 Lib "Kernel" Alias "WritePrivateProfileString" () As Integer
Declare Function extfn0605 Lib "Kernel" Alias "WriteProfileString" () As Integer

Sub sub007C (p011C As String, p0120 As ListBox)
Dim l0122 As Variant
Dim l0126 As String
Dim l0128 As Variant
If  p0120.ListCount = 0 Then
p0120.AddItem p011C
Exit Sub
End If
Do Until l0122 = (p0120.ListCount)
Let l0126$ = p0120.List(l0122)
If  Trim(LCase(l0126$)) = Trim(LCase(p011C)) Then Let l0128 = "NO"
Let l0122 = l0122 + 1
Loop
If  l0128 = "NO" Then Exit Sub
p0120.AddItem p011C
End Sub

Sub sub008A (p012E As ListBox)
Dim l0130 As Integer
Dim l0132 As String
Dim l0134 As String
Dim l0136 As Variant
For l0130% = 0 To 25
l0132$ = String$(256, " ")
l0136 = extfn0466(l0130%, l0132$) & l0134$
If  Len(Trim$(l0132$)) <= 1 Then GoTo LAC4
l0132$ = Left$(Trim$(l0132$), Len(Trim(l0132$)) - 1)
sub007C l0132$, p012E
Next l0130%
LAC4:
End Sub

Sub sub0095 (p013C As ListBox)
Dim l013E As Integer
Dim l0140 As Integer
Dim l0142 As Integer
Dim l0144 As Integer
Dim l0146 As Variant
Dim l014A As Variant
Dim l014E As String
Dim l0150 As Integer
Dim l0152 As String
Dim l0154 As Integer
Dim l0156 As Integer
Dim l0158 As String
Dim l015A As Integer
Dim l015C As Variant
l013E% = extfn04BA("AOL Frame25", 0&)
Debug.Print l013E%
l0140% = extfn0510(l013E%, "List Rooms")
Debug.Print l0140%
l0142% = extfn0474(l0140%)
Debug.Print l0142%
l0144% = extfn0524(l0142%, "_AOL_Listbox")
If  l0144% = 0 Then Exit Sub
If  l0144% = 0 Then GoTo 19
l0146 = extfn0398(l0144%, gc0076, 0&, 0&)
For l014A = 0 To l0146 - 1
l014E$ = String$(64, 0)
l0150% = extfn0466%(l014A, l014E$)
l0152$ = Left$(l014E$, l0150%)
l013E% = extfn04BA("AOL Frame25", 0&)
l0154% = extfn0524(l013E%, "MDIClient")
l0156 = extfn0510(l0154%, "Welcome, ")
l0158 = String$(22, 0)
l015A% = extfn0491(l0156%, l0158$, 22)
l0158$ = Left$(l0158$, l015A%)
l0158$ = Mid$(l0158$, InStr(l0158$, ", "), InStr(l0158$, "!") - 1)
l0158$ = Left$(l0158$, Len(l0158$) - 1)
l0158$ = Right$(l0158$, Len(l0158$) - 2)
If  l0158$ = l0152$ Then GoTo 18
For l015C = 0 To p013C.ListCount - 1
 If  l0152$ = p013C.List(l015C) Then GoTo 18
Next l015C
p013C.AddItem l0152$
18 :
Next l014A
19 :
End Sub

Sub sub00A9 (p0160 As Integer)
Dim l0162 As String
Dim l0164 As Variant
Dim l0168 As Variant
If  l0162$ = "OFF" Then
Exit Sub
End If
sub00ED (.1#)
l0164 = extfn02DF(p0160, gc0064, 0, 0&)
l0168 = extfn02DF(p0160, gc0066, 0, 0&)
End Sub

Function fn00B5 () As Variant
Dim l0170 As Variant
Dim l0174 As Integer
Dim l0176 As Integer
Dim l0178 As Integer
l0170 = 1
l0174% = extfn04BA("AOL Frame25", 0&)
l0176% = extfn0524(l0174%, "AOL Toolbar")
l0178% = extfn0524(l0176%, "_AOL_Icon")
Do
    gv00CC% = DoEvents()
    l0178% = extfn04C8(l0178%, gc0084)
    l0170 = l0170 + 1
Loop Until l0178% = 0
If  l0170 = 19 Or fn0111("Exit Unlimited Use area") <> 0 Then
    fn00B5 = 30
ElseIf l0170 = 21 Or fn0111("Exit Free Area") <> 0 Then
    fn00B5 = 25
End If
End Function

Sub sub00C3 (p017A As String)
Dim l017C As Variant
Dim l0180 As Variant
Dim l0184 As Integer
Dim l0186 As Integer
Dim l0188 As Integer
Dim l018A As Integer
Dim l018C As Variant
Dim l0190 As Integer
Dim l0192 As Integer
Dim l0194 As Integer
Dim l0196 As Integer
Dim l0198 As Integer
Dim l019A As Integer
Dim l019C As Integer
Dim l019E As Integer
Dim l01A0 As Variant
Dim l01A4 As Integer
Dim l01A6 As Integer
Dim l01A8 As Integer
l017C = "NOT DONE"
l0180 = "YET"
l0184% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l0184%, "Compose Mail")
Do
DoEvents
l0186% = extfn0524(l0184%, "MDIClient")
l0188% = extfn0510(l0186%, "Compose Mail")
l018A% = extfn0524(l0188%, "_AOL_Edit")
Loop Until l018A% <> 0
sub0269 (.6#)
l018C = extfn02F3(l018A%, gc0068, 0, p017A)
Do
DoEvents
l0190% = extfn0524(l0184%, "MDIClient")
l0192% = extfn0510(l0190%, "Compose Mail")
l0194% = extfn0524(l0192%, "_AOL_Edit")
Loop Until l0194% <> 0
Do
DoEvents
l0196% = extfn04C8(l0194%, gc0084)
l0198% = extfn04C8(l0196%, gc0084)
l019A% = extfn04C8(l0198%, gc0084)
l019C% = extfn04C8(l019A%, gc0084)
Loop Until l019C% <> 0
l018C = extfn02F3(l019C%, gc0068, 0, "²º·‚¸‚•²º·‚¸‚• O.c.W Application")
Do
DoEvents
l019E% = extfn0524(l0188%, "RICHCNTL")
Loop Until l019E% <> 0
l01A0 = extfn02F3(l019E%, gc0068, 0, l017C + l0180)
Do
DoEvents
l01A4% = extfn0524(l0188%, "_AOL_ICON")
Loop Until l01A4% <> 0
l01A6% = extfn02DF(l01A4%, gc0064, 0&, 0&)
l01A6% = extfn02DF(l01A4%, gc0066, 0&, 0&)
sub0269 (.6#)
l01A8% = extfn0398(l0188%, gc006E, 0, 0)
End Sub


Function fn00CD () As String
Dim l01AC As Integer
Dim l01AE As String
Dim l01B0 As Variant
l01AC% = fn0101()
If  l01AC% = 0 Then
    fn00CD = ""
End If
l01AE$ = Space$(extfn0546(l01AC%))
l01B0 = extfn0491(l01AC%, l01AE$, 255)
If  InStr(l01AE$, Chr$(0)) <> 0 Then
    l01AE$ = Left$(l01AE$, InStr(l01AE$, Chr$(0)) - 1)
End If
fn00CD = l01AE$
End Function

Function fn00D8 () As Integer
Dim l01B6 As Integer
Dim l01B8 As Integer
Dim l01BA As Integer
Dim l01BC As Integer
Dim l01BE As String
 l01B6% = extfn04BA("AOL Frame25", 0&)
 If  l01B6% = 0 Then
  fn00D8 = 0
  Exit Function
 End If
 l01B8% = extfn0524(l01B6%, "MDIClient")
 l01BA% = extfn0510(l01B8%, "Welcome")
 : l01BE$ = String$(200, 0)
 l01BC% = extfn0491(l01BA%, l01BE$, 199)
 l01BE$ = Left$(l01BE$, l01BC%)
 l01BC% = InStr(1, l01BE$, "!", 1)
 If  l01BC% = 0 Then
  fn00D8 = 0
 Else
  fn00D8 = 1
 End If
End Function

Function fn00E1 () As Variant
Dim l01C4 As Integer
Dim l01C6 As Integer
Dim l01C8 As Variant
l01C4% = extfn04BA(0&, "America  Online")
l01C6% = extfn0524(l01C4%, "_AOL_TREE")
If  l01C6% = 0 Then MsgBox "Open your Mailbox and try again"
DoEvents
l01C8 = extfn02DF(l01C6%, gc0076, 0, 0)
If  l01C8 > 1 And l01C8 <> 550 Then
MsgBox "You have " & l01C8 & " messages in your Mailbox.", 0
End If
If  l01C8 = 1 Then
MsgBox "You have " & l01C8 & " message in your Mailbox.", 0
End If
If  l01C8 < 1 Then
MsgBox "You have no messages in your Mailbox.", 0
End If
If  l01C8 = 550 Then
MsgBox "Your Mailbox iz full."
End If
End Function

Sub sub00ED (ByVal pv01CC As Variant)
Dim l01D0 As Variant
l01D0 = Timer
Do
 DoEvents
Loop Until ((Timer) >= (l01D0 + pv01CC))
End Sub

Sub sub00F6 (p01D6 As ListBox, ByVal p01D8 As String)
Dim l01DA As Variant
Do
If  LCase$(p01D6.List(l01DA)) = LCase$(p01D8$) Then p01D6.RemoveItem (l01DA)
l01DA = 1 + l01DA
Loop Until l01DA >= p01D6.ListCount
End Sub

Function fn0101 () As Variant
Dim l01E2 As Integer
Dim l01E4 As Integer
Dim l01E6 As Integer
Dim l01E8 As Integer
Dim l01EA As Integer
Dim l01EC As Integer
Dim l01EE As Integer
Dim l01F0 As Integer
Dim l01F2 As Integer
l01E2% = extfn04BA("AOL Frame25", 0&)
l01E4% = extfn0524(l01E2%, "MDIClient")
l01E6% = extfn0524(l01E4%, "AOL Child")
Do
    l01E8% = extfn0524(l01E6%, "_AOL_Listbox")
    l01EA% = extfn0524(l01E6%, "_AOL_View")
    l01EC% = extfn0524(l01E6%, "_AOL_Edit")
    l01EE% = extfn0524(l01E6%, "_AOL_Icon")
    l01F0% = extfn0524(l01E6%, "_AOL_Glyph")
    l01F2% = extfn0524(l01E6%, "_AOL_Image")
    If  l01E8% = 0 Or l01EA% = 0 Or l01EC% = 0 Or l01EE% = 0 Or l01F2% = 0 Then
        l01E6% = extfn04C8(l01E6%, gc0084)
    End If
Loop Until l01E6% = 0 Or (l01E8% <> 0 And l01EA% <> 0 And l01EC% <> 0 And l01EE% <> 0 And l01F2% <> 0)
If  l01E6% = 0 Then
    fn0101 = 0
Else
    fn0101 = l01E6%
End If
End Function

Function fn0111 (ByVal p01F8 As String) As Variant
Dim l01FA As Integer
Dim l01FC As Integer
Dim l01FE As Integer
Dim l0200 As Integer
Dim l0202 As Integer
Dim l0204 As Integer
Dim l0206 As Integer
Dim l0208 As String
Dim l020A As Integer
l01FA% = extfn04BA("AOL Frame25", 0&)
l01FC% = extfn0434(l01FA%)
l01FE% = extfn02CB(l01FC%)
For l0200% = 0 To l01FE% - 1
    l0202% = extfn037C(l01FC%, l0200%)
    l0204% = extfn02CB(l0202%)
    For l0206% = 0 To l0204% - 1
        l0208$ = Space$(50)
        l020A% = extfn05CD(l0202%, l0206%, l0208$, 50, gc00B8)
        l0208$ = Trim$(l0208$): l0208$ = Left(l0208$, Len(l0208$) - 1)
        If  l0208$ = "" Then l0208$ = " -"
        If  InStr(l0208$, p01F8$) Then
            fn0111 = 1
            Exit Function
        End If
    Next l0206%
Next l0200%
fn0111 = 0
End Function

Function fn0120 (ByVal pv020E As Variant) As String
fn0120 = Trim(Left$(pv020E, InStr(pv020E, Chr$(0)) - 1))
End Function

Function fn0129 (p0214 As String, p0216 As String, p0218 As String) As String
Dim l021A As String
   l021A = String(255, Chr(0))
   fn0129 = Left(l021A, extfn056E(p0214$, ByVal p0216$, "", l021A, Len(l021A), p0218$))
End Function

Sub sub0137 (p021C As Integer, p021E As Variant)
Dim l0222 As Variant
Do
If  p021E = 0 Then Exit Sub
l0222 = 1 + l0222
p021C% = extfn04C8(p021C%, gc0084)
Loop Until l0222 >= p021E - 1
End Sub

Function fn0141 () As String
Dim l0228 As Integer
Dim l022A As Integer
Dim l022C As Variant
Dim l0230 As String
Dim l0232 As Variant
Dim l0236 As Variant
Dim l023A As String
Dim l023C As String
    On Error GoTo L2170
    l0228% = extfn04BA("AOL Frame25", 0&)
    l022A% = extfn0510(l0228%, "Welcome, ")
    l022C = extfn0546(l022A%)
    l0230$ = Space(l022C)
    l0232 = extfn0491(l022A%, l0230$, l022C)
    l0236 = InStr(l0230$, ", ")
    If  l0236 = False Then Exit Function
    l023A$ = Mid$(l0230$, l0236 + 2)
    l023C$ = Mid$(l023A$, 1, Len(l023A$) - 1)
    l023C$ = Trim(l023C$)
    fn0141 = l023C$
L2170:
Resume Next
End Function

Function fn014A (p0240 As Integer) As String
Dim l0242 As Variant
Dim l0246 As String
Dim l0248 As Variant
l0242 = extfn0398(p0240, gc0070, 0&, 0&)
l0246$ = Space$(l0242 + 1)
l0248 = extfn02F3(p0240, gc0062, l0242 + 1, l0246$)
fn014A = l0246$
End Function

Function fn0158 (p024E As Integer) As String
Dim l0250 As Variant
Dim l0254 As String
Dim l0256 As Variant
Dim l025A As Integer
Dim l025C As String
l0250 = extfn0398(p024E, gc0070, 0&, 0&)
l0254$ = Space$(l0250 + 1)
l0256 = extfn02F3(p024E, gc0062, l0250 + 1, l0254$)
l025A% = InStr(l0254$, "Fwd: ")
If  l025A% <> 0 Then
l025C$ = Mid$(l0254$, 6, (Len(l0254$) - l025A%))
fn0158 = l025C$
End If
If  l025A% = 0 Then
fn0158 = l0254$
End If
End Function


Sub sub016B (p025E As Variant)
Dim l0262 As Variant
l0262 = extfn0538(p025E, gv00B0)
End Sub

Sub sub0176 (ByVal p0266 As String)
Dim l0268 As Integer
Dim l026A As Integer
Dim l026C As Integer
Dim l026E As Integer
Dim l0270 As Integer
Dim l0272 As Integer
Dim l0274 As Integer
Dim l0276 As String
Dim l0278 As Integer
Dim l027A As Integer
l0268% = extfn04BA("AOL Frame25", 0&)
l026A% = extfn0434(l0268%)
l026C% = extfn02CB(l026A%)
For l026E% = 0 To l026C% - 1
l0270% = extfn037C(l026A%, l026E%)
l0272% = extfn02CB(l0270%)
For l0274% = 0 To l0272% - 1
l0276$ = Space$(50)
l0278% = extfn05CD(l0270%, l0274%, l0276$, 50, gc00B8)
l0276$ = Trim$(l0276$): l0276$ = Left(l0276$, Len(l0276$) - 1)
If  l0276$ = "" Then l0276$ = " -"
If  InStr(l0276$, p0266$) Then
l027A% = extfn03F4(l0270%, l0274%)
l0278% = extfn02DF(l0268%, gc0060, l027A%, 0)
Exit Sub
End If
Next l0274%
Next l026E%
End Sub

Sub sub0182 ()
Dim l027C As Integer
Dim l027E As Integer
Dim l0280 As Integer
Dim l0282 As Integer
Dim l0284 As Integer
Dim l0286 As Integer
Dim l0288 As Integer
Dim l028A As Variant
Dim l028E As Integer
Dim l0290 As Integer
Dim l0292 As Integer
l027C% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l027C%, "Send an instant Message")
Do
DoEvents
l027E% = extfn0510(l027C%, "Send Instant Message")
Loop Until l027E% <> 0
Do
DoEvents
l0280% = extfn0524(l027E%, "_AOL_Edit")
Loop Until l0280% <> 0
l0282% = extfn02F3(l0280%, gc0068, 0, "$Im_Off")
l0284% = extfn0524(l027E%, "RICHCNTL")
If  l0284% <> 0 Then
l0286% = extfn02F3(l0284%, gc0068, 0, "O.c.W OwNz yer Ass")
l0288% = extfn055D(l0284%, gc0084)
l028A = extfn02DF(l0288%, gc0064, 0, 0&)
l028A = extfn02DF(l0288%, gc0066, 0, 0&)
sub00ED (.2#)
l028E% = extfn0398(l027E%, gc006E, 0, 0)
sub00ED (.1#)
End If
If  l0284% = 0 Then
l0290% = extfn055D(l0280%, gc0084)
l0286% = extfn02F3(l0290%, gc0068, 0, "O.c.W OwNz yer Ass")
l0288% = extfn055D(l0290%, gc0084)
l028A = extfn02DF(l0288%, gc0064, 0, 0&)
l028A = extfn02DF(l0288%, gc0066, 0, 0&)
sub0269 (.4#)
l028E% = extfn0398(l027E%, gc006E, 0, 0)
End If
Do
DoEvents
l0292% = extfn04BA("#32770", 0&)
Loop Until l0292% <> 0
sub01AD
sub0269 (.2#)
sub01AD
sub0269 (.2#)
End Sub


Sub sub018B ()
Dim l0294 As Integer
Dim l0296 As Integer
Dim l0298 As Integer
Dim l029A As Integer
Dim l029C As Integer
Dim l029E As Integer
Dim l02A0 As Integer
Dim l02A2 As Variant
Dim l02A6 As Integer
Dim l02A8 As Integer
Dim l02AA As Integer
l0294% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l0294%, "Send an instant Message")
Do
DoEvents
l0296% = extfn0510(l0294%, "Send Instant Message")
Loop Until l0296% <> 0
Do
DoEvents
l0298% = extfn0524(l0296%, "_AOL_Edit")
Loop Until l0298% <> 0
l029A% = extfn02F3(l0298%, gc0068, 0, "$Im_On")
l029C% = extfn0524(l0296%, "RICHCNTL")
If  l029C% <> 0 Then
l029E% = extfn02F3(l029C%, gc0068, 0, "O.c.W OwNz yer Ass")
l02A0% = extfn055D(l029C%, gc0084)
l02A2 = extfn02DF(l02A0%, gc0064, 0, 0&)
l02A2 = extfn02DF(l02A0%, gc0066, 0, 0&)
sub00ED (.2#)
l02A6% = extfn0398(l0296%, gc006E, 0, 0)
sub00ED (.1#)
End If
If  l029C% = 0 Then
l02A8% = extfn055D(l0298%, gc0084)
l029E% = extfn02F3(l02A8%, gc0068, 0, "O.c.W OwNz yer Ass")
l02A0% = extfn055D(l02A8%, gc0084)
l02A2 = extfn02DF(l02A0%, gc0064, 0, 0&)
l02A2 = extfn02DF(l02A0%, gc0066, 0, 0&)
sub0269 (.4#)
l02A6% = extfn0398(l0296%, gc006E, 0, 0)
End If
Do
DoEvents
l02AA% = extfn04BA("#32770", 0&)
Loop Until l02AA% <> 0
sub0269 (.1#)
l02A6% = extfn0398(l02AA%, gc006E, 0, 0)
sub0269 (.3#)
If  l02AA% <> 0 Then
Do
DoEvents
l02AA% = extfn04BA("#32770", 0&)
Loop Until l02AA% <> 0
sub0269 (.1#)
l02A6% = extfn0398(l02AA%, gc006E, 0, 0)
End If
End Sub


Sub sub0193 (p02AE As Form)
extsub04E4 p02AE.hWnd, -1, 0, 75, 560, 35, &H50
End Sub

Sub sub01A0 ()
Dim l02B0 As Integer
Dim l02B2 As Variant
Dim l02B6 As Integer
Do
DoEvents
Do
DoEvents
l02B0% = extfn04BA("_AOL_MODAL", "America Online")
Loop Until l02B0% <> 0
l02B6% = extfn0398(l02B0%, gc006E, l02B2, 0)
sub0269 (.1#)
Loop Until l02B0% = 0
End Sub

Sub sub01AD ()
Dim l02B8 As Integer
Dim l02BA As Integer
l02B8% = extfn04BA("#32770", 0&)
If  l02B8% <> 0 Then
l02BA% = extfn0398(l02B8%, gc006E, 0, 0)
End If
End Sub

Sub sub01B8 ()
If  fn00B5() = 25 Then
    sub0176 "Exit Free Area"
End If
If  fn00B5() = 30 Then
    sub0176 "Exit Unlimited Use area"
End If
gv00CC% = DoEvents()
End Sub

Sub sub01C4 ()
Dim l02BC As Variant
Dim l02C0 As String
Dim l02C2 As Variant
Dim l02C6 As Variant
Dim l02CA As Variant
Dim l02CE As Variant
Dim l02D2 As Variant
    Select Case Err
        Case l02BC
            l02C0$ = "Error load media device driver."
        Case l02C2
            l02C0$ = "The device is not open or is not known."
        Case l02C6
            l02C0$ = "Invalid device ID."
        Case l02CA
            l02C0$ = "Invalid filename."
        Case l02CE
            l02C0$ = "Action not available for this device."
        Case Else
            l02C0$ = "Unknown Error (" + Str$(Err) + ")."
    End Select
    MsgBox l02C0$, 48, l02D2
End Sub

Sub sub01CF (p02D8 As ListBox, p02DA As String)
Dim l02DC As Variant
Dim l02E0 As Integer
Dim l02E2 As String
Dim l02E4 As Variant
Dim l02E8 As Variant
Dim l02EC As Variant
Dim l02F0 As Variant
Dim l02F4 As Integer
Dim l02F6 As Variant
Dim l02FA As Integer
Dim l02FC As Integer
Dim l02FE As Integer
Dim l0300 As String
Dim l0302 As String
Dim l0304 As Variant
Dim l0308 As Variant
Do
For l02DC = 0 To p02D8.ListCount - 1
If  l02E0% = 1 Then Exit Sub
l02E2$ = p02D8.List(0)
p02D8.ListIndex = 0
Next l02DC
l02E4 = extfn04BA("#32770", "America Online")
l02E8 = extfn0510(l02E4, "OK")
l02EC = extfn02DF(l02E8, gc0064, 0, 0&)
l02F0 = extfn02DF(l02E8, gc0066, 0, 0&)
l02F4% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l02F4%, "Send an instant Message")
Do
l02F6 = extfn0510(l02F4%, "Send Instant Message")
l02FA% = extfn0524(l02F6, "_AOL_Edit")
DoEvents
Loop Until l02FA% <> 0
l02FA% = extfn0524(l02F6, "_AOL_Edit")
Do
l02FC% = extfn0524(l02F6, "RICHCNTL")
l02FE% = extfn0510(l02F6, "Send")
DoEvents
sub00ED (.001#)
Loop Until l02FC% <> 0 Or l02FE% <> 0
If  l02FC% <> 0 Then
sub0222 l02FA%, l02E2$
sub0222 l02FC%, ((p02DA$) & Chr(13) & "           ·•¤^º°''˜¨¨GrOoVe Mass Im²º·‚¸‚•²º·‚¸‚•·" & Chr(13) & "        ·•¤^º°''˜¨¨by Flowie and Funkyeti²º·‚¸‚•²º·‚¸‚•·")
sub00ED (.001#)
sub0137 l02FC%, 1
sub00A9 l02FC%
Else
sub0222 l02FA%, l0300$
sub0137 l02FA%, 1
sub0222 l02FA%, l0302$
sub0137 l02FA%, 1
sub00A9 l02FA%
End If
sub00ED (.3#)
l0304 = extfn02DF(l02F6, gc006E, 0, 0)
l0308 = p02D8.List(0)
Call sub00F6(p02D8, (l0308))
Loop Until p02D8.ListCount = 0
If  p02D8.ListCount = 0 Then
Exit Sub
End If
End Sub

Sub sub01E5 ()
Dim l030C As Integer
Dim l030E As Integer
Dim l0310 As Integer
Dim l0312 As Integer
Dim l0314 As Integer
Dim l0316 As Integer
Dim l0318 As Integer
Dim l031A As Integer
Dim l031C As Integer
Dim l031E As Integer
Dim l0320 As Integer
l030C% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l030C%, "Preferences")
Do: DoEvents
l030E% = extfn0524(l030C%, "MDIClient")
l0310% = extfn0510(l030E%, "Preferences")
l0312% = extfn0510(l0310%, "Mail")
l0314% = extfn04C8(l0312%, gc0084)
If  l0312% <> 0 And l0314% <> 0 Then Exit Do
Loop
sub00A9 (l0314%)
Do: DoEvents
l0316% = extfn04BA("_AOL_Modal", "Mail Preferences")
l0318% = extfn0510(l0316%, "Close mail after it has been sent")
l031A% = extfn0510(l0316%, "Confirm mail after it has been sent")
l031C% = extfn0510(l0316%, "OK")
If  l031C% <> 0 And l0318% <> 0 And l031A% <> 0 Then Exit Do
Loop
l031E% = extfn0398(l0318%, gc0074, 1, 0)
l031E% = extfn0398(l031A%, gc0074, 0, 0)
sub00A9 (l031C%)
Do: DoEvents
l0316% = extfn04BA("_AOL_Modal", "Mail Preferences")
Loop Until l0316% = 0
l0320% = extfn0398(l0310%, gc006E, 0, 0)
End Sub

Sub sub01F5 (p0322 As Variant)
Dim l0326 As String
Dim l0328 As Variant
Dim l032C As Integer
Dim l032E As Integer
l0326$ = p0322
   l032C% = gc009A Or l0328
   l032E% = extfn0456(l0326$, l032C%)
End Sub

Sub sub0200 ()
Dim l0330 As Integer
Dim l0332 As Integer
Dim l0334 As Integer
Dim l0336 As Integer
Dim l0338 As Integer
Dim l033A As Integer
Dim l033C As Integer
Dim l033E As Integer
Dim l0340 As Integer
Dim l0342 As Integer
Dim l0344 As Integer
l0330% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l0330%, "Preferences")
Do: DoEvents
l0332% = extfn0524(l0330%, "MDIClient")
l0334% = extfn0510(l0332%, "Preferences")
l0336% = extfn0510(l0334%, "Mail")
l0338% = extfn04C8(l0336%, gc0084)
If  l0336% <> 0 And l0338% <> 0 Then Exit Do
Loop
sub00ED (.1#)
sub00A9 (l0338%)
Do: DoEvents
l033A% = extfn04BA("_AOL_Modal", "Mail Preferences")
l033C% = extfn0510(l033A%, "Close mail after it has been sent")
l033E% = extfn0510(l033A%, "Confirm mail after it has been sent")
l0340% = extfn0510(l033A%, "OK")
If  l0340% <> 0 And l033C% <> 0 And l033E% <> 0 Then Exit Do
Loop
l0342% = extfn0398(l033C%, gc0074, 0, 0)
l0342% = extfn0398(l033E%, gc0074, 0, 0)
sub00A9 (l0340%)
Do: DoEvents
l033A% = extfn04BA("_AOL_Modal", "Mail Preferences")
Loop Until l033A% = 0
l0344% = extfn0398(l0334%, gc006E, 0, 0)
End Sub

Sub sub020F (p0346 As Variant, p034A As Variant)
Dim l034E As String
Dim l0350 As Variant
Dim l0354 As Variant
Dim l0358 As Variant
Dim l035C As Variant
Dim l0360 As Variant
Dim l0364 As Variant
Dim l0368 As Variant
Dim l036C As String
Dim l036E As Variant
Dim l0372 As Variant
l034E$ = p034A
l0350 = extfn0434(p0346)
l0354 = extfn02CB(l0350)
For l0358 = 0 To l0354 - 1
l035C = extfn037C(l0350, l0358)
l0360 = extfn02CB(l035C)
For l0364 = 0 To l0360 - 1
    l0368 = extfn03F4(l035C, l0364)
    l036C$ = String$(100, " ")
    l036E = extfn05CD(l035C, l0368, l036C$, 100, 1)
    If  InStr(UCase(l036C$), UCase(l034E$)) Then
        l0372 = l0368
        GoTo L4182
    End If
Next l0364
Next l0358
L4182:
l036E = extfn02DF(p0346, &H111, l0372, 0)
End Sub

Sub sub0222 (p0376 As Variant, p037A As String)
Dim l037C As Variant
l037C = extfn02F3(p0376, gc0068, 0, p037A$)
End Sub

Sub sub022A (p0380 As String, p0382 As String)
Dim l0384 As Integer
Dim l0386 As Integer
Dim l0388 As Integer
Dim l038A As Integer
Dim l038C As Integer
Dim l038E As Integer
Dim l0390 As Integer
Dim l0392 As Variant
Dim l0396 As Integer
Dim l0398 As Integer
l0384% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l0384%, "Send an instant Message")
Do
DoEvents
l0386% = extfn0510(l0384%, "Send Instant Message")
Loop Until l0386% <> 0
Do
DoEvents
l0388% = extfn0524(l0386%, "_AOL_Edit")
Loop Until l0388% <> 0
l038A% = extfn02F3(l0388%, gc0068, 0, p0380)
l038C% = extfn0524(l0386%, "RICHCNTL")
If  l038C% <> 0 Then
l038E% = extfn02F3(l038C%, gc0068, 0, p0382)
l0390% = extfn055D(l038C%, gc0084)
l0392 = extfn02DF(l0390%, gc0064, 0, 0&)
l0392 = extfn02DF(l0390%, gc0066, 0, 0&)
sub00ED (.15#)
l0396% = extfn0398(l0386%, gc006E, 0, 0)
sub00ED (.01#)
End If
If  l038C% = 0 Then
l0398% = extfn055D(l0388%, gc0084)
l038E% = extfn02F3(l0398%, gc0068, 0, p0382)
l0390% = extfn055D(l0398%, gc0084)
sub00ED (.01#)
l0392 = extfn02DF(l0390%, gc0064, 0, 0&)
l0392 = extfn02DF(l0390%, gc0066, 0, 0&)
sub00ED (.04#)
l0396% = extfn0398(l0386%, gc006E, 0, 0)
sub00ED (.01#)
End If
End Sub

Sub sub0234 (p039A As String, p039C As String, p039E As String)
Dim l03A0 As Integer
Dim l03A2 As Integer
Dim l03A4 As Integer
Dim l03A6 As Integer
Dim l03A8 As Variant
Dim l03AC As Integer
Dim l03AE As Integer
Dim l03B0 As Integer
Dim l03B2 As Integer
Dim l03B4 As Integer
Dim l03B6 As Integer
Dim l03B8 As Integer
Dim l03BA As Integer
Dim l03BC As Variant
Dim l03C0 As Integer
Dim l03C2 As Integer
Dim l03C4 As Integer
Dim l03C6 As Integer
Dim l03C8 As Integer
l03A0% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l03A0%, "Compose Mail")
Do
DoEvents
l03A2% = extfn0524(l03A0%, "MDIClient")
l03A4% = extfn0510(l03A2%, "Compose Mail")
l03A6% = extfn0524(l03A4%, "_AOL_Edit")
Loop Until l03A6% <> 0
sub0269 (.6#)
l03A8 = extfn02F3(l03A6%, gc0068, 0&, p039A)
Do
DoEvents
l03AC% = extfn0524(l03A0%, "MDIClient")
l03AE% = extfn0510(l03AC%, "Compose Mail")
l03B0% = extfn0524(l03AE%, "_AOL_Edit")
Loop Until l03B0% <> 0
Do
DoEvents
l03B2% = extfn04C8(l03B0%, gc0084)
l03B4% = extfn04C8(l03B2%, gc0084)
l03B6% = extfn04C8(l03B4%, gc0084)
l03B8% = extfn04C8(l03B6%, gc0084)
Loop Until l03B8% <> 0
l03A8 = extfn02F3(l03B8%, gc0068, 0&, p039C)
l03BA% = extfn0524(l03A4%, "RICHCNTL")
If  l03BA% <> 0 Then
l03BC = extfn02F3(l03BA%, gc0068, 0&, p039E)
End If
If  l03BA% = 0 Then
l03C0% = extfn04C8(l03B8%, gc0084)
l03C2% = extfn04C8(l03C0%, gc0084)
l03C4% = extfn04C8(l03C2%, gc0084)
l03A8 = extfn02F3(l03C4%, gc0068, 0&, p039E)
End If
Do
DoEvents
l03C6% = extfn0524(l03A4%, "_AOL_ICON")
Loop Until l03C6% <> 0
l03C8% = extfn02DF(l03C6%, gc0064, 0&, 0&)
l03C8% = extfn02DF(l03C6%, gc0066, 0&, 0&)
l03C8% = extfn02DF(l03C6%, gc0064, 0&, 0&)
l03C8% = extfn02DF(l03C6%, gc0066, 0&, 0&)
End Sub

Sub sub0240 ()
Dim l03CA As Integer
Dim l03CC As Integer
Dim l03CE As Integer
Dim l03D0 As Integer
Dim l03D2 As Integer
Dim l03D4 As Integer
Dim l03D6 As Integer
Dim l03D8 As Integer
Dim l03DA As Integer
Dim l03DC As Integer
Dim l03DE As Integer
l03CA% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l03CA%, "Preferences")
Do: DoEvents
l03CC% = extfn0524(l03CA%, "MDIClient")
l03CE% = extfn0510(l03CC%, "Preferences")
l03D0% = extfn0510(l03CE%, "Mail")
l03D2% = extfn04C8(l03D0%, gc0084)
If  l03D0% <> 0 And l03D2% <> 0 Then Exit Do
Loop
sub00ED (.1#)
sub00A9 (l03D2%)
Do: DoEvents
l03D4% = extfn04BA("_AOL_Modal", "Mail Preferences")
l03D6% = extfn0510(l03D4%, "Close mail after it has been sent")
l03D8% = extfn0510(l03D4%, "Confirm mail after it has been sent")
l03DA% = extfn0510(l03D4%, "OK")
If  l03DA% <> 0 And l03D6% <> 0 And l03D8% <> 0 Then Exit Do
Loop
l03DC% = extfn0398(l03D6%, gc0074, 1, 0)
l03DC% = extfn0398(l03D8%, gc0074, 1, 0)
sub00A9 (l03DA%)
Do: DoEvents
l03D4% = extfn04BA("_AOL_Modal", "Mail Preferences")
Loop Until l03D4% = 0
l03DE% = extfn0398(l03CE%, gc006E, 0, 0)
End Sub


Sub sub024E (p03E0 As Variant)
Dim l03E4 As Variant
l03E4 = extfn0538(p03E0, gc00B2)
End Sub


Sub sub0259 (p03E8 As String)
Dim l03EA As Variant
Dim l03EE As Variant
Dim l03F2 As Integer
Dim l03F4 As Integer
Dim l03F6 As Integer
Dim l03F8 As Integer
Dim l03FA As Variant
Dim l03FE As Integer
Dim l0400 As Integer
Dim l0402 As Integer
Dim l0404 As Integer
Dim l0406 As Integer
Dim l0408 As Integer
Dim l040A As Integer
Dim l040C As Integer
Dim l040E As Variant
Dim l0412 As Integer
Dim l0414 As Integer
Dim l0416 As Integer
l03EA = " NOT DONE"
l03EE = "YET"
l03F2% = extfn04BA("AOL Frame25", 0&)
Call sub020F(l03F2%, "Compose Mail")
Do
DoEvents
l03F4% = extfn0524(l03F2%, "MDIClient")
l03F6% = extfn0510(l03F4%, "Compose Mail")
l03F8% = extfn0524(l03F6%, "_AOL_Edit")
Loop Until l03F8% <> 0
sub0269 (.6#)
l03FA = extfn02F3(l03F8%, gc0068, 0, p03E8)
Do
DoEvents
l03FE% = extfn0524(l03F2%, "MDIClient")
l0400% = extfn0510(l03FE%, "Compose Mail")
l0402% = extfn0524(l0400%, "_AOL_Edit")
Loop Until l0402% <> 0
Do
DoEvents
l0404% = extfn04C8(l0402%, gc0084)
l0406% = extfn04C8(l0404%, gc0084)
l0408% = extfn04C8(l0406%, gc0084)
l040A% = extfn04C8(l0408%, gc0084)
Loop Until l040A% <> 0
l03FA = extfn02F3(l040A%, gc0068, 0, "²º·‚¸‚•²º·‚¸‚• O.c.W InfoRmaTioN")
Do
DoEvents
l040C% = extfn0524(l03F6%, "RICHCNTL")
Loop Until l040C% <> 0
l040E = extfn02F3(l040C%, gc0068, 0, l03EA + l03EE)
Do
DoEvents
l0412% = extfn0524(l03F6%, "_AOL_ICON")
Loop Until l0412% <> 0
l0414% = extfn02DF(l0412%, gc0064, 0&, 0&)
l0414% = extfn02DF(l0412%, gc0066, 0&, 0&)
sub0269 (.6#)
l0416% = extfn0398(l03F6%, gc006E, 0, 0)
End Sub

Sub sub0263 (ByVal p0418 As String)
Dim l041A As Integer
Dim l041C As Integer
Dim l041E As Integer
Dim l0420 As Variant
l041A% = fn0101()
l041C% = extfn0524(l041A%, "_AOL_Edit")
l041E% = extfn02F3(l041C%, gc0068, 0, p0418$)
l0420 = extfn02DF(l041C%, gc005E, 13, 0)
l0420 = extfn02DF(l041C%, gc005E, 13, 0)
sub0269 (.0001#)
End Sub

Sub sub0269 (ByVal pv0424 As Variant)
Dim l0428 As Variant
l0428 = Timer
Do
 DoEvents
Loop Until ((Timer) >= (l0428 + pv0424))
End Sub

Sub sub0274 ()
Dim l042C As Integer
Dim l042E As Integer
Dim l0430 As Variant
Dim l0434 As Variant
Dim l0438 As Variant
Do
DoEvents
l042C% = extfn04BA("AOL Frame25", 0&)
l042E% = extfn0524(l042C%, "_AOL_Tree")
Loop Until l042E% <> 0
Do
DoEvents
l0430 = extfn0398(l042E%, gc0076, 0, 0&)
Call sub0269(1.5#)
l0434 = extfn0398(l042E%, gc0076, 0, 0&)
Call sub0269(1.5#)
l0438 = extfn0398(l042E%, gc0076, 0, 0&)
Loop Until l0430 = l0434 And l0434 = l0438
End Sub

Function fn0280 (p0440 As Integer) As Variant
Dim l0444 As String * 255
Dim l0446 As Integer
l0446 = extfn0491(p0440, l0444, 255)
fn0280 = (l0444)
End Function

