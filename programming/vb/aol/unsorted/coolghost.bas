Attribute VB_Name = "coolghost"
Sub ghoston()
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aolicon As Long
Dim blwin As Long
Dim aolicons As Long
Dim aoltab As Long
Dim blpref As Long
Dim aoltabpage As Long
Dim aolradiobox As Long
Dim aolicont As Long
Dim getit As Long
On Error Resume Next
AOL4KW "buddy"
Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
ClickIt aolicon
    EnterKey aolicon
Loop Until aolicon <> 0
Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Loop Until aolicon <> 0
Do
blwin = AOLChildByTitle("Buddy List Setup")
Loop Until blwin <> 0
aolicons = FindWindowEx(blwin, 0&, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
ClickIt aolicons
Do
blpref = AOLChildByTitle("Buddy List Preferences")
Loop Until blpref <> 0
aoltab = FindWindowEx(blpref, 0&, "_aol_tabcontrol", vbNullString)
Call SendMessageLong(aoltab, WM_KEYDOWN, RIGHT_VKEY, 0&)
Call SendMessageLong(aoltab, WM_KEYUP, RIGHT_VKEY, 0&)
Call SendMessageLong(aoltab, WM_KEYDOWN, RIGHT_VKEY, 0&)
Call SendMessageLong(aoltab, WM_KEYUP, RIGHT_VKEY, 0&)

aoltab = FindWindowEx(blpref, 0&, "_aol_tabcontrol", vbNullString)
aoltabpage = FindWindowEx(aoltab, 0&, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltab, aoltabpage, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltab, aoltabpage, "_aol_tabpage", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, 0&, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
getit& = SendMessageByNum(aolradiobox, BM_GETCHECK, 0&, 0&)
        If getit& = 0 Then
     Call SendMessageLong(aolradiobox, BM_SETCHECK, True, 0&)
     End If
    
aolicont = FindWindowEx(blpref, 0&, "_aol_icon", vbNullString)
ClickIt aolicont
    EnterKey aolicon
Sleep 100&
Win_CloseWin blwin
End Sub
Sub ghostoff()
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aolicon As Long
Dim blwin As Long
Dim aolicons As Long
Dim aoltab As Long
Dim blpref As Long
Dim aoltabpage As Long
Dim aolradiobox As Long
Dim aolicont As Long
Dim getit As Long
On Error Resume Next
AOL4KW "buddy"
Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
ClickIt aolicon
    EnterKey aolicon
Loop Until aolicon <> 0
Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Loop Until aolicon <> 0
Do
blwin = AOLChildByTitle("Buddy List Setup")
Loop Until blwin <> 0
aolicons = FindWindowEx(blwin, 0&, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
aolicons = FindWindowEx(blwin, aolicons, "_aol_icon", vbNullString)
ClickIt aolicons
Do
blpref = AOLChildByTitle("Buddy List Preferences")
Loop Until blpref <> 0
aoltab = FindWindowEx(blpref, 0&, "_aol_tabcontrol", vbNullString)
Call SendMessageLong(aoltab, WM_KEYDOWN, RIGHT_VKEY, 0&)
Call SendMessageLong(aoltab, WM_KEYUP, RIGHT_VKEY, 0&)
Call SendMessageLong(aoltab, WM_KEYDOWN, RIGHT_VKEY, 0&)
Call SendMessageLong(aoltab, WM_KEYUP, RIGHT_VKEY, 0&)

aoltab = FindWindowEx(blpref, 0&, "_aol_tabcontrol", vbNullString)
aoltabpage = FindWindowEx(aoltab, 0&, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltab, aoltabpage, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltab, aoltabpage, "_aol_tabpage", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, 0&, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
getit& = SendMessageByNum(aolradiobox, BM_GETCHECK, 0&, 0&)
        If getit& = 0 Then
     Call SendMessageLong(aolradiobox, BM_SETCHECK, False, 0&)
     End If
    
aolicont = FindWindowEx(blpref, 0&, "_aol_icon", vbNullString)
ClickIt aolicont
    EnterKey aolicon
Sleep 100&
Win_CloseWin blwin
End Sub
