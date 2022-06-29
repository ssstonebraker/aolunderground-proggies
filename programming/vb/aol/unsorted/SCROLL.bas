' SCROLL.FRM
Option Explicit

Sub done_Click ()
Scroll.Enabled = False
Form9.Hide
Form3.Show
End Sub

Sub Form_Load ()
extsubAC18 Form9.hWnd, -1, 0, 0, 0, 0, &H50
End Sub

Sub Scroll_Click ()
sub4680 "•·´¯`·.¸.·• AoSkam Scroller Ac†ivateD •·´¯`·.¸.·•"
subB1C8 (.0001#)
sub4680 "" + Text1 + ""
subB1C8 (.0001#)
sub4680 "" + Text2 + ""
subB1C8 (.0001#)
sub4680 "" + Text3 + ""
subB1C8 (2#)
sub4680 "" + Text4 + ""
subB1C8 (.0001#)
sub4680 "" + Text5 + ""
subB1C8 (.0001#)
sub4680 "" + Text6 + ""
subB1C8 (.0001#)
sub4680 "•·´¯`·.¸.·• AoSkam Scroller DeAc†ivateD •·´¯`·.¸.·•"
End Sub

Sub Stop_Click ()
Scroll.Enabled = False
End Sub
