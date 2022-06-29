Attribute VB_Name = "Position"
'This .bas provides the best way to position your
'forms relatively with AOL.  Using it, you can make
'forms non-intrusive to the users workspace.  I
'feel the best way to make an AOL Add-on is to make
'a toolbar that takes up a small amount of space,
'and then position it out of the way using functions
'in this .bas. The AOL4_ChatPosition is also good
'when used with a timer to make it follow the
'chat room.
'                           You Savior,
'                           ChiChis
'
'VB Version - VB 4.0 32 bit or higher
'AOL Version - Any AOL 4.0
'E-mail - MrChiChis@AOL.Com

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Rct As RECT
Sub AOL4_GrayBackBottom(frm As Form, FormHeight)
'This positions your form at the very bottom of AOL
'(right above the start bar)
Dim wndRect As RECT, lRet As Long
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
lRet = GetWindowRect(MDI%, wndRect)
frm.Top = wndRect.Bottom * Screen.TwipsPerPixelY - FormHeight
End Sub

Sub AOL4_ChatPosition(frm As Form)
'This positions a form's left and top exactly above
'the AOL chat text box.  Best used when there is a
'textbox with the font "Arial" font size set at 10
'and width set at 5295 (twips) in the upper left
'corner of your form.
'Warning: If the chatroom is not open you will not
'be able to see the form, because it can't find
'the room.
'Solution: Make an error message come up if
'AOL4_FindChat returns 0
Dim wndRect As RECT, lRet As Long
lRet = GetWindowRect(AOL4_FindRoom, wndRect)
With frm
  .Top = ((wndRect.Top + (wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY) - 650
  .Left = wndRect.Left * Screen.TwipsPerPixelX + 170
End With
End Sub
Function AOL4_FindChat()
'Useful for AOL4_ChatPosition me and other stuff
'such as chat sends (not included in this .bas)
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
firs% = Getwindow(MDI%, 5)
LISTERS% = FindChildByClass(firs%, "RICHCNTL")
LISTERE% = FindChildByClass(firs%, "RICHCNTL")
LISTERB% = FindChildByClass(firs%, "_AOL_Listbox")
Do While (LISTERS% = 0 Or LISTERE% = 0 Or LISTERB% = 0) And (L <> 100)
    DoEvents
    firs% = Getwindow(firs%, 2)
    LISTERS% = FindChildByClass(firs%, "RICHCNTL")
    LISTERE% = FindChildByClass(firs%, "RICHCNTL")
    LISTERB% = FindChildByClass(firs%, "_AOL_Listbox")
    If LISTERS% And LISTERE% And LISTERB% Then Exit Do
    L = L + 1
Loop
If (L < 100) Then
    AOL4_FindChat = firs%
    Exit Function
End If
AOL4_FindChat = 0
End Function
Function FindChildByClass(parentw, childhand)
firs% = Getwindow(parentw, GW_MAX)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = Getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
While firs%
firss% = Getwindow(parentw, GW_MAX)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = Getwindow(firs%, GW_HWNDNEXT)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
Wend
FindChildByClass = 0
Greed:
room% = firs%
FindChildByClass = room%
End Function
Sub AOL4_GrayBackTop(frm As Form)
'This positions your form right below the
'AOL toolbar
Dim wndRect As RECT, lRet As Long
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
lRet = GetWindowRect(MDI%, wndRect)
frm.Top = wndRect.Top * Screen.TwipsPerPixelY
End Sub
