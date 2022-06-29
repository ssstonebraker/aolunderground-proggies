Attribute VB_Name = "cpu_fire_up"
'gettickcount = gets the time that the cpu
'has been on, in milliseconds...
Public Declare Function GetTickCount Lib "kernel32" () As Long

'setwindowpos = setting the back to front
'position of a window, in this case we
'use the top most position, because well
'we want it to stay on top =)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2 'disable movement
Public Const SWP_NOSIZE = &H1 'disable resizing
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE ' set flags

Public Sub lastboot(hr As Label, min As Label)
 Dim lngHours As Long, lngMinutes As Long
 lngCount = GetTickCount 'boot time in milliseconds
 lngHours = ((lngCount / 1000) / 60) / 60 'full hours since boot
 lngMinutes = ((lngCount / 1000) / 60) Mod 60 'leftover minutes
 hr.Caption = lngHours ' set the label to # of hours
 min.Caption = lngMinutes 'set label to # of minutes
End Sub

Public Sub stayontop(f As Form)
 'i added this so the form will stay on top,
 'when you run the project, because its just
 'easier to see it that way when its on top
 Call SetWindowPos(f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
