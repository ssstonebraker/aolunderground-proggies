Attribute VB_Name = "Module1"
Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Declare Function SetWindowPos Lib "User32" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Sub AlwaysOnTop(frmID As Form, OnTop As Integer)
' Pass any non-zero value to Place on top
' Pass zero to remove top-mostness
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop Then
        OnTop = SetWindowPos(frmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(frmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
    
    End Sub


