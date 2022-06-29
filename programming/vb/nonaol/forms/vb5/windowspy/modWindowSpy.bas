Attribute VB_Name = "modWindowSpy"
'modWindowSpy:                                                             '_'_'_'
'Saiñt,                                                                  _/__   __\_
'   I Didn't Write This Myself In Fact I Don't Even Know Who Wrote This {~ <@> <@> ~}   (mE)
'I Am Just Make It So You Can Learn What Going On And How This All       \_  < >  _/
'Works.  I Am No Expert By No Means So If Something Is Wrong Then That     \ -_- /
'Is Why.                                                              __ __|\_'_/|__ __
'Send All Comments Or Question To (V)E And If I Made Any Mistakes    /  |  \\___//  |  \
'Then Send Them Too I Would Like To Learn To.                       |   |   \   /   |   |
'                                                                   |   |    \_/    |   |
'                                                                   |   |   Saiñt   |   |
'
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer

Public Const HWND_TOPMOST = -1

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Type POINTAPI
   X As Long
   Y As Long
End Type

Function WindowSPY(WinHdl As Label, WinClass As Label, WinTxt As Label, WinStyle As Label, WinIDNum As Label, WinPHandle As Label, WinPText As Label, WinPClass As Label, WinModule As Label)

Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, R As Long
Static hWndLast As Long
    Call GetCursorPos(pt32) 'Find Where The Mouse Is
    ptx = pt32.X    'Is The X Coordinate
    pty = pt32.Y    'Is The Y Coordinate
    hWndOver = WindowFromPointXY(ptx, pty) '= The Windows Handle Mouse Is On
    If hWndOver <> hWndLast Then 'If This Windows Handle Is The Same As The Last Windows Handle
        hWndLast = hWndOver 'Windows Handle Last = Windows Handle The One The Mouse Is On
        WinHdl.Caption = " Window Handle: " & hWndOver  'WinHdl(the Label) Caption = The Windows Handle
        R = GetWindowText(hWndOver, sWindowText, 100) 'Finds The Windows Text
        WinTxt.Caption = " Window Text: " & Left(sWindowText, R)    'WinTxt(the Label) Caption = The Windows Text
        R = GetClassName(hWndOver, sClassName, 100) 'Finds The Windows Class Name
        WinClass.Caption = " Window Class Name: " & Left(sClassName, R) 'WinClass(the Label) Caption = The Windows Class Name
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)   'Finds The Windows Style
        WinStyle.Caption = " Window Style: " & lWindowStyle 'WinStyle(the Label) Caption = The Windows Style
        hWndParent = GetParent(hWndOver)    'Finds the Windows Parent (Where Is My Mommy)
            If hWndParent <> 0 Then 'If Windows Parent = anything then
                wID = GetWindowWord(hWndOver, GWW_ID)   'Finds Windows ID Number
                WinIDNum.Caption = " Window ID Number: " & wID  'WinIDNum(the Label) Caption = The Windows ID Number
                WinPHandle.Caption = " Parent Window Handle: " & hWndParent 'WinPHandle(the Label) Caption = The Windows Parent Handle
                R = GetWindowText(hWndParent, sParentWindowText, 100)   'Finds Windows Parents Text
                WinPText.Caption = " Parent Window Text: " & Left(sParentWindowText, R) 'WinPText(the Label) Caption = The Windows Parent Window Text
                R = GetClassName(hWndParent, sParentClassName, 100) 'Finds The Windows Parents Class Name
                WinPClass.Caption = " Parent Window Class Name: " & Left(sParentClassName, R)   'WinPClass(the Label) Caption = The Windows Parents Class Name
            Else    'Else
                WinIDNum.Caption = " Window ID Number: N/A" 'If Window Dosen't Have A Parent Then It dose Not Apply
                WinPHandle.Caption = " Parent Window Handle: N/A"   'Ditto
                WinPText.Caption = " Parent Window Text: N/A"   'Ditto
                WinPClass.Caption = " Parent Window Class Name : N/A"   'Ditto
            End If  'End The If
        hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)  'Finds The Windows ID Number
        R = GetModuleFileName(hInstance, sModuleFileName, 100)  'Finds The Windows Module Name
        WinModule.Caption = " Module: " & Left(sModuleFileName, R) 'WinModule(the Label) Caption = The Windows Module Name
    End If  'End The If
End Function    'Well That Is It - Latez
