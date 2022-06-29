Attribute VB_Name = "mWinAPI"
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
ByVal Y3 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
(ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260
Private Const AMAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * AMAX_PATH
End Type

Public Enum FILE_ATTRIBUTES
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Public Type POINTAPI
    lxPos As Long
    lyPos As Long
End Type

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

