Attribute VB_Name = "Module1"
' Written By G man
' IHSDOGG@excite.com
' The codes on the subs don't exactly work and
' I can't figure out why.  This will work on the
' other hand if you copy the code from the sub and
' paste it into the correct area of your project.
' If you can get the subs to work email me the fix
' and I'll put you in a fix version 2 of this .bas if
' I ever do one.
'
' To put your icon in the system tray you need 3
' things.  2 command buttons (if you choose to
' have the user put them in or take em out)  and
' a picturebox with an icon as the picture.
' Copy the whole sub to create and delete the icons
' into their buttons, and copy the PictureMouseMove
' code into the pictures mousemove proc.  Change
' all necessary items to fit your project.
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
"Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
NOTIFYICONDATA) As Long
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Sub DeleteSystrayicon(picture)
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = Picture1.hwnd '<---- change
Tic.uID = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub
Private Sub PictureMouseMove()
X = X / Screen.TwipsPerPixelX
Select Case X
Case WM_RBUTTONDOWN
Caption = "right"
Case WM_MOUSEMOVE
Caption = "mouse"
End Select
End Sub


Private Sub CreateSystrayicon()
Dim picture As Long
Dim message As Long
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = Picture1.hwnd  '<---- change the number
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = Picture1.picture '<---- change
Tic.szTip = "CHANGE ME" & Chr$(0)  '<---- change
erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub


