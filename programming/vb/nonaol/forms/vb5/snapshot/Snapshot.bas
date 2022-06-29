Attribute VB_Name = "Snapshot"

Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const VK_MENU = &H12
Const VK_SNAPSHOT = &H2C
Const KEYEVENTF_KEYUP = &H2
Sub snap1()
Dim frm As New Snap
Dim frmm As New main
' Presses Alt.
    keybd_event VK_MENU, 0, 0, 0
    DoEvents
    
    ' Presses Print Scrn.
    keybd_event VK_SNAPSHOT, 1, 0, 0
    DoEvents
        'this is used to copy the form only
        'keybd_event VK_SNAPSHOT, 0, 0, 0
        'DoEvents
    ' Releases Alt.
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    DoEvents

    frm.Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    frm.Show
    frm.WindowState = vbMaximized
    frmm.WindowState = vbMinimized
End Sub
