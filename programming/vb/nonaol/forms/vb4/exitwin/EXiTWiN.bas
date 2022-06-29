Attribute VB_Name = "EXiTWiN"
' Windows exit sample by NiVeK
' Please thank the cooL site that has my examples and proggs
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal DWreserved As Long)
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1




