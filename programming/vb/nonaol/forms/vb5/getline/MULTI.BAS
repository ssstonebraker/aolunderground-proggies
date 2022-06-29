Attribute VB_Name = "DeclarationsForProject"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA

