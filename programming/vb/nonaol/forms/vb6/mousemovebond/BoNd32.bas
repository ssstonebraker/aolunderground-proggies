Attribute VB_Name = "BoNd"
'What's up? This is the bas file for my Mouse Move example form.
'Below is the declaration for moving a form with no menu bar.
'Any questions or comments, please feel free to e-mail me at I be BoNd@aol.com

Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long


