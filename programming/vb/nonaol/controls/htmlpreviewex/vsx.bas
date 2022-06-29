Attribute VB_Name = "vsx"
'vsx.bas by vsx
'ues the form included for an example
'http://www.bigdazz.com

'--------
'form moveform
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012

'---------
'for formontop
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long







































Public Const HWND_TOPMOST = -1
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type

























Public Sub formontop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub moveform(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
'place in FORM_RESIZE (form_load wont work)
'form3d me
Sub form3d(frmForm As Form)
Const cPi = 3.1415926
       Dim intLineWidth As Integer
       intLineWidth = 2
       Dim intSaveScaleMode As Integer
       intSaveScaleMode = frmForm.ScaleMode
       frmForm.ScaleMode = 3
       Dim intScaleWidth As Integer
       Dim intScaleHeight As Integer
       intScaleWidth = frmForm.ScaleWidth
       intScaleHeight = frmForm.ScaleHeight
       frmForm.Cls
       frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
       frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
       frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
       frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
       Dim intCircleWidth As Integer
       intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
       frmForm.FillStyle = 0
       frmForm.FillColor = QBColor(15)
       frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), -3.1415926, -3.90953745777778
       frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), -0.78539815, -1.5707963
       frmForm.Line (0, intScaleHeight)-(0, 0), 0
       frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
       frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
       frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
       frmForm.ScaleMode = intSaveScaleMode
End Sub




