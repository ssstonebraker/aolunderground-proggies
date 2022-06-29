Attribute VB_Name = "upchat"
' Upchat.bas Verson 1.0 ( there probobly will not be any more versons)
'Well heres a lil Auto UpChat that I made. I really dont know
'why I was really bord.  I know that all of the sudden if I
'send this out to a bunch of people then this will be some
'prog for a dumb A$$ Programmer wanna be, that dosent even want
'to take some time to sit down and learn some API and simple
'Subs & Functions ( like I am doing my self) but any ways here
'its all yours! I thought the DogBar was a nice lil thing to
'include but hey I just made this so that I may share some of
'my hard wizdom with others. If you ever for any reason need to
'talk to me you can send me "1" count em "1" instant message
'if I am not bussy ill be glad to explane some thing  in depth.
'I am not like those untouchables =) like i think ive talked
'with a few "leeto" people like "Azar","Diox","GPX",and "Flow",
'"Chudd","Pat or JK"(I used his spy for this thanks man =D )
'on like 1 or 2 ocasions. By the way since I am rambling on
'here I would like to thank sonics new prog "Up Dat encode 5.0"
'for this consept=D - thanks "sonic" and well I love GPX's
'Bixtch, and programming bords there awesome!
'ps.   this is for AOL 4.0 =)

'################################################################################
'############*******************************************#########################
'############*              Contact 2ooo               *#########################
'############*        E-Mail: JCool710@aol.com         *#########################
'############*          Aim : IiIiBarcodeiIiI          *#########################
'############*      Web Site: www.jctech.simplenet.com *#########################
'############|********************____*****************|#########################
'#####********||||||||||||||||||||____||||||||||||||||*******####################
'#####*|||||||||||||||||||||||||||____||||||||||||||||||||||*####################
'################################################################################
Option Explicit
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wflags As Long) As Long

Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186


Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

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
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Sub StayOnTop(TheForm As Form)
Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub NotOnTop(TheForm As Form)
Call SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub


Public Function find_aolmodal() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim aolframe&
Dim aolmodal&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
Dim Winkid1&, Winkid2&, Winkid3&, Winkid4&, Winkid5&, FindOtherWin&
FindOtherWin& = GetWindow(aolmodal&, GW_HWNDFIRST)
Do While FindOtherWin& <> 0
       DoEvents
       Winkid1& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
       Winkid2& = FindWindowEx(FindOtherWin&, 0&, "_aol_gauge", vbNullString)
       Winkid3& = FindWindowEx(FindOtherWin&, 0&, "_aol_static", vbNullString)
       Winkid4& = FindWindowEx(FindOtherWin&, 0&, "_aol_checkbox", vbNullString)
       Winkid5& = FindWindowEx(FindOtherWin&, 0&, "_aol_button", vbNullString)
       If (Winkid1& <> 0) And (Winkid2& <> 0) And (Winkid3& <> 0) And (Winkid4& <> 0) And (Winkid5& <> 0) Then
              find_aolmodal = FindOtherWin&
              Exit Function
       End If
       FindOtherWin& = GetWindow(FindOtherWin&, GW_HWNDNEXT)
Loop
find_aolmodal = 0
' example on how to use:
' If find_aolmodal() <> 0 Then
' what to do if window is found
' Else
' what to do if window is not found
' End If
End Function
Sub GetWinTxt()
Dim aolframe&
Dim aolmodal&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(aolmodal&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(aolmodal&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
Form1.Label1.Caption = "" + Right(TheText$, 3)

End Sub
Sub disableuploadwin()
Dim aolframe&
Dim aolmodal&
Dim aolgauge&
Dim upp%
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("aol frame25", "_aol_modal")
aolgauge& = FindWindow("_aol_modal", "_AOL_Gauge")
If aolmodal& <> 0 Then upp% = aolmodal&
Call EnableWindow(aolframe&, 1)
Call EnableWindow(upp%, 0)
     Exit Sub


End Sub

Sub enableuploadwin()
Dim aolframe&
Dim aolmodal&
Dim aolgauge&
Dim upp%
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("aol frame25", "_aol_modal")
aolgauge& = FindWindow("_aol_modal", "_AOL_Gauge")
If aolgauge& <> 0 Then upp% = aolmodal&
Call EnableWindow(aolframe&, 1)
Call EnableWindow(upp%, 0)
  Form1.Text1.Text = ""
   Exit Sub
End Sub
Sub getstatictxt()
Dim TheText$
Dim TL As Long
Dim aolframe&
Dim aolmodal&
Dim aolstatic&

aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolstatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolmodal&, aolstatic&, "_aol_static", vbNullString)
TL = SendMessageLong(aolstatic&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(aolstatic&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
Form1.Text1.Text = TheText$

If aolstatic& = 0 Then
     Exit Sub
End If


End Sub


Sub miniuploadwin()
Dim aolframe&
Dim aolmodal&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
Call ShowWindow(aolmodal&, SW_MINIMIZE)
End Sub
Sub getfilename()
Dim aolframe&
Dim aolmodal&
Dim aolstatic&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolstatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(aolstatic&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(aolstatic&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
Form1.Text3.Text = Right(TheText$, 13)
End Sub
