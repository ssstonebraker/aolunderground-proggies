Attribute VB_Name = "modIMEnter"
Option Explicit

'This module was made by Kyle LaDuke.
'It was made because im bored out of my mind right now.
'It was put together in ... i lost track of time but not to long
'To start it call HookKeyboard and to stop it call UnhookKeyboard

'The sad thing is i dont even use AOL anymore but yet im still making
'stuff for it sometimes.  Well i thought this could ease the lifes of
'many people so here you go.  The code was written by me .. all except
'the first section of API's.  I used API viewer to get those like we all
'do.  But I had to decare some of them myself.. those are the ones that
'follow the first section as i pointed out.  if you have anything you
'cant make but would like to see done send me your idea and ill do it
'for you then release the source.

' - Kyle LaDuke (ProgramDeveloper@hotmail.com)

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
'The following line is a modified version of CopyMemory
Private Declare Sub CopyStruct Lib "kernel32" Alias "RtlMoveMemory" (Struct As Any, ByVal ptr As Long, ByVal cbSize As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwthreadid As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Const HC_ACTION = 0

Private Const WH_KEYBOARD_LL = 13

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private lngHook As Long

'These are used in the hook call back for the keyboard.  I dim'ed 'em here so
'that they would only need to be done once for the many times the callback
'procedure will be called.
Dim kbhlParam As KBDLLHOOKSTRUCT, gtiInfo As GUITHREADINFO


' -- The following things weren't in API Viewer so I got the commented out --'
' -- code from MSDN and then converted it into Visual Basic code.  I       --'
' -- thought I 'd show you it just incase you wanted to see it for one     --'
' -- reason or another.                                                    --'

'BOOL WINAPI GetGUIThreadInfo(
'  DWORD idThread,       // thread identifier
'  PGUITHREADINFO lpgui  // thread information
');

Private Declare Function GetGUIThreadInfo Lib "user32" (ByVal idThread As Long, lpgui As GUITHREADINFO) As Long

'typedef struct tagGUITHREADINFO {
'    DWORD   cbSize;
'    DWORD   flags;
'    HWND    hwndActive;
'    HWND    hwndFocus;
'    HWND    hwndCapture;
'    HWND    hwndMenuOwner;
'    HWND    hwndMoveSize;
'    HWND    hwndCaret;
'    RECT    rcCaret;
'} GUITHREADINFO, *PGUITHREADINFO;

Private Type GUITHREADINFO
  cbSize As Long
  flags As Long
  hwndActive As Long
  hwndFocus As Long
  hwndCapture As Long
  hwndMenuOwner As Long
  hwndMoveSize As Long
  hwndCaret As Long
  rcCarect As RECT
End Type

'typedef struct tagKBDLLHOOKSTRUCT {
'    DWORD     vkCode;
'    DWORD     scanCode;
'    DWORD     flags;
'    DWORD     time;
'    ULONG_PTR dwExtraInfo;
'} KBDLLHOOKSTRUCT, *PKBDLLHOOKSTRUCT;

Private Type KBDLLHOOKSTRUCT
  vkCode As Long
  scanCode As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type

'Only procedures follow this line ----------------------------------------

'The parameters of this function are predefined.  They may not be changed.
Private Function KeyboardProc(ByVal idHook As Integer, ByVal wParam As Long, ByVal lParam As Long) As Long
  If idHook = HC_ACTION Then
  
    CopyStruct kbhlParam, lParam, Len(kbhlParam)

    If kbhlParam.vkCode = vbKeyReturn Then
      gtiInfo.cbSize = Len(gtiInfo)
      
      GetGUIThreadInfo 0&, gtiInfo
      
      If IsIM(gtiInfo.hwndFocus) = True Then
        If wParam = WM_KEYUP Then _
          ClickSend gtiInfo.hwndFocus
          
        KeyboardProc = 1
        Exit Function
      End If
  
    End If
  End If
  
  KeyboardProc = CallNextHookEx(lngHook, idHook, wParam, lParam)
End Function

Public Sub HookKeyboard()
  lngHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardProc, App.hInstance, 0&)
End Sub

Public Sub UnhookKeyboard()
  UnhookWindowsHookEx lngHook
End Sub

Private Function IsIM(lngWindowHandle) As Boolean
  Dim lngWindow As Long, lngAOLFrame As Long, lngAOLMDI As Long, lngContainer As Long, strCaption As String
  
  lngWindow = GetParent(lngWindowHandle)
  lngAOLFrame = FindWindow("AOL Frame25", vbNullString)
  lngAOLMDI = FindWindowEx(lngAOLFrame, 0&, "MDIClient", vbNullString)
  lngContainer = GetParent(lngWindow)

  strCaption = Space(25)
  
  GetWindowText lngWindow, strCaption, 25
  
  If Left(strCaption, Len("Send Instant Message")) = "Send Instant Message" Then
    IsIM = True
  End If
  
  If Left(strCaption, Len(">IM From:")) = ">IM From:" Then
    IsIM = True
  End If
  
  If Left(strCaption, Len(" IM To:")) = " IM To:" Then
    IsIM = True
  End If

End Function

Private Sub ClickSend(lngWindowHandle As Long)
  Dim intCurrent As Integer, lngParentWindow As Long, lngAOLIcon As Long
  
  lngParentWindow = GetParent(lngWindowHandle)
  
  lngAOLIcon = FindWindowEx(lngParentWindow, 0&, "_AOL_Icon", "")

  For intCurrent = 0 To 8
    lngAOLIcon = FindWindowEx(lngParentWindow, lngAOLIcon, "_AOL_Icon", "")
  Next intCurrent
   
  SendMessage lngAOLIcon, WM_LBUTTONDOWN, 0&, 0&
  SendMessage lngAOLIcon, WM_LBUTTONUP, 0&, 0&
  SendMessage lngAOLIcon, WM_KEYDOWN, vbKeySpace, 0&
  SendMessage lngAOLIcon, WM_KEYUP, vbKeySpace, 0&

End Sub
