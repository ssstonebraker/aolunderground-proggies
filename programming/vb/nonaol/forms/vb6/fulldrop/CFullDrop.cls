VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFullDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©1999 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 API Declarations
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const WM_CTLCOLORLISTBOX = &H134
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const SM_CYBORDER = 6

' Member variables
Private m_hWnd As Long

Private Sub Class_Terminate()
   ' Always unhook!
   If m_hWnd Then
      Call UnhookWindow(m_hWnd)
   End If
End Sub

Public Property Let hWnd(ByVal NewVal As Long)
   ' Unhook previous window if need be.
   If m_hWnd Then
      Call UnhookWindow(m_hWnd)
   End If

   ' Store handle and hook new window.
   If IsWindow(NewVal) Then
      m_hWnd = NewVal
      Call HookWindow(m_hWnd, Me)
   End If
End Property

Public Property Get hWnd() As Long
   ' Return handle for window we're monitoring.
   hWnd = m_hWnd
End Property

' **************************************************************
'  Subclassing a Form or UserControl
'    Copy the following, and add handling for additional
'    messages of interest.
' **************************************************************
Friend Function WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   Static rLst As RECT
   Static rCbo As RECT
   Static nHeight As Long
   Static nWidth As Long
   Static nCount As Long
   Dim Result As Long
   
   Select Case msg
      ' Add handlers here for each message you're interested in.
      
      Case WM_CTLCOLORLISTBOX
         ' get stats on dropdown list
         nHeight = SendMessage(lp, LB_GETITEMHEIGHT, 0, ByVal 0&)
         nCount = SendMessage(lp, LB_GETCOUNT, 0, ByVal 0&)
         Call GetWindowRect(lp, rLst)
         
         ' adjust and store height/width
         nHeight = (nCount * nHeight) + _
                   (GetSystemMetrics(SM_CYBORDER) * 2)
         nWidth = rLst.Right - rLst.Left
         
         ' get the combo's position
         Call GetWindowRect(hWnd, rCbo)
         
         ' check to make sure we're onscreen
         If rCbo.Bottom + nHeight > (Screen.Height / Screen.TwipsPerPixelY) Then
            ' need to "drop" above combo
            rLst.Top = rCbo.Top - nHeight
         End If
         
         ' move that sucker
         Call MoveWindow(lp, rLst.Left, rLst.Top, nWidth, nHeight, 1&)
         
         ' allow default processing
         Result = InvokeWindowProc(hWnd, msg, wp, lp)

      Case Else
         ' Pass along to default window procedure.
         Result = InvokeWindowProc(hWnd, msg, wp, lp)
   End Select

   ' Return desired result code to Windows.
   WindowProc = Result
End Function


