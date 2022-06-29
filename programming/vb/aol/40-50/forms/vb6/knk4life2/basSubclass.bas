Attribute VB_Name = "basSubclass"
Option Explicit

'Variable to hold the address of the previous
'window procedure
Private g_OldWindowProc As Long

'Constant for use with GetWindowLong and SetWindowLong
Private Const GWL_WNDPROC = (-4)

'These three API functions are required for subclassing
Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) _
As Long

Private Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

'Constants for mouse messages
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209


Public Function MyWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'This is our application-defined procedure for processing messages.
'ALL messages sent to the subclassed window will first come to
'this procedure.  It is here that we process the message(s) we
'choose and send the rest on to Windows for default processing.

'This function returns 0 if our application-defined procedure
'processes the message.  If the message is sent on to Windows
'for default processing, the return value is the return value
'of the CallWindowProc function, which varies depending on the
'actual message.  Again, you should consult the SDK for information
'about particular messages.

Dim lRet As Long

'Determine the message that was received
Select Case uMsg
    Case TaskBar.CallbackMessage
        'Determine what mouse event occurred
        Select Case lParam
            'Left button double-click
            Case WM_LBUTTONDBLCLK
                  Beep
            'Right button up
            Case WM_RBUTTONUP
                If Form1.Visible Then Form1.Hide
                Form1.PopupMenu Form1.mnuPopUp, , , , Form1.max
        End Select
    'Case WM_OTHERMESSAGES
        'Add any additional case statements for other messages you
        'may want to process.
    Case Else
        'Message is not one we want to process.  Send it on to
        'Windows for default processing.  This is ABSOLUTELY
        'neccessary; otherwise, your application will NOT respond
        'Made By KnK
        'E-Mail me at Bill@knk.tierranet.com
        'This was DL from http://knk.tierranet.com/knk4o
        'to messages such as mouse clicks, key presses,
        'menu items being selected, or ANY other message that is
        'sent to the subclassed window.
        MyWindowProc = CallWindowProc(g_OldWindowProc, hwnd, uMsg, wParam, lParam)
End Select

End Function


Public Function Hook(Frm As Form) As Boolean

'Enables subclassing of the specified form.  The function
'returns True if successful; False, otherwise.
'The parameter Frm is the name of the form you want to
'subclass.

'Note:  While this example is set up to subclass forms only,
'it should be noted that it is possible to subclass ANY window.
'This could include text boxes, list boxes, command buttons, etc.
'since these are really just specialized windows (i.e. they have
'an hWnd property).

Dim lRet As Long

'Assign function a default return value of True
Hook = True

'Get the address for the previous window precedure
g_OldWindowProc = GetWindowLong(Frm.hwnd, GWL_WNDPROC)
If g_OldWindowProc = 0 Then
    Hook = False
    Exit Function
End If

'Set our application-defined function as the new window procedure
'This is what creates the subclass.  VB5's new AddressOf operator
'is what makes it possible to subclass.  It would be to your
'advantage to read the SDK on the SetWindowLong and
'CallWindowProc functions.  The SDK is included on the
'Developer Network CD-ROM included with all editions of VB5 except
'the Learning Edition.  Also, carefully read VB Help on the use of
'AddressOf.  Note particularly that the procedure used with
'AddressOf MUST be in a standard module (a .bas file).  The
'procedure CANNOT be located in form (.frm file) or
'class (.cls file) modules.
If SetWindowLong(Frm.hwnd, GWL_WNDPROC, AddressOf MyWindowProc) = 0 Then
    Hook = False
End If


End Function


Public Function Unhook(Frm As Form) As Boolean

'This function disables subclassing of the specified form.

Dim lRet As Long

Unhook = True

If SetWindowLong(Frm.hwnd, GWL_WNDPROC, g_OldWindowProc) = 0 Then
    Unhook = False
End If

End Function


