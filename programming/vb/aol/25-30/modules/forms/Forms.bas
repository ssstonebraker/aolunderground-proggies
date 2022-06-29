'****************************************************
'* FORMS.BAS Version 2.1 Date: 06/01/95             *
'* VB Tips & Tricks                                 *
'* 8430-D Summerdale Road San Diego CA 92126-5415   *
'* Compuserve: 74227,1557                           *
'* America On-Line: DPMCS                           *
'* InterNet: DPMCS@AOL.COM                          *
'*==================================================*
'*This module contains common functions for use with*
'*forms.                                            *
'****************************************************
Option Explicit

Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

'*******************************************************
'* Procedure Name: KeepOnTop                           *
'*-----------------------------------------------------*
'* Created: 4/18/94   By: KeepOnTop                    *
'* Modified:          By:                              *
'*=====================================================*
'*Keep form on top. Note that this is switched off if  *
'*form is minimised, so place in resize event as well. *
'*******************************************************
Sub keepontop (frmIn As Form)
'Keep form on top. Note that this is switched off if form is
'minimised, so place in resize event as well.
Const wFlags = SWP_NOMOVE Or SWP_NOSIZE

    SetWindowPos frmIn.hWnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags    'Window will stay on top

    DoEvents

End Sub

'*******************************************************
'* Procedure Name: PaintBackGround                     *
'*-----------------------------------------------------*
'* Created:           By: KARL M. GARAND               *
'* Modified: 3/01/95  By: David McCarter               *
'*=====================================================*
'*This code paint the backgound of a form from black   *
'*to blue.                                             *
'*******************************************************
Sub PaintBackGround (frmIn As Form)
Dim I As Integer
Dim Y As Integer

    frmIn.AutoRedraw = True
    frmIn.DrawStyle = 6
    frmIn.DrawMode = 13
    frmIn.DrawWidth = 2
    frmIn.ScaleMode = 3
    frmIn.ScaleHeight = (256 * 2)
    For I = 0 To 255
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, 0, I), BF
        Y = Y + 2
    Next I

End Sub

'*******************************************************
'* Procedure Name: RemoveOnTop                         *
'*-----------------------------------------------------*
'* Created: 4/18/94   By:                              *
'* Modified:          By:                              *
'*=====================================================*
'*Removes the form from being on top.                  *
'*******************************************************
Sub RemoveOnTop (frmIn As Form)
Const wFlags = SWP_NOMOVE Or SWP_NOSIZE

    SetWindowPos frmIn.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags

    DoEvents

End Sub

