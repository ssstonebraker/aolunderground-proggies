'User and GDI Functions for Explode to work
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As RECT)
Declare Function GetDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
Declare Sub SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)

Sub CenterForm (frm As Form)
    Dim X, Y                   ' New top, left for the form
    X = (Screen.Width - frm.Width) / 2
    Y = (Screen.Height - frm.Height) / 2
    frm.Move X, Y             ' Change location of the form
End Sub

Sub Explode (frm As Form, CFlag As Integer)
Const STEPS = 150 'Lower Number Draws Faster, Higher Number Slower
Dim FRect As RECT
Dim FWidth, FHeight As Integer
Dim I, X, Y, Cx, Cy As Integer
Dim hScreen, Brush As Integer, OldBrush

' If CFlag = True, then explode from center of form, otherwise
' explode from upper left corner.
    GetWindowRect frm.hWnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    FHeight = FRect.Bottom - FRect.Top
    
' Create brush with Form's background color.
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    
' Draw rectangles in larger sizes filling in the area to be occupied
' by the form.
    For I = 1 To STEPS
        Cx = FWidth * (I / STEPS)
        Cy = FHeight * (I / STEPS)
        If CFlag Then
            X = FRect.Left + (FWidth - Cx) / 2
            Y = FRect.Top + (FHeight - Cy) / 2
        Else
            X = FRect.Left
            Y = FRect.Top
        End If
        Rectangle hScreen, X, Y, X + Cx, Y + Cy
    Next I
    
' Release the device context to free memory.
' Make the Form visible

    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    frm.Show 1

End Sub

