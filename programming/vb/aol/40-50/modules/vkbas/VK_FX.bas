Attribute VB_Name = "VK_FX"

'this version  9-1-98

'coded by KRhyME and SkaFia
'|¯|    |¯|\¯\ |¯|    |_|/¯//¯/|¯|\¯\ /¯/|¯|
'| |/¯/ | |/ / | |\¯\   / / | ||_|| ||  |/_/_
'|_|\_\ |_|\_\ |_| |_| /_/  |_|   |_| \______|
'HEAD of the [Voltron Kru]
'Voltron Kru '98
'www.voltronkru.com
'voltronkru@juno.com

'This Bas file requires Voltron.bas to work
'Voltron.Bas is the core bas file for the
'Voltron Kru. You can get all our bas files at
'www.voltronkru.com

'many ideas for this bas came from other Voltron Kru
'members. I would Like to thank SkaFia for all the
'things he did for the series of Bas files.
'Please do not steal our codes without giving us
'credit. I would like to say thank you to KnK for
'making so many files avaible to the public, The makers
'of DiVe32.bas (the first bas i used), Toast, Magus,
'and all the other great programmers out there who
'have infuinced us

'Please join our VB mailing list
'www.voltronkru.com


Option Explicit
Dim nXCoord(50) As Integer
Dim nYCoord(50) As Integer
Dim nXSpeed(50) As Integer
Dim nYSpeed(50) As Integer

Sub HowTo_CircleForm()

'THIS CODE WILL MAKE CIRCLE OR OVEL SHAPED FORMS
'PLACE IN THE FORM LOAD
'SetWindowRgn hwnd, _
'  CreateEllipticRgn(0, 0, 300, 200), True
'
End Sub


Sub HowTO_StarField()
'this is how to make a star field
'on a form.....just uncomment, and edit


'Private Sub Form_Load()
'    Dim nIndex As Integer
'    ' At form load, the initial coordinates of all the stars needs to be set to off screen.
'    ' The timer event will recognise this and bring the stars back on screen.
'    For nIndex = 0 To 49
'        nXCoord(nIndex) = -1
'        nYCoord(nIndex) = -1
'    Next
'    ' Call the randomize method to tell VB to get ready to think of some random numbers
'    Randomize
'    Timer1.Enabled = True
'End Sub
'

'Private Sub Timer1_Timer()
'    ' The timer event performs three functions here.
'    '       1. Stars that are off screen are remade at the centre of the screen
'    '       2. Stars previously drawn are erase by redrawing them in black
'    '       3. Each star's position is recalculated and the star redrawn.
'    Dim nIndex As Integer
'    For nIndex = 0 To 49
'        'erase the previously drawn star
'        PSet (nXCoord(nIndex), nYCoord(nIndex)), &H0&
'        ' If the star number nIndex is off screen, then bring it back
'        If nXCoord(nIndex) < 0 Or nXCoord(nIndex) > frmMain.ScaleWidth Or nYCoord(nIndex) < 0 Or nYCoord(nIndex) > frmMain.ScaleHeight Then
'            nXCoord(nIndex) = frmMain.ScaleWidth \ 2
'            nYCoord(nIndex) = frmMain.ScaleHeight \ 2
'            ' Decide on some random speeds for the new star
'            nXSpeed(nIndex) = Int(Rnd(1) * 200) - 100   ' Gives a speed between -100 and 100
'            nYSpeed(nIndex) = Int(Rnd(1) * 200) - 100   ' Gives a speed between -100 and 100
'        End If
'        ' Now redraw the star so that it appears to move
'        nXCoord(nIndex) = nXCoord(nIndex) + nXSpeed(nIndex)
'        nYCoord(nIndex) = nYCoord(nIndex) + nYSpeed(nIndex)
'        PSet (nXCoord(nIndex), nYCoord(nIndex)), &HFFFFFF
'    ' Move on to the next star
'    Next
'End Sub
End Sub

Sub Form_Explode(F As Form, Movement As Integer)
'this will explode a form
    Dim myRect As rect
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect F.hWnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(F.BackColor)
    
    For i = 1 To Movement
        cx = formWidth * (i / Movement)
        cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub


Public Sub Form_Implode(F As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'The larger the "Movement" value, the slower the "Implosion"
    Dim myRect As rect
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect F.hWnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(F.BackColor)
    
        For i = Movement To 1 Step -1
        cx = formWidth * (i / Movement)
        cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub


Sub Form_ScrollDown(frm As Form, startNUM, endNUM)
'This will make the form slowly scroll down
'You can use a timeout to stop it and put it in a
'timer
Dim X
Dim Y
frm.Show
frm.Height = startNUM
X = frm.Height
For Y = X To endNUM
frm.Height = frm.Height + 20
timeout (0.0001)
If frm.Height >= endNUM Then GoTo out:
Next Y
out:
End Sub

Sub Form_ScrollUp(frm As Form, startNUM, endNUM)
'This will make the form slowly scroll up
'You can use a timeout to stop it and put it in a
'timer
Dim X
Dim Y
frm.Show
frm.Height = startNUM
X = frm.Height
For Y = X To endNUM
frm.Height = frm.Height - 20
timeout (0.0001)
'If frm.Height <= endNUM Then GoTo out:
Next Y
out:
End Sub

Sub Form_Suckin(frm As Form)

Do
DoEvents
frm.Height = frm.Height - 50
frm.Width = frm.Width - 50
Loop Until frm.Height < 450 And frm.Width < 1700
End Sub

Sub FormExit_Down(winform As Form)
'the form files down off the screen, and
'ends the program
Do
winform.Top = Trim(Str(Int(winform.Top) + 300))
DoEvents
Loop Until winform.Top > 7200
If winform.Top > 7200 Then End
End Sub


Sub FormExit_Left(winform As Form)
'the form files left off the screen, and
'ends the program
Do
winform.Left = Trim(Str(Int(winform.Left) - 300))
DoEvents
Loop Until winform.Left < -6300
If winform.Left < -6300 Then End
End Sub


Sub FormExit_right(winform As Form)
'the form files right off the screen, and
'ends the program
Do
winform.Left = Trim(Str(Int(winform.Left) + 300))
DoEvents
Loop Until winform.Left > 9600
If winform.Left > 9600 Then End
End Sub


Sub FormExit_up(winform As Form)
'the form files up off the screen, and
'ends the program
Do
winform.Top = Trim(Str(Int(winform.Top) - 300))
DoEvents
Loop Until winform.Top < -4500
If winform.Top < -4500 Then End
End Sub

Sub FormDraw3DBorder(F As Form)
'adds a 3d border to a form

Dim iOldScaleMode As Integer
Dim iOldDrawWidth As Integer
    iOldScaleMode = F.ScaleMode
    iOldDrawWidth = F.DrawWidth
    F.ScaleMode = vbPixels
    F.DrawWidth = 1
    F.Line (0, 0)-(F.ScaleWidth, 0), QBColor(15)
    F.Line (0, 0)-(0, F.ScaleHeight), QBColor(15)
    F.Line (0, F.ScaleHeight - 1)-(F.ScaleWidth - 1, F.ScaleHeight - 1), QBColor(8)
    F.Line (F.ScaleWidth - 1, 0)-(F.ScaleWidth - 1, F.ScaleHeight), QBColor(8)

    F.ScaleMode = iOldScaleMode
    F.DrawWidth = iOldDrawWidth
End Sub

Public Sub Form_MakeTransparent(frm As Form)
'makes a form transparent.....
       Dim rctClient As rect, rctFrame As rect
       Dim hClient As Long, hFrame As Long
       '     '// Grab client area and frame area
       GetWindowRect frm.hWnd, rctFrame
       GetClientRect frm.hWnd, rctClient
       '     '// Convert client coordinates to screen coordinates
       Dim lpTL As POINTAPI, lpBR As POINTAPI
       lpTL.X = rctFrame.Left
       lpTL.Y = rctFrame.Top
       lpBR.X = rctFrame.Right
       lpBR.Y = rctFrame.Bottom
       ScreenToClient frm.hWnd, lpTL
       ScreenToClient frm.hWnd, lpBR
       rctFrame.Left = lpTL.X
       rctFrame.Top = lpTL.Y
       rctFrame.Right = lpBR.X
       rctFrame.Bottom = lpBR.Y
       rctClient.Left = Abs(rctFrame.Left)
       rctClient.Top = Abs(rctFrame.Top)
       rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
       rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
       rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
       rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
       rctFrame.Top = 0
       rctFrame.Left = 0
       '     '// Convert RECT structures to region handles
       hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
       hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
       '     '// Create the new "Transparent" region
       CombineRgn hFrame, hClient, hFrame, RGN_XOR
       '     '// Now lock the window's area to this created region
       SetWindowRgn frm.hWnd, hFrame, True
End Sub

Public Sub Form_Move(TheForm As Form)
'WILL HELP YOU MOVE A FORM WITHOUT
'A TITLE BAR, PLACE IN MOUSEDOWN
       ReleaseCapture
       Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
       '
End Sub

Sub shrink(label As label, startSIZE, endSIZE)
'this function makes the text in a lebel shrink.
'the startSIZE has to be greater that the endSize.
'if the endSIZE is 0, the label dispears

label.Visible = True
label.FontSize = startSIZE
Dim X
Do
X = label.FontSize - 2
label.FontSize = X
    timeout (0.001)
    If label.FontSize < (endSIZE Or 2) Then Exit Do
Loop
If endSIZE = 0 Then
    label.Visible = False
End If
End Sub


Sub Grow(label As label, startSIZE, endSIZE)
'this function makes a label grow
'the font size starts a startSIZE,
'and ends and endSIZE
'MADE SURE THE FONT OF THE LABEL
'CAN GO BIG ENOUGH
label.Visible = True
label.FontSize = startSIZE
Dim X
Do While label.FontSize < endSIZE
    label.FontSize = label.FontSize + 2
        timeout (0.001)
    
Loop
End Sub

Sub bounce(label As label, minSIZE, MAXSIZE, numOFbounces)
'this function makes a label, grow, and shrink
'giving it a "BOUNCING" effect
'minSIZE is the smallest it goes, maxSIZE is the largest the label goes
'numOFbounces is the number of times the label bounces
'to make a label bounce forever, call this
'function in a timer, and have numOFbounces = 1
'MADE SURE THE FONT OF THE LABEL
'CAN GO BIG ENOUGH

label.FontSize = minSIZE
Dim X
Dim Y
Dim num
Start:
If (num >= numOFbounces) Then GoTo out:
Do
X = label.FontSize + 2
label.FontSize = X
    timeout (0.001)
    If label.FontSize >= MAXSIZE Then Exit Do
Loop

Do
X = label.FontSize - 2
label.FontSize = X
    timeout (0.001)
    If label.FontSize < (minSIZE Or 2) Then Exit Do
Loop
num = num + 1
GoTo Start:
out:
End Sub

Sub bounce2(label As label, big_size, small_size, numOFbounces)
'this is another bounce function
'Makesure the labels font will support
'the font sizes first
Dim X
For X = 1 To numOFbounces
Call Grow(label, small_size, big_size)
Call shrink(label, big_size, small_size)
Next X
End Sub

Sub FlipPictureHorizontal(pic1 As PictureBox, pic2 As PictureBox)
'pic1 = the existing pic
'pic2 = the pic to be fliped
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Dim px%
    Dim py%
    Dim retval%
    px% = pic1.ScaleWidth
    py% = pic1.ScaleHeight
    retval% = StretchBlt(pic2.hdc, px%, 0, -px%, py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub

Sub FlipPictureVertical(pic1 As PictureBox, pic2 As PictureBox)
'pic1 = the existing pic
'pic2 = the pic to be fliped
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Dim px%
    Dim py%
    Dim retval%
    px% = pic1.ScaleWidth
    py% = pic1.ScaleHeight
    retval% = StretchBlt(pic2.hdc, 0, py%, px%, -py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub

Sub PicRotate45(pic1 As PictureBox, pic2 As PictureBox)
'rotate 45 degrees
'pic1 = the existing pic
'pic2 = the pic to be rotated
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Call bmp_rotate(pic1, pic2, 3.14 / 4)
               End Sub


               Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta!)
                ' bmp_rotate(pic1, pic2, theta)
                ' Rotate the image in a picture box.
                '   pic1 is the picture box with the bitmap to rotate
                '   pic2 is the picture box to receive the rotated bitmap
                '   theta is the angle of rotation
                Dim c1x As Integer, c1y As Integer
                Dim c2x As Integer, c2y As Integer
                Dim a As Single
                Dim p1x As Integer, p1y As Integer
                Dim p2x As Integer, p2y As Integer
                Dim n As Integer, r   As Integer

                c1x = pic1.ScaleWidth \ 2
                c1y = pic1.ScaleHeight \ 2
                c2x = pic2.ScaleWidth \ 2
                c2y = pic2.ScaleHeight \ 2

                If c2x < c2y Then n = c2y Else n = c2x
                n = n - 1
                Dim pic1hdc%
                Dim pic2hdc%
                pic1hdc% = pic1.hdc
                pic2hdc% = pic2.hdc

                For p2x = 0 To n
                  For p2y = 0 To n
                    If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
                    r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
                    p1x = r * Cos(a + theta!)
                    p1y = r * Sin(a + theta!)
                    
                    Dim c0&
                    Dim c1&
                    Dim c2&
                    Dim c3&
                    Dim xret&
                    Dim t%
                    
                    c0& = GetPixel(pic1hdc%, c1x + p1x, c1y + p1y)
                    c1& = GetPixel(pic1hdc%, c1x - p1x, c1y - p1y)
                    c2& = GetPixel(pic1hdc%, c1x + p1y, c1y - p1x)
                    c3& = GetPixel(pic1hdc%, c1x - p1y, c1y + p1x)
                    If c0& <> -1 Then xret& = SetPixel(pic2hdc%, c2x + p2x, c2y + p2y, c0&)
                    If c1& <> -1 Then xret& = SetPixel(pic2hdc%, c2x - p2x, c2y - p2y, c1&)
                    If c2& <> -1 Then xret& = SetPixel(pic2hdc%, c2x + p2y, c2y - p2x, c2&)
                    If c3& <> -1 Then xret& = SetPixel(pic2hdc%, c2x - p2y, c2y + p2x, c3&)
                  Next
                  t% = DoEvents()
                Next
               End Sub

