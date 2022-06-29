VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCoolbar"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' fCoolBar - written by Phil Ricketts - PRicketts@compuserve.com [100042,2041]
' ============================================================================
' This code has been released into the public domain by the author

' Immitation CoolBar Class
' ==================================================================================
' 0.00 - 12/3/97    Initial version written as a single procedure in VB3
' 0.01 - 13/3/97    Updated with flag variables to reduce unnecessary drawing and
'                   flickering
' 0.02 - 17/3/97    Added TransBlit() to convert hard coded colours in the
'                   bitmap to the equivalent Windows system colours
' 0.03 - 19/3/97    Rewritten as a class in VB4 16 bit
' 0.04 - 3/4/97     Added SetCapture() in order to remove the Timer
' 0.05 - 5/4/97     Added ToolTips - and the Timer again !
' 0.06 - 6/4/97     Added a help file
' 0.07 - 12/4/97    Added 32 bit support
' 0.08 - 13/4/97    Changed the ToolTip to use a separate form instead of BitBlting
'                   on to the Screen
' 0.09 - 17/4/97    Added code to manually calculate the minimum allowable width
'                   and height for the buttons based on the text to appear on each
'                   button, this is different if using large fonts
' 0.10 - 21/4/97    Added a property to switch on or off the conversion of the
'                   bitmap to system colours
'                   Added temporary Debug information to try to trace the colour
'                   conversion problem
' 0.11 - 22/4/97    Added Vertical separator bars which can group buttons together
'                   Removed the debug code
' 0.12 - 24/4/97    Added the Lefthand vertical bars 'handle'
'                   Small change to the TransBlit function to try to solve the
'                   colour conversion problem
' 0.13 - 10/5/97    Added another PictureBox to hold the original shades of grey
'                   copy of the Bitmap, this doesn't ever get changed, and if the
'                   system colours change while the program is running then a
'                   refresh will force the Bitmap to have its colours changed again
' 0.14 - 11/5/97    Class no longer assumes the ScaleMode for the form is PIXEL, and
'                   will cope with an MDI form which cannot have its ScaleMode
'                   changed.
'                   Stopped changing Black to ButtonText colour on the Bitmap
'                   Added a call to .Refresh from the Form_Paint event, this forces
'                   a complete redraw of the toolbar which in turn will check if the
'                   system colours have changed
' 0.15 - 13/5/97    Changed to initialising the PictureBoxes at runtime to get
'                   around the Windows95 TransBlit problem
'                   Added a property to switch on or off the HandleBar
' 0.16 - 14/5/97    Added an error handler to the Paint routine to help detect MDI
'                   Parent forms, an error is raised by VB5 but not by VB4
' 0.17 - 17/5/97    Added a Click Event handler so that the code that was previously
'                   in the MouseUp event can now go in the Click event
' ==================================================================================

Private Const csVersion = "0.17"

' Still to do
' ===========
' Add different button styles, Win3Style, Win95Style, Win97Style
' Remove additional PictureBoxes and use memory DCs for the working Bitmaps

' Methods
' =======
' .AddButton sCaption, sToolTipText - Adds a button to the array of buttons
' .AddBar iAfterWhichButton - Adds a vertical separator bar
' .Create form, picToolBar, picOriginalBitmap, picWorkingBitmap, picScratch, timer
'   - initialises the class and tells it which controls to use
' iActiveButton = .ClickEvent() - tells the class that we have clicked on the toolbar
' .EnableButton iButton, bEnabled - tells the class to enable or disable a button
' .KeyDownEvent - detects a key being pressed while the mouse is over the toolbar
' .KeyUpEvent - detects a key being pressed while the mouse is over the toolbar
' .LoweredOffset iLeft, iDown - tells the class to use a different number of pixels
'  as the offset from the normal position when drawing the button in the down position
' iActiveButton = .MouseDownEvent(iButton, iShift, rX, rY) - tells the class which
'  button it is over and draws that button in the down position and returns the index
'  of the button
' iActiveButton = .MouseMoveEvent(iButton, iShift, rX, rY) - tells the class which
'  button it is over and draws that button in the up position and returns the index
'  of the button
' iActiveButton = .MouseUpEvent(iButton, iShift, rX, rY) - tells the class which
'  button it is over and draws that button in the up position and returns the index
'  of the button
' .RaisedOffset iLeft, iDown - tells the class to use a different number of pixels
'  as the offset from the normal position when drawing the button in the up position
' .Refresh - tells the class to redraw itself
' .ResizeEvent - tells the class to redraw itself
' .SetBitmapSize iWidth, iHeight - tells the class what size of bitmap to expect to
'  find in the master bitmap, the bitmap stores a grid of bitmaps which must all be
'  the same size, in the example they are 20 by 20 and this is the default for the
'  class, but they can be any size, this is needed so that the class knows where to
'  find the bitmap for each button
' .SetBigButtonSize iWidth, iHeight - overrides the default calculated size for the
'  buttons when drawn with text
' .SetSmallButtonSize iWidth, iHeight - overrides the default calculated size fot the
'  buttons when drawn without text
' .TimerEvent - tells the class to show the tool tip text
'
' Properties
' ==========
' .BarStyle = iStyle - tells the class what style of 3D effects to use to use to draw
'  the toolbar
' .SeparatorStyle = iStyle - tells the class which style to use for the vertical
'  separator bar
' .ShowHandle = bTrueOrFalse - tells the class to display or hide the vertical handle
'  at the left side of the toolbar
' .ShowText = bTrueFalse - tells the class to displat or hide the text on the buttons
' .ShowToolTips = bTrueFalse - tells the class to show or hide the tooltip text when
'  the mouse pauses over a buttons
' .UseSystemColours = bTrueFalse - tells the class to draw the toolbar using the system
'  colours instead of the shades of grey, this property was added for debugging only,
'  and must be called before the call to the Create method (write only)
' .Version - read only property, returns the version of this class
'

#If Win16 Then
    
    Private Type POINTAPI
        X As Integer
        Y As Integer
    End Type
    Private Type BITMAP
        bmType As Integer
        bmWidth As Integer
        bmHeight As Integer
        bmWidthBytes As Integer
        bmPlanes As String * 1
        bmBitsPixel As String * 1
        bmBits As Long
    End Type
    
    Private Declare Function GetSysColor Lib "user" (ByVal nIndex As Integer) As Long
    Private Declare Function GetVersion Lib "kernel" () As Long
    Private Declare Function GetCursorPos Lib "user" (lpPoint As POINTAPI) As Integer
    
    Private Declare Function BitBlt Lib "gdi" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Long) As Integer
    Private Declare Function SetBkColor Lib "gdi" (ByVal hDC As Integer, ByVal cColor As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi" (ByVal hDC As Integer) As Integer
    Private Declare Function CreateDCz Lib "gdi" Alias "CreateDC" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Integer
    Private Declare Function DeleteDC Lib "gdi" (ByVal hDC As Integer) As Integer
    Private Declare Function CreateBitmap Lib "gdi" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal cbPlanes As Integer, ByVal cbBits As Integer, lpvBits As Any) As Integer
    Private Declare Function CreateCompatibleBitmap Lib "gdi" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
    Private Declare Function SelectObject Lib "gdi" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Private Declare Function DeleteObject Lib "gdi" (ByVal hObject As Integer) As Integer
    Private Declare Function GetObjectAPI Lib "gdi" Alias "GetObject" (ByVal hObject As Integer, ByVal nCount As Integer, bmp As Any) As Integer
    Private Declare Function SetCapture Lib "user" (ByVal hWnd As Integer) As Integer
    Private Declare Sub ReleaseCapture Lib "user" ()
#Else
    
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
    End Type
    
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function GetVersion Lib "kernel32" () As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    
    Private Declare Function BitBlt Lib "gdi32" (ByVal srchDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long, ByVal srcH As Long, ByVal desthDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal op As Long) As Long
    Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal cColor As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateDCz Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Integer
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cbPlanes As Long, ByVal cbBits As Long, lpvBits As Any) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObj As Any) As Long
    Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseCapture Lib "user32" () As Long
#End If

' BitBlt flag
Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086
Private Const NOTSRCCOPY = &H330008

' API constants - system colour constant flags
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNTEXT = 18

' Constants for the different Bar Styles available
Private Const BAR3DNONE = 0
Private Const BAR3DCUSTOM = 1
Private Const BAR3DRAISED = 2
Private Const BAR3DLOWERED = 3

' constants for the different Separator Bar styles
Private Const SEP_SHORT = 0
Private Const SEP_LONG = 1

Private Const PIXEL = 3

' the form and controls to use
Private cFrm As Form
Private cTimer As Timer
Private cPic As PictureBox
Private cPicBMP As PictureBox
Private cPicMasterBMP As PictureBox
Private cPicScratch As PictureBox

' a flag to check before trying to draw -
' this is set by calling the Create method
Private cbInitialised As Boolean

' the colours to use for drawing the buttons
Private clHilightColour As Long
Private clShadowColour As Long
Private clButtonColour As Long
Private clTextColour As Long
Private cbUseSystemColours As Boolean

' the number of buttons, and the text to draw on each - keep the text short, and make
' sure the buttons are wide enough to show the text
Private ciMaxButtons As Integer
Private cbShowText As Integer
Private csButtonText() As String
Private csButtonTip() As String
Private cbButtonEnabled() As Integer
Private ciButtonLeft() As Integer
Private ciMaxSeparators As Integer
Private ciSeparatorToLeftOfButtonNo() As Integer
Private ciSeparatorWidth As Integer
Private ciSeparatorBarStyle As Integer

' the offset from the top left corner of the PictureBox to the top left corner of the
' first button
Private ciButtonXOffset As Integer
Private ciButtonYOffset As Integer
Private cbShowHandle As Boolean
Private ciHandleWidth As Integer

' the size of the buttons - they are all the same size
Private ciBigButtonWidth As Integer
Private ciBigButtonHeight As Integer
Private ciSmallButtonWidth As Integer
Private ciSmallButtonHeight As Integer

Private ciBarHeight As Integer
Private ciBarStyle As Integer

Private ciRaisedXOffset As Integer
Private ciRaisedYOffset As Integer
Private ciLoweredXOffset As Integer
Private ciLoweredYOffset As Integer
Private ciImageWidth As Integer
Private ciImageHeight As Integer
Private ciImageVertOffset As Integer
Private ciTextVertOffset As Integer
    
Private cbShowToolTips As Boolean
Private ciToolTipShowAfter As Integer
Private csToolTipText As String
Private cbTipVisible As Boolean
Private ciTipMouseXOffset As Integer
Private ciTipMouseYOffset As Integer

' these store the last known active button and mouse state - used to cut
' down on unneccessary redrawing
Private ciLastSelectedButton As Integer
Private ciLastButtonState As Integer
Private crLastMouseMoveX As Single
Private crLastMouseMoveY As Single

Public Property Get Version() As String

    Version = csVersion
    
End Property

Public Property Let UseSystemColours(ByVal bUseSystemColours As Boolean)

    cbUseSystemColours = bUseSystemColours
    Call ChangeBitmapColours(cPicBMP, cPicScratch)
    Call Refresh
    
End Property

Public Property Let ShowHandle(ByVal bShowHandle As Boolean)

    cbShowHandle = bShowHandle
    Call CalculateButtonCoOrds
    Call Refresh
    
End Property

Public Property Get ShowHandle() As Boolean

    ShowHandle = cbShowHandle
    
End Property

Private Sub CalculateButtonCoOrds()

    ' works out the Left Co-Ords of all of the buttons
    ' must be called before the first Paint
    ' and must be called if the buttons get resized
    
    Dim iButtonCount As Integer
    Dim iSeparatorCount As Integer
    Dim iLeftOffset As Integer
    
    ' work out the left co-ord of all of the buttons (assuming no separators)
    For iButtonCount = 1 To ciMaxButtons
        If cbShowHandle = True Then
            iLeftOffset = ciHandleWidth + ciButtonXOffset
        Else
            iLeftOffset = ciButtonXOffset
        End If
        If cbShowText = True Then
            ciButtonLeft(iButtonCount) = iLeftOffset + (ciBigButtonWidth * (iButtonCount - 1))
        Else
            ciButtonLeft(iButtonCount) = iLeftOffset + (ciSmallButtonWidth * (iButtonCount - 1))
        End If
    Next
        
    ' push the buttons to the right to take account of the separators
    For iSeparatorCount = 1 To ciMaxSeparators
        For iButtonCount = ciSeparatorToLeftOfButtonNo(iSeparatorCount) To ciMaxButtons
            ciButtonLeft(iButtonCount) = ciButtonLeft(iButtonCount) + ciSeparatorWidth
        Next
    Next
    
End Sub

Private Function Paint(ByVal MouseButton As Integer, ByVal MouseX As Single, ByVal MouseY As Single) As Integer

    ' this is the main routine for drawing the CoolBar, this routine will work out
    ' from the co-ords passed in which button the mouse is over, and will redraw the
    ' bar as necessary in order to raise or lower the buttons, this should be called
    ' from the MouseMove, MouseDown and MouseUp events, it should also be called from
    ' the Resize event of the CoolBar PictureBox, since we will have to redraw the bar
    ' if the form gets resized
    
    ' working variables to hold the x and y co-ords of the graphic and the text when
    ' redrawing each button
    Dim iImageLeft As Integer
    Dim iImageTop As Integer
    Dim iTextLeft As Integer
    Dim iTextTop As Integer

    ' a working variable to hold the button the mouse is over - or 0 if it isn't over any
    ' this is also returned as the result of the function, and that value can be tested for
    ' in the mouse up event which is where the mouse click code should go or in the mouse
    ' over code if you want to update the status bar
    Dim iSelectedButton As Integer
    
    ' used as a counter
    Dim iButtonCount As Integer
    Dim iSeparatorCount As Integer
    
    ' working variables to hold the co-ords that each button will be drawn at
    Dim iLeft As Integer
    Dim iTop As Integer
    Dim iRight As Integer
    Dim iBottom As Integer
    
    Dim iButtonWidth As Integer
    Dim iButtonHeight As Integer
    
    Dim iStartFormScaleMode As Integer
    Dim rScaleHeightFactor As Single
    
    #If Win16 Then
        Dim iRet As Integer
    #Else
        Dim iRet As Long
    #End If
    
    If cbShowText = True Then
        iButtonWidth = ciBigButtonWidth
        iButtonHeight = ciBigButtonHeight
    Else
        iButtonWidth = ciSmallButtonWidth
        iButtonHeight = ciSmallButtonHeight
    End If
    
    If cbUseSystemColours = True Then
        If clHilightColour <> GetSysColor(COLOR_BTNHIGHLIGHT) Or _
          clShadowColour <> GetSysColor(COLOR_BTNSHADOW) Or _
          clButtonColour <> GetSysColor(COLOR_BTNFACE) Or _
          clTextColour <> GetSysColor(COLOR_BTNTEXT) Then
            ' the system colours have changed since we last changed the
            ' colours of the bitmap, so change them again
            cPicBMP.Picture = cPicMasterBMP.Picture
            Call ChangeBitmapColours(cPicBMP, cPicScratch)
            Debug.Print "Change in SystemColours detected"
            Call Refresh
        End If
    End If
    
    If ciMaxButtons > 0 Then
        
        ' make sure the Picture Box is big enough
        rScaleHeightFactor = 1
        
        On Error Resume Next
        
            ' this error handler is necessary because VB5 does not catch errors
            ' when the ScaleMode property is accessed on MDI Parent forms
            ' (it doesn't exist) - VB4 catches this and discards it without
            ' raising an error
        
            iStartFormScaleMode = cFrm.ScaleMode
            If cFrm.ScaleMode <> PIXEL Then
                cFrm.ScaleMode = PIXEL
                If Err <> 0 Or cFrm.ScaleMode <> PIXEL Then
                    ' must be an MDI form ... ScaleMode property wont change to PIXEL
                    rScaleHeightFactor = Screen.TwipsPerPixelY
                End If
            End If
            ciBarHeight = iButtonHeight + (ciButtonYOffset * 2)
            If cPic.Height <> (ciBarHeight + 1) * rScaleHeightFactor Then
                cPic.Height = (ciBarHeight + 1) * rScaleHeightFactor
            End If
            cFrm.ScaleMode = iStartFormScaleMode
            
        On Error GoTo 0
    
        ' determine if the mouse is over one of the buttons
        iSelectedButton = 0
        For iButtonCount = 1 To ciMaxButtons
            iLeft = ciButtonLeft(iButtonCount)
            iTop = ciButtonYOffset
            iRight = iLeft + iButtonWidth
            iBottom = iTop + iButtonHeight
            If MouseX >= iLeft And MouseY >= iTop And MouseX <= iRight And MouseY <= iBottom Then
                ' we are over one of the buttons
                iSelectedButton = iButtonCount
            End If
        Next
        If MouseX > cPic.ScaleWidth Then
            ' this detects if the form isn't wide enough to show all of the buttons
            iSelectedButton = 0
        End If
        
        ' decide which tooltip text to show - if any
        If iSelectedButton = 0 Then
            csToolTipText = ""
        Else
            csToolTipText = csButtonTip(iSelectedButton)
        End If
        
        ' start drawing - erase the control and redraw the 3D lines
        If MouseButton = 0 And MouseX = 0 And MouseY = 0 Then
            ' sentinel values - erase and redraw the control
            ' when calling this from the Resize event and the timer it should be called with
            ' these sentinel values to force a complete redraw
            
            cPic.BackColor = clButtonColour
            
            Select Case ciBarStyle
            Case BAR3DCUSTOM:
                cPic.Line (0, 0)-(cPic.ScaleWidth, cPic.ScaleHeight), cPic.BackColor, BF
                cPic.Line (0, 0)-(cPic.ScaleWidth, 0), clShadowColour, B
                cPic.Line (0, 1)-(cPic.ScaleWidth, 1), clHilightColour, B
                cPic.Line (0, ciBarHeight - 2)-(cPic.ScaleWidth, ciBarHeight - 2), clShadowColour, B
                cPic.Line (0, ciBarHeight - 1)-(cPic.ScaleWidth, ciBarHeight - 1), clHilightColour, B
            Case BAR3DRAISED:
                cPic.Line (0, 0)-(cPic.ScaleWidth, cPic.ScaleHeight), cPic.BackColor, BF
                cPic.Line (0, 0)-(cPic.ScaleWidth, 0), clHilightColour, B
                cPic.Line (0, ciBarHeight - 1)-(cPic.ScaleWidth, ciBarHeight - 1), clShadowColour, B
            Case BAR3DLOWERED:
                cPic.Line (0, 0)-(cPic.ScaleWidth, cPic.ScaleHeight), cPic.BackColor, BF
                cPic.Line (0, 0)-(cPic.ScaleWidth, 0), clShadowColour, B
                cPic.Line (0, ciBarHeight - 1)-(cPic.ScaleWidth, ciBarHeight - 1), clHilightColour, B
            Case Else:
                ' no 3D effects
                cPic.Line (0, 0)-(cPic.ScaleWidth, cPic.ScaleHeight), cPic.BackColor, BF
            End Select
            
            If cbShowHandle = True Then
                iLeft = 3
                iTop = ciButtonYOffset
                iRight = iLeft + 5
                iBottom = iTop + iButtonHeight
                cPic.Line (iLeft, iTop)-(iLeft, iBottom), clHilightColour, BF
                cPic.Line (iLeft + 3, iTop)-(iLeft + 3, iBottom), clHilightColour, BF
                cPic.Line (iLeft, iTop)-(iRight, iTop), clHilightColour, BF
                cPic.Line (iLeft + 2, iTop)-(iLeft + 2, iBottom), clShadowColour, BF
                cPic.Line (iRight, iTop)-(iRight, iBottom), clShadowColour, BF
                cPic.Line (iLeft, iBottom)-(iRight, iBottom), clShadowColour, BF

            End If
            
        Else
            ' NOTE: we dont want to even test for this if the sentinel values are passed in
            ' ALSO: there is no point is redrawing anything if nothing has changed since last time
            If iSelectedButton = ciLastSelectedButton And MouseButton = ciLastButtonState Then
                ' no change from last time
                Exit Function
            Else
                ciLastSelectedButton = iSelectedButton
                ciLastButtonState = MouseButton
            End If
        End If
    
        ' draw the Images and the text
        For iButtonCount = 1 To ciMaxButtons
                
            iLeft = ciButtonLeft(iButtonCount)
            iTop = ciButtonYOffset
            iRight = iLeft + iButtonWidth
            iBottom = iTop + iButtonHeight
            cPic.Line (iLeft, iTop)-(iRight, iBottom), cPic.BackColor, B
                
            iImageLeft = iLeft + ((iButtonWidth - ciImageWidth) / 2)
            iImageTop = iTop + ciImageVertOffset
                
            ' if the mouse is over one of the buttons then draw it in the up or down position
            ' depending on the MouseButton state otherwise draw it in the flat position
            If iButtonCount = iSelectedButton And cbButtonEnabled(iButtonCount) = True Then
                ' adjust the Image offset
                ' this button will be drawn in either the up or down position
                If MouseButton <> 1 Then
                    iImageLeft = iImageLeft + ciRaisedXOffset
                    iImageTop = iImageTop + ciRaisedYOffset
                Else
                    iImageLeft = iImageLeft + ciLoweredXOffset
                    iImageTop = iImageTop + ciLoweredYOffset
                End If
            End If
    
            ' erase the client area of the button then blit the new Image
            cPic.Line (iLeft + 1, iTop + 1)-(iRight - 1, iBottom - 1), cPic.BackColor, BF
            If cbButtonEnabled(iButtonCount) = True Then
                If iButtonCount <> iSelectedButton Then
                    ' blit the Image from the first row - enabled but not selected button
                    iRet = BitBlt(cPic.hDC, iImageLeft, iImageTop, ciImageWidth, ciImageHeight, cPicBMP.hDC, ciImageWidth * (iButtonCount - 1), 0, SRCCOPY)
                Else
                    ' blit the Image from the second row - selected button
                    iRet = BitBlt(cPic.hDC, iImageLeft, iImageTop, ciImageWidth, ciImageHeight, cPicBMP.hDC, ciImageWidth * (iButtonCount - 1), ciImageHeight, SRCCOPY)
                End If
            Else
                ' blit the Image from the third row - disabled button
                iRet = BitBlt(cPic.hDC, iImageLeft, iImageTop, ciImageWidth, ciImageHeight, cPicBMP.hDC, ciImageWidth * (iButtonCount - 1), ciImageHeight * 2, SRCCOPY)
            End If
    
            If cbShowText = True Then
                ' draw the text
                iTextLeft = iLeft + ((iButtonWidth - cPic.TextWidth(csButtonText(iButtonCount))) / 2)
                iTextTop = iImageTop + ciImageHeight + ciTextVertOffset
                If iButtonCount = iSelectedButton And cbButtonEnabled(iButtonCount) = True Then
                    ' adjust the TextOffset
                    ' this button will be drawn in either the up or down position
                    If MouseButton <> 1 Then
                        iTextLeft = iTextLeft + ciRaisedXOffset
                    Else
                        iTextLeft = iTextLeft + ciLoweredXOffset
                    End If
                End If
            
                If cbButtonEnabled(iButtonCount) = True Then
                    ' draw the text enabled
                    cPic.CurrentX = iTextLeft
                    cPic.CurrentY = iTextTop
                    cPic.ForeColor = clTextColour
                    cPic.Print csButtonText(iButtonCount)
                Else
                    ' draw the text disabled
                    cPic.CurrentX = iTextLeft
                    cPic.CurrentY = iTextTop
                    cPic.ForeColor = clHilightColour
                    cPic.Print csButtonText(iButtonCount)
        
                    cPic.CurrentX = iTextLeft - 1
                    cPic.CurrentY = iTextTop - 1
                    cPic.ForeColor = clShadowColour
                    cPic.Print csButtonText(iButtonCount)
                    
                    cPic.ForeColor = clTextColour
                End If
            End If
        Next
    
        ' draw the separator bars (if any)
        For iSeparatorCount = 1 To ciMaxSeparators
            Select Case ciSeparatorBarStyle
            Case SEP_LONG:
                ' this will draw a vertical separator the full height of the toolbar
                iLeft = ciButtonLeft(ciSeparatorToLeftOfButtonNo(iSeparatorCount))
                iLeft = iLeft - (ciSeparatorWidth / 2) - 1
                iRight = iLeft + 1
                Select Case ciBarStyle
                Case BAR3DCUSTOM, BAR3DRAISED, BAR3DLOWERED:
                    iTop = 1
                    iBottom = cPic.ScaleHeight - 2
                Case Else:
                    iTop = 0
                    iBottom = cPic.ScaleHeight - 1
                End Select
                cPic.Line (iLeft, iTop)-(iLeft, iBottom), clShadowColour
                cPic.Line (iRight, iTop)-(iRight, iBottom), clHilightColour
            Case SEP_SHORT:
                ' this will draw a separator bar just the height of the button
                iLeft = ciButtonLeft(ciSeparatorToLeftOfButtonNo(iSeparatorCount))
                iLeft = iLeft - (ciSeparatorWidth / 2) - 1
                iTop = ciButtonYOffset
                iRight = iLeft + 1
                iBottom = iTop + iButtonHeight
                cPic.Line (iLeft, iTop)-(iLeft, iBottom), clShadowColour, B
                cPic.Line (iRight, iTop)-(iRight, iBottom), clHilightColour, B
            End Select
        Next
        
        ' finally draw the selected button border - unless it is disabled
        If iSelectedButton > 0 And cbButtonEnabled(iSelectedButton) = True Then
            iLeft = ciButtonLeft(iSelectedButton)
            iTop = ciButtonYOffset
            iRight = iLeft + iButtonWidth
            iBottom = iTop + iButtonHeight
            If MouseButton <> 1 Then
                ' draw the button up
                cPic.Line (iLeft, iTop)-(iRight, iBottom), clHilightColour, B
                cPic.Line (iLeft, iBottom)-(iRight, iBottom), clShadowColour, B
                cPic.Line (iRight, iTop)-(iRight, iBottom), clShadowColour, B
            Else
                ' draw the button down
                cPic.Line (iLeft, iTop)-(iRight, iBottom), clShadowColour, B
                cPic.Line (iLeft, iBottom)-(iRight, iBottom), clHilightColour, B
                cPic.Line (iRight, iTop)-(iRight, iBottom), clHilightColour, B
            End If
        End If
        
        ' set the return value - this should be looked for in the MouseUp event to
        ' trigger code, or the MouseMove event if you want to update a Status Bar
        If iSelectedButton > 0 And cbButtonEnabled(iSelectedButton) = True Then
            Paint = iSelectedButton
        End If
        
    End If

End Function

Private Sub TransBlit(picDest As PictureBox, picSrc As PictureBox, ByVal destX As Integer, ByVal destY As Integer, ByVal TransColor As Long)

    ' this was cribbed from MSDN - it will allow you to draw transparent bitmaps
    ' I'm using it to change the standard 3D colours (shades of grey) to the actual
    ' system colours being used

    #If Win16 Then
    
        Dim SrcDC As Integer
        Dim saveDC As Integer
        Dim maskDC As Integer
        Dim invDC As Integer
        Dim resultDC As Integer
        Dim bmp As BITMAP
        Dim hResultBmp As Integer
        Dim hSaveBmp As Integer
        Dim hMaskBmp As Integer
        Dim hInvBmp As Integer
        Dim hPrevBmp As Integer
        Dim hSrcPrevBmp As Integer
        Dim hSavePrevBmp As Integer
        Dim hDestPrevBmp As Integer
        Dim hMaskPrevBmp As Integer
        Dim hInvPrevBmp As Integer
        Dim OrigColor As Long
        Dim iRet As Integer
        
    #Else
    
        Dim SrcDC As Long
        Dim saveDC As Long
        Dim maskDC As Long
        Dim invDC As Long
        Dim resultDC As Long
        Dim bmp As BITMAP
        Dim hResultBmp As Long
        Dim hSaveBmp As Long
        Dim hMaskBmp As Long
        Dim hInvBmp As Long
        Dim hPrevBmp As Long
        Dim hSrcPrevBmp As Long
        Dim hSavePrevBmp As Long
        Dim hDestPrevBmp As Long
        Dim hMaskPrevBmp As Long
        Dim hInvPrevBmp As Long
        Dim OrigColor As Long
        Dim iRet As Long
    
    #End If
    
    'Retrieve bitmap to get width (bmp.bmWidth) & height (bmp.bmHeight)
    iRet = GetObjectAPI(picSrc.Picture, Len(bmp), bmp)
    
    SrcDC = CreateCompatibleDC(picSrc.hDC)
    saveDC = CreateCompatibleDC(picSrc.hDC)
    maskDC = CreateCompatibleDC(picSrc.hDC)
    invDC = CreateCompatibleDC(picSrc.hDC)
    resultDC = CreateCompatibleDC(picSrc.hDC)

    'Create monochrome bitmaps for the mask-related bitmaps:
    
    hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    
    'Create color bitmaps for final result & stored copy of source
    hResultBmp = CreateCompatibleBitmap(picDest.hDC, bmp.bmWidth, bmp.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(picDest.hDC, bmp.bmWidth, bmp.bmHeight)
    hSrcPrevBmp = SelectObject(SrcDC, picSrc.Picture)
    hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
    hInvPrevBmp = SelectObject(invDC, hInvBmp)
    hDestPrevBmp = SelectObject(resultDC, hResultBmp)
    
    iRet = BitBlt(saveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCCOPY)
    
    'Create mask: set background color of source to transparent color.
    OrigColor = SetBkColor(SrcDC, TransColor)
    iRet = BitBlt(maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCCOPY)
    TransColor = SetBkColor(SrcDC, OrigColor)
        
    'Create inverse of mask to AND w/ source & combine w/ background.
    iRet = BitBlt(invDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, NOTSRCCOPY)
    'Copy background bitmap to result & create final transparent bitmap
    iRet = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, picDest.hDC, destX, destY, SRCCOPY)
    'AND mask bitmap w/ result DC to punch hole in the background by
    'painting black area for non-transparent portion of source bitmap.
    iRet = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, SRCAND)
    'AND inverse mask w/ source bitmap to turn off bits associated
    'with transparent area of source bitmap by making it black.
    iRet = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, SRCAND)
    'XOR result w/ source bitmap to make background show through.
    iRet = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, SRCPAINT)
    iRet = BitBlt(picDest.hDC, destX, destY, bmp.bmWidth, bmp.bmHeight, resultDC, 0, 0, SRCCOPY)
    iRet = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, saveDC, 0, 0, SRCCOPY)
    
    hPrevBmp = SelectObject(SrcDC, hSrcPrevBmp)
    hPrevBmp = SelectObject(saveDC, hSavePrevBmp)
    hPrevBmp = SelectObject(resultDC, hDestPrevBmp)
    hPrevBmp = SelectObject(maskDC, hMaskPrevBmp)
    hPrevBmp = SelectObject(invDC, hInvPrevBmp)
    
    iRet = DeleteObject(hSaveBmp)
    iRet = DeleteObject(hMaskBmp)
    iRet = DeleteObject(hInvBmp)
    iRet = DeleteObject(hResultBmp)
    iRet = DeleteDC(SrcDC)
    iRet = DeleteDC(saveDC)
    iRet = DeleteDC(invDC)
    iRet = DeleteDC(maskDC)
    iRet = DeleteDC(resultDC)
    
End Sub

Private Sub ChangeBitmapColours(picKeep As PictureBox, picScratch As PictureBox)

    ' using two Picture Boxes, copy the first picture (which was created using shades of grey)
    ' converting first one colour to the actual system 3D colours on this machine
    
    If cbUseSystemColours = True Then
        clHilightColour = GetSysColor(COLOR_BTNHIGHLIGHT)
        clShadowColour = GetSysColor(COLOR_BTNSHADOW)
        clButtonColour = GetSysColor(COLOR_BTNFACE)
        clTextColour = GetSysColor(COLOR_BTNTEXT)
    Else
        ' shades of grey
        clHilightColour = QBColor(15)
        clShadowColour = QBColor(8)
        clButtonColour = QBColor(7)
        clTextColour = QBColor(0)
        picKeep.Picture = cPicMasterBMP.Image
        Exit Sub
    End If

    picKeep.Picture = cPicMasterBMP.Image
    picScratch.Picture = picKeep.Image
    picScratch.Picture = LoadPicture("")
    
    ' change light grey to the button face colour
    picScratch.BackColor = clButtonColour
    Call TransBlit(picScratch, picKeep, 0, 0, QBColor(7))
    picScratch.Refresh
    picKeep.Picture = picScratch.Image
    
    ' change dark grey to the button shadow colour
    picScratch.BackColor = clShadowColour
    Call TransBlit(picScratch, picKeep, 0, 0, QBColor(8))
    picScratch.Refresh
    picKeep.Picture = picScratch.Image
    
    ' change bright white to the button highlight colour
    picScratch.BackColor = clHilightColour
    Call TransBlit(picScratch, picKeep, 0, 0, QBColor(15))
    picScratch.Refresh
    picKeep.Picture = picScratch.Image
    
End Sub

Public Sub TimerEvent()

    ' this event is used to determine if we should show or hide the tool tip info
    
    ' NOTE: the text that is to be displayed is set into a class level variable
    ' whenever the toolbar is redrawn
    If cbTipVisible = False Then
        Call ShowToolTip
        cTimer.Enabled = False
    End If
    
End Sub

Public Sub ResizeEvent()

    Call Refresh

End Sub

Public Function MouseUpEvent(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) As Integer
    
   Call MouseMoveEvent(0, 0, -1, -1)
   MouseUpEvent = Paint(0, X, Y)

End Function

Public Function ClickEvent() As Integer
   
    Dim pt As POINTAPI
    Dim iX As Integer
    Dim iY As Integer
    Dim iStartFormScaleMode As Integer
    Dim rScaleHeightFactor As Single
    Dim rScaleWidthFactor As Single
    Dim iFormBorderWidth As Integer
    Dim iTitleHeight As Integer
    
    #If Win16 Then
        Dim iRet As Integer
    #Else
        Dim iRet As Long
    #End If

    #If Win16 Then
        Call GetCursorPos(pt)
    #Else
        iRet = GetCursorPos(pt)
    #End If

    On Error Resume Next
    
        ' this error handler is necessary because VB5 does not catch errors
        ' when the ScaleMode property is accessed on MDI Parent forms
        ' (it doesn't exist) - VB4 catches this and discards it without
        ' raising an error
    
        rScaleHeightFactor = 1
        rScaleWidthFactor = 1
        iStartFormScaleMode = cFrm.ScaleMode
        If cFrm.ScaleMode <> PIXEL Then
            cFrm.ScaleMode = PIXEL
            If Err <> 0 Or cFrm.ScaleMode <> PIXEL Then
                ' must be an MDI form ... ScaleMode property wont change to PIXEL
                rScaleHeightFactor = Screen.TwipsPerPixelY
                rScaleWidthFactor = Screen.TwipsPerPixelX
            End If
        End If
        cFrm.ScaleMode = iStartFormScaleMode
        
    On Error GoTo 0
    
    iFormBorderWidth = ((cFrm.Width / Screen.TwipsPerPixelX) - (cFrm.ScaleWidth / rScaleWidthFactor)) / 2
    iTitleHeight = ((cFrm.Height / Screen.TwipsPerPixelY) - (cFrm.ScaleHeight / rScaleHeightFactor)) - iFormBorderWidth
    iX = pt.X - (cFrm.Left / Screen.TwipsPerPixelX) - iFormBorderWidth - (cPic.Left / rScaleHeightFactor)
    iY = pt.Y - (cFrm.Top / Screen.TwipsPerPixelY) - iTitleHeight
    
    ClickEvent = Paint(1, iX, iY)

End Function

Public Sub KeyDownEvent()

    ' needed to allow Alt+ keyboard shortcuts to be passed through
    
    Call MouseMoveEvent(0, 0, -1, -1)
    
End Sub

Public Sub KeyUpEvent()

    ' needed to allow Alt+ keyboard shortcuts to be passed through

    Call MouseMoveEvent(0, 0, -1, -1)

End Sub

Public Function MouseMoveEvent(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) As Integer
    
    ' determine if there is any need to redraw the Cool Bar
    
    #If Win16 Then
        Dim iRet As Integer
    #Else
        Dim iRet As Long
    #End If
    
    If cbInitialised = True And X <> crLastMouseMoveX Or Y <> crLastMouseMoveY Then
        
        crLastMouseMoveX = X
        crLastMouseMoveY = Y
        
        ' determine if the mouse is inside the picture box
        ' if it is then make sure that the next mouse movement event gets sent to the PictureBox
        ' - even if the mouse isn't over the control
        If X >= 0 And X <= cPic.ScaleWidth And Y >= 0 And Y <= cPic.ScaleHeight Then
            ' inside the rect
            iRet = SetCapture(cPic.hWnd)
            cTimer.Enabled = False
            If cbShowToolTips = True Then
                cTimer.Enabled = True
                cTimer.Interval = ciToolTipShowAfter * 1000
            End If
        Else
            ' outside
            Call ReleaseCapture
            cTimer.Enabled = False
        End If
    
        MouseMoveEvent = Paint(Button, X, Y)
    
        ' reset the flag that prevents the tip from being reshown
        If cbTipVisible = True Then
            Call HideToolTip
        End If
    End If
    
End Function

Public Function MouseDownEvent(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) As Integer

    Call MouseMoveEvent(0, 0, -1, -1)
    MouseDownEvent = Paint(Button, X, Y)

End Function

Public Sub Create(frm As Form, pic As PictureBox, picMaster As PictureBox, picBMP As PictureBox, picScratch As PictureBox, tmr As Timer)

    ' this should be called before anything gets shown - most people have their
    ' button colours set to shades of grey, but sometimes people set them to
    ' other colours, I have drawn the bitmaps using the standard 16 colours and
    ' I have used shades of grey for the background and 3D look, this routine
    ' will change the 4 colours (white, light grey, dark grey and black to the
    ' system colours for button highlight, button face, button shadow and
    ' button text)

    Set cFrm = frm
    Set cPic = pic
    Set cPicBMP = picBMP
    Set cPicMasterBMP = picMaster
    Set cPicScratch = picScratch
    Set cTimer = tmr
    
    ' the colours to use for drawing
    If cbUseSystemColours = True Then
        clHilightColour = GetSysColor(COLOR_BTNHIGHLIGHT)
        clShadowColour = GetSysColor(COLOR_BTNSHADOW)
        clButtonColour = GetSysColor(COLOR_BTNFACE)
        clTextColour = GetSysColor(COLOR_BTNTEXT)
    Else
        ' shades of grey
        clHilightColour = QBColor(15)
        clShadowColour = QBColor(8)
        clButtonColour = QBColor(7)
        clTextColour = QBColor(0)
    End If
    
    ' ensure that the controls are set up the way we want them to be
    cPic.BorderStyle = 0
    cPic.AutoRedraw = True
    cPic.FontBold = False
    cPic.ScaleMode = 3
    cPic.Parent.ScaleMode = 3
    cPic.ForeColor = clTextColour
    cPic.BackColor = clButtonColour
    
    cPicMasterBMP.BorderStyle = 0
    cPicMasterBMP.ScaleMode = 3
    cPicMasterBMP.AutoRedraw = False
    cPicMasterBMP.Visible = False
    
    cPicBMP.BorderStyle = 0
    cPicBMP.ScaleMode = 3
    cPicBMP.AutoRedraw = True
    cPicBMP.Visible = False
    cPicBMP.Width = cPicMasterBMP.Width
    cPicBMP.Height = cPicMasterBMP.Height
    
    cPicScratch.BorderStyle = 0
    cPicScratch.ScaleMode = 3
    cPicScratch.AutoRedraw = True
    cPicScratch.Visible = False
    cPicScratch.Width = cPicBMP.Width
    cPicScratch.Height = cPicBMP.Height
    
    If cbUseSystemColours = True Then
        Call ChangeBitmapColours(picBMP, picScratch)
        cPicBMP.Picture = cPicMasterBMP.Picture
        Call ChangeBitmapColours(picBMP, picScratch)
    End If
    
    cTimer.Enabled = False
    
    cbInitialised = True
    
    Call Refresh
    
End Sub

Private Sub Class_Initialize()

    cbInitialised = False
    ciBarStyle = BAR3DCUSTOM
    
    ' determine the minimum size of the buttons when shown with text
    ciBigButtonWidth = 60
    ciBigButtonHeight = 40

    ' the default is to show the buttons with the text
    cbShowText = True
    
    cbShowHandle = False
    ciHandleWidth = 6
    ciButtonXOffset = 6
    ciButtonYOffset = 4
    
    ' the dimensions of each bitmap in the block of bitmaps
    ciImageWidth = 20
    ciImageHeight = 20
    
    ' work out the size of the buttons when shown without text
    ciSmallButtonWidth = ciImageWidth + 8
    ciSmallButtonHeight = ciImageHeight + 8
    
    ' setup how many pixels to insert between buttons for a separator
    ciSeparatorWidth = 6
    ciSeparatorBarStyle = SEP_SHORT
    
    ' the offset from the normal position to draw the bitmap and text when
    ' the button is raised or lowered
    ciRaisedXOffset = 0
    ciRaisedYOffset = 0
    ciLoweredXOffset = 1
    ciLoweredYOffset = 1
    ' the offset from the top of the button to the top of the bitmap
    ciImageVertOffset = 4
    ' the offset from the bottom of the bitmap to the top of the text
    ciTextVertOffset = 2

    ' by default we should show the tooltips after 1 second and vertically
    ' offset it from the cursor by 20 pixels
    cbShowToolTips = True
    ciToolTipShowAfter = 1
    ciTipMouseXOffset = 0
    ciTipMouseYOffset = 20
    
    ' by default we should manipulate the bitmap so that it is drawn using
    ' the system colours instead of just using shades of grey
    cbUseSystemColours = True
    
End Sub

Public Sub Refresh()

    ' force a complete redraw of the toolbar
    
    Dim iRet As Integer
    
    iRet = Paint(0, 0, 0)
    
End Sub

Public Sub AddButton(ByVal sCaption As String, ByVal sToolTipText As String)

    ciMaxButtons = ciMaxButtons + 1
    
    ReDim Preserve csButtonText(ciMaxButtons)
    ReDim Preserve csButtonTip(ciMaxButtons)
    ReDim Preserve cbButtonEnabled(ciMaxButtons)
    ReDim Preserve ciButtonLeft(ciMaxButtons)
    
    csButtonText(ciMaxButtons) = sCaption
    csButtonTip(ciMaxButtons) = sToolTipText
    cbButtonEnabled(ciMaxButtons) = True
    ciButtonLeft(ciMaxButtons) = 0
    
    If ciBigButtonWidth < cPic.TextWidth(sCaption) + 4 Then
        ' 4 to allow 2 pixels to the left and right of the text
        ciBigButtonWidth = cPic.TextWidth(sCaption) + 4
    End If
    If ciBigButtonHeight < ciImageVertOffset + ciImageHeight + ciTextVertOffset + cPic.TextHeight(sCaption) + 2 Then
        ' 2 to allow for a bottom border
        ciBigButtonHeight = ciImageVertOffset + ciImageHeight + ciTextVertOffset + cPic.TextHeight(sCaption) + 2
    End If
    
    Call CalculateButtonCoOrds

End Sub

Public Sub AddBar(ByVal iButtonToPrecede As Integer)

    ' adds a vertical bar before a specified button
    
    ciMaxSeparators = ciMaxSeparators + 1
    ReDim Preserve ciSeparatorToLeftOfButtonNo(ciMaxSeparators)
    
    ciSeparatorToLeftOfButtonNo(ciMaxSeparators) = iButtonToPrecede
    
    Call CalculateButtonCoOrds
    
End Sub

Public Property Let ShowText(ByVal fShow As Boolean)

    cbShowText = fShow
    Call CalculateButtonCoOrds
    Call Refresh

End Property

Public Property Get ShowText() As Boolean

    ShowText = cbShowText
    
End Property

Public Property Let SeparatorStyle(ByVal iStyle As Integer)

    Select Case iStyle
    Case SEP_SHORT, SEP_LONG:
        ciSeparatorBarStyle = iStyle
    Case Else:
        ciSeparatorBarStyle = SEP_SHORT
    End Select
    Call Refresh
    
End Property

Public Property Get SeparatorStyle() As Integer

    SeparatorStyle = ciSeparatorBarStyle
    
End Property

Public Sub EnableButton(ByVal iButtonNo As Integer, ByVal bEnabled As Boolean)

    If iButtonNo >= 1 And iButtonNo <= ciMaxButtons Then
        cbButtonEnabled(iButtonNo) = bEnabled
    End If
    
    Call Refresh
    
End Sub

Public Property Let BarStyle(ByVal iStyle As Integer)

    Select Case iStyle
    Case BAR3DNONE, BAR3DCUSTOM, BAR3DRAISED, BAR3DLOWERED:
        ciBarStyle = iStyle
    Case Else:
        ciBarStyle = BAR3DCUSTOM
    End Select
    
    Call Refresh
    
End Property

Public Property Get BarStyle() As Integer

    BarStyle = ciBarStyle
    
End Property

Public Property Let ShowToolTips(ByVal fShowTips As Boolean)

    cbShowToolTips = fShowTips
    If fShowTips <> True Then
        Call HideToolTip
    End If
    
End Property

Public Property Get ShowToolTips() As Boolean

    ShowToolTips = cbShowToolTips
    
End Property

Public Sub RaisedOffset(ByVal iXOffset As Integer, ByVal iYOffset As Integer)

    ciRaisedXOffset = iXOffset
    ciRaisedYOffset = iYOffset
    
End Sub

Public Sub LoweredOffset(ByVal iXOffset As Integer, ByVal iYOffset As Integer)

    ciLoweredXOffset = iXOffset
    ciLoweredYOffset = iYOffset
    
End Sub

Public Sub SetBitmapSize(ByVal iWidth As Integer, ByVal iHeight As Integer)

    ' we are changing the bitmap size from the default
    ' adjust the maximum height to suit and if the maximum width is dependent on the
    ' size of the bitmap (rather than the text) then adjust that also
    ciBigButtonHeight = ciBigButtonHeight + (iHeight - ciImageHeight)
    If ciBigButtonWidth <= ciSmallButtonWidth Then
        ciBigButtonWidth = ciBigButtonWidth - (iWidth - ciBigButtonWidth)
    End If
    
    ' adjust the minimum size of the button to suit the new size of the bitmap
    ciSmallButtonHeight = ciSmallButtonHeight - (iHeight - ciImageHeight)
    ciSmallButtonWidth = ciSmallButtonWidth - (iWidth - ciSmallButtonWidth)
    
    ciImageWidth = iWidth
    ciImageHeight = iHeight
    
End Sub

Public Sub SetBigButtonSize(ByVal iWidth As Integer, ByVal iHeight As Integer)

    ' this method allows you to override the calculated defaults for the button size
    
    ciBigButtonWidth = iWidth
    ciBigButtonHeight = iHeight

End Sub

Public Sub SetSmallButtonSize(ByVal iWidth As Integer, ByVal iHeight As Integer)

    ' this method allows you to override the calculated defaults for the button size
    
    ciSmallButtonWidth = iWidth
    ciSmallButtonHeight = iHeight

End Sub

Private Sub ShowToolTip()

    ' shows the tooltip form
    
    Dim pt As POINTAPI
    
    #If Win16 Then
        Dim iRet As Integer
    #Else
        Dim iRet As Long
    #End If

    If Len(Trim$(csToolTipText)) > 0 Then
        
        csToolTipText = " " & Trim$(csToolTipText) & " "
            
        #If Win16 Then
            Call GetCursorPos(pt)
        #Else
            iRet = GetCursorPos(pt)
        #End If
    
        ' NOTE: the order of events should be ... position the ToolTip form
        ' then set the caption to be displayed, when the caption is changed
        ' this will trigger the code to show the tooltip form
        frmToolTip.Left = (pt.X + ciTipMouseXOffset) * Screen.TwipsPerPixelX
        frmToolTip.Top = (pt.Y + ciTipMouseYOffset) * Screen.TwipsPerPixelY
        frmToolTip.lblToolTip.Caption = csToolTipText

        cbTipVisible = True

    End If

End Sub

Private Sub HideToolTip()

    ' hides the ToolTip form if it is visible
    
    If cbTipVisible = True Then
        ' NOTE: setting the ToolTip Caption to the empty string will trigger
        ' code which will unload the ToolTip form
        frmToolTip.lblToolTip.Caption = ""
        cbTipVisible = False
    End If

End Sub

