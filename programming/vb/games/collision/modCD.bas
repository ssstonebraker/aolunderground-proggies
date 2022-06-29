Attribute VB_Name = "modCD"
Option Explicit
'============================================================
'== Author  : Richard Lowe
'== Date    : July 99
'== Contact : riklowe@hotmail.com
'============================================================
'== Desciption
'==
'== This module contains the CollisionDetect function
'==
'============================================================
'== Version History
'============================================================
'== 1.0  28-July-99  RL  Initial Release.
'============================================================

'------------------------------------------------------------
'Dimension variables
'------------------------------------------------------------
Dim r As Integer, c As Integer

Dim hNewBMP As Long
Dim hPrevBMP As Long
Dim tmpObj As Long

Dim hMemDC As Long
Dim nRet As Long

Dim blnCollision As Boolean
            
Dim iMaskWidth As Integer
Dim iMaskHeight As Integer

Dim iM1SrcX As Integer
Dim iM1SrcY As Integer

Dim iM2SrcX As Integer
Dim iM2SrcY As Integer

Dim iDestX As Integer
Dim iDestY As Integer

Dim iStartBlankWidth As Integer
Dim iStartBlankHeight As Integer

Function CollisionDetect(ByVal X1 As Integer, ByVal Y1 As Integer, picMask As PictureBox, ByVal X2 As Integer, ByVal Y2 As Integer, picMask1 As PictureBox, picBlank As PictureBox) As Boolean
'============================================================
'== Desciption
'==
'== Name    : CollisionDetect
'==
'== Inputs
'== X1       X position in pixels of the first sprite mask
'== Y1       Y position in pixels of the first sprite mask
'== picMask  Picturebox Object of the first sprite mask
'== X2       X position in pixels of the second sprite mask
'== Y2       Y position in pixels of the second sprite mask
'== picMask  Picturebox Object of the second sprite mask
'== picBlank Picturebox Object of a blank sprite
'==
'== Returns
'== TRUE     If The pixels of sprites intersect
'== FALSE    If The pixels of sprites do not intersect
'==
'== Notes
'== Remove or comment out the section of code marked *** in a real program
'== It is only included here to dislay the contents of the memory DC
'==
'============================================================
    
'------------------------------------------------------------
'This section of code calculates the overlapping mask section
'size, and defines the X and Y coordinates to be used to copy
'from each of the sprites into the memory DC.
'
'These calcs have to take into account the orientation of the
'two sprites
'------------------------------------------------------------
    
    If X1 <= X2 Then
        iMaskWidth = X1 + picMask.ScaleWidth - X2
        iM1SrcX = picMask.ScaleWidth - iMaskWidth
        iM2SrcX = 0
        iDestX = 0
        iStartBlankWidth = iMaskWidth
    Else
        iMaskWidth = X2 + picMask.ScaleWidth - X1
        iM1SrcX = 0
        iM2SrcX = picMask.ScaleWidth - iMaskWidth
        iDestX = 0
        iStartBlankWidth = iMaskWidth
    End If
    
    If Y1 <= Y2 Then
        iMaskHeight = Y1 + picMask.ScaleHeight - Y2
        iM1SrcY = picMask.ScaleHeight - iMaskHeight
        iM2SrcY = 0
        iDestX = 0
        iStartBlankHeight = iMaskHeight
    Else
        iMaskHeight = Y2 + picMask.ScaleHeight - Y1
        iM1SrcY = 0
        iM2SrcY = picMask.ScaleHeight - iMaskHeight
        iDestX = 0
        iStartBlankHeight = iMaskHeight
    End If
    
'------------------------------------------------------------
'Create a memory DC
'------------------------------------------------------------

    hMemDC = CreateCompatibleDC(Screen.ActiveForm.hdc)
    hNewBMP = CreateCompatibleBitmap(Screen.ActiveForm.hdc, picMask.ScaleWidth, picMask.ScaleHeight)
    hPrevBMP = SelectObject(hMemDC, hNewBMP)
    
'------------------------------------------------------------
'Blank the memory dc, and draw the two sprite sections into it.
'------------------------------------------------------------

    BitBlt hMemDC, 0, 0, picBlank.ScaleWidth, picBlank.ScaleHeight, picBlank.hdc, 0, 0, vbNotSrcCopy
    BitBlt hMemDC, iDestX, iDestY, iMaskWidth, iMaskHeight, picMask.hdc, iM1SrcX, iM1SrcY, vbSrcPaint
    BitBlt hMemDC, iDestX, iDestX, iMaskWidth, iMaskHeight, picMask1.hdc, iM2SrcX, iM2SrcY, vbSrcPaint
    
'------------------------------------------------------------
'Send it to a picture box, to make it visible
'This is only required in this demo
'***
    BitBlt frmMain.picCD.hdc, 0, 0, picBlank.ScaleWidth, picBlank.ScaleHeight, hMemDC, 0, 0, vbSrcCopy

'------------------------------------------------------------
'Examine the memory DC, and see if it contains any non white
'pixels. If so, set Collision = true and exit
'------------------------------------------------------------
    
    blnCollision = False
    For c = 0 To iMaskHeight - 1
        For r = 0 To iMaskWidth - 1
            If GetPixel(hMemDC, r, c) <> 16777215 Then
                blnCollision = True
                Exit For
            Else
            End If
        Next
        
        If blnCollision = True Then
            Exit For
        End If
        
    Next

'------------------------------------------------------------
'Clean up resources
    tmpObj = SelectObject(hMemDC, hPrevBMP)
    tmpObj = DeleteObject(hPrevBMP)
    tmpObj = DeleteDC(hMemDC)
    
'------------------------------------------------------------

    CollisionDetect = blnCollision
    
End Function


