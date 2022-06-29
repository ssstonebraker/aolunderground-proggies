Option Explicit

'Game Constants
'~~~~~~~~~~~~~~
'Flags for monitoring movement keys
Global Const KEY_CUR_LEFT_FLAG = 1
Global Const KEY_CUR_RIGHT_FLAG = 2
Global Const KEY_FIRE_FLAG = 4

'KeyCode values for important keys
Global Const KEY_CUR_LEFT = 37
Global Const KEY_CUR_RIGHT = 39
Global Const KEY_FIRE = 32
Global Const KEY_ABORT = 65
Global Const KEY_PAUSE = 80
Global Const KEY_QUIT = 81

'Game status
Global Const GAME_PLAYING = 0
Global Const GAME_STOPPED = 1
Global Const GAME_PAUSED = 2

'Sprite ID's

Global Const PLAYER_ID = 0
Global Const BULLET_ID = 1
Global Const FIRST_INVADER_ID = 2
Global Const LAST_INVADER_ID = 31
Global Const BONUS_SHIP_ID = 32
Global Const FIRST_INVADER_BULLET_ID = 33
Global Const LAST_INVADER_BULLET_ID = 35
Global Const EXPLOSION_ID = 36
Global Const BOSS_ID = 37

'How many points until an extra life is granted
Global Const EXTRA_LIFE = 500
Global Const START_LIVES = 3
Global Const MAX_LIVES = 5

'Odd VB costants
Global Const VBModal = 1        'To diaplay forms as Modal

'Windows API rectangle structure
Type RECT
    iLeft As Integer
    iTop As Integer
    iRight As Integer
    iBottom As Integer
End Type

'Windows API Point structure
Type POINTAPI
    iX As Integer
    iY As Integer
End Type

'Constants for BitBlt() copy modes
Global Const SRCCOPY = &HCC0020
Global Const SRCAND = &H8800C6
Global Const SRCPAINT = &HEE0086
Global Const NOTSRCCOPY = &H330008
Global Const SRCERASE = &H440328
Global Const SRCINVERT = &H660046

'Constants for objects Scale Mode
Global Const TWIPS = 1
Global Const PIXELS = 3
Global Const RES_INFO = 2
Global Const MINIMIZED = 1

Global Const SND_ASYNC = &H1

'My User Defined Types
'~~~~~~~~~~~~~~~~~~~~~

'Defines an image in loaded gfx bitmap
Type VBGfx
    iX As Integer           'TopLeft of this gfx
    iY As Integer           'TopRight of this gfx
    iW As Integer           'Width
    iH As Integer           'Height
End Type

'Defines a sprite on the screen
Type VBSprite
    iInUse As Integer       'Set if sprite is being used
    iActive As Integer      '0=Sprite is off, 1=Sprite is on
    iSaveOn As Integer      '0=No saves, 1=wipe as we go, 2=Bgrnd save
    iGfxX As Integer        'X position of sprite in Gfx bitmap (pixel coords)
    iGfxY As Integer        'Y position of sprite in Gfx bitmap (pixel coords)
    iTrans As Integer       'Set if doing transparent blits
    iW As Integer           'Width of the sprite
    iH As Integer           'Height of the sprite
    iX As Integer           'X position of sprite (pixel coords)
    iY As Integer           'Y position of sprite (pixel coords)
    lColour As Long         'Background colour if wiping & not restoring background
    iSaveDC As Integer      'DC for background save
    iSaveBmp As Integer     'BitMap for background save
    iSaveSav As Integer     'BitMap from DC
    iUser1 As Integer       'Varies according to sprite type
End Type

'Game preferences
Type prefs
    iTimer As Integer       'Timer value that controls game loop
    iIGap As Integer        'Invaders separation
    iISpeed As Integer      'Invaders initial speed
    iIBSpeed As Integer     'Invaders bullet speed
    fIBFreq As Single       'Invaders bullet frequency
    iIDrop As Integer       'Invaders drop rate
    iPSpeed As Integer      'Players speed
    iPBSpeed As Integer     'Players bullet speed
End Type

'Game Global Variables
'~~~~~~~~~~~~~~~~~~~~~

Global giKeyStatus As Integer   'Holds movement key flags
Global giGameStatus As Integer  'Playing, Paused or Stopped
Global giLevel As Integer       'What level player is on
Global giLives As Integer       'How many lives the player has left
Global giScore As Integer       'Players score
Global giHiScore As Integer     'Highest Score
Global gsHiName As String * 20  'Name of player with high score
Global giFiring As Integer      'Set when player has fired a bullet
Global giInvaders As Integer    'Number of invaders left to kill
Global giFireLock As Integer    'Set to disable fire button detection
Global GamePrefs As prefs       'Game preferences

'Invaders graphics
Dim miGfxDC As Integer
Dim miGfxBmp As Integer
Dim miGfxSav  As Integer
Dim miMaskDC As Integer
Dim miMaskBmp As Integer
Dim miMaskSav As Integer
Dim mVBGfx(22) As VBGfx         'Holds positions of all gfx images
Global gVBSpr(40) As VBSprite      'Holds all sprite details

'16 Bit API functions used by MVaders
Declare Function BitBlt% Lib "GDI" (ByVal hDestDC%, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal dwRop&)
Declare Function SetBkColor& Lib "GDI" (ByVal hDC%, ByVal crColor&)
Declare Function CreateCompatibleDC% Lib "GDI" (ByVal hDC%)
Declare Function DeleteDC% Lib "GDI" (ByVal hDC%)
Declare Function CreateBitmap% Lib "GDI" (ByVal nWidth%, ByVal nHeight%, ByVal nPlanes%, ByVal nBitCount%, ByVal lpBits As Any)
Declare Function CreateCompatibleBitmap% Lib "GDI" (ByVal hDC%, ByVal nWidth%, ByVal nHeight%)
Declare Function SelectObject% Lib "GDI" (ByVal hDC%, ByVal hObject%)
Declare Function DeleteObject% Lib "GDI" (ByVal hObject%)
Declare Function sndPlaySound Lib "MMSystem" (lpsound As Any, ByVal flag As Integer) As Integer
Declare Function PtInRect Lib "User" (lpRect As RECT, ptRect As Any) As Integer

'32 Bit API functions
'Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
'Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub CenterForm (frm As Form)

'Center the form on the screen
frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2

End Sub

Sub FreeAllSprites ()

Dim iMax As Integer
Dim i As Integer

iMax = UBound(gVBSpr) - 1

For i = 0 To iMax
    FreeSprite i
Next i

End Sub

Sub FreeGfx ()

'Purpose    To release all resources used to hold sprite graphics in memory.
'Entry      None, uses module level variables
'Exit       None, all resources released
'Notes      Clears all the following modal variables:
'           miGfxDC, miGfxBmp, miGfxSav, miMaskDC, miMaskBmp, miMaskSav

Dim i As Integer

'If there is a Gfx DC, free it
If miGfxDC Then

    'Swap BitMap back in
    i = SelectObject(miGfxDC, miGfxSav)

    'And free the DC
    i = DeleteDC(miGfxDC)
    miGfxDC = 0
        
End If

'If there is a GfxBmp, free it
If miGfxBmp Then
    i = DeleteObject(miGfxBmp)
    miGfxBmp = 0
End If

'Clear the swap pointer just to be complete
miGfxSav = 0

'If there is a Mask DC, free it
If miMaskDC Then

    'Swap BitMap back in
    i = SelectObject(miMaskDC, miMaskSav)

    'And free the DC
    i = DeleteDC(miMaskDC)
    miMaskDC = 0
        
End If

'If there is a MaskBmp, free it
If miMaskBmp Then
    i = DeleteObject(miMaskBmp)
    miMaskBmp = 0
End If

'Clear the swap pointer just to be complete
miMaskSav = 0

End Sub

Sub FreeSprite (riId As Integer)

Dim i As Integer
        
'Only proceed if sprite is used
If gVBSpr(riId).iInUse Then

    'If there is a background DC, free it
    If gVBSpr(riId).iSaveDC Then

        'Swap BitMap back in
        i = SelectObject(gVBSpr(riId).iSaveDC, gVBSpr(riId).iSaveSav)

        'And free the DC
        i = DeleteDC(gVBSpr(riId).iSaveDC)
        gVBSpr(riId).iSaveDC = 0
        
    End If

    'If there is a background Bmp, free it
    If gVBSpr(riId).iSaveBmp Then
        i = DeleteObject(gVBSpr(riId).iSaveBmp)
        gVBSpr(riId).iSaveBmp = 0
    End If

    'Clear the swap pointer just to be complete
    gVBSpr(riId).iSaveSav = 0

End If
    
'Mark as no longer active or in use
gVBSpr(riId).iActive = 0
gVBSpr(riId).iInUse = 0

End Sub

Sub GetHiScore (riVal As Integer, rsName As String)

Dim sFName As String
Dim iFNum As Integer

'Trap error if file not accessible
On Error GoTo GetHiScore_Err

'Name of ini file
sFName = App.Path & "\MVaders.dat"
iFNum = FreeFile

'Open the file
Open sFName For Input As #iFNum

'Read data from the file
Input #iFNum, riVal, rsName
Input #iFNum, GamePrefs.iTimer, GamePrefs.iIGap, GamePrefs.iISpeed, GamePrefs.iIBSpeed, GamePrefs.fIBFreq, GamePrefs.iIDrop, GamePrefs.iPSpeed, GamePrefs.iPBSpeed

'Close the file
Close #iFNum

Exit Sub
GetHiScore_Err:

'Default high score details
riVal = 200
rsName = "Mark Meany"

'Default game preferences
GamePrefs.iTimer = 50
GamePrefs.iIGap = 50
GamePrefs.iISpeed = 4
GamePrefs.iIBSpeed = 12
GamePrefs.fIBFreq = .9
GamePrefs.iIDrop = 20
GamePrefs.iPSpeed = 10
GamePrefs.iPBSpeed = 17

Exit Sub
End Sub

Function iCheckBullet (riBullet As Integer, riStart As Integer, riStop As Integer)

'This is a very basic collision check that looks to see
'if the center of a bullet sprite is contained in the
'bounding rectangle of a range of sprites

Dim i As Integer
Dim j As Integer
Dim iX As Integer
Dim iY As Integer
Dim tRect As RECT
Dim tPoint As POINTAPI
Dim lPoint As Long
Dim iRetVal As Integer

'Default to no collisions
iRetVal = -1

'Define the hot spot for bullet
tPoint.iX = gVBSpr(riBullet).iX + gVBSpr(riBullet).iW \ 2
tPoint.iY = gVBSpr(riBullet).iY + gVBSpr(riBullet).iH \ 2
lPoint = tPoint.iX + CLng(tPoint.iY) * &H10000

'Check all sprites for collision
For i = riStart To riStop
    If gVBSpr(i).iActive Then
        'Get bounding rectangle
        tRect.iLeft = gVBSpr(i).iX
        tRect.iTop = gVBSpr(i).iY
        tRect.iRight = gVBSpr(i).iX + gVBSpr(i).iW
        tRect.iBottom = gVBSpr(i).iY + gVBSpr(i).iH

        If PtInRect(tRect, lPoint) Then
            iRetVal = i
            Exit For
        End If
    End If
Next i

iCheckBullet = iRetVal

End Function

Function iGetSprite (riId As Integer, riGfx As Integer, riTrans As Integer) As Integer

'Purpose    To allocate a sprite from the sprite system.  Intialises resources required
'           for background saves.
'Entry      riId  -- Sprite identifier
'           riGfx -- Image to use for this sprite
'           riTrans -- Set True if transparent blitting is to be used
'Exit
'Notes

Dim iBmp As Integer
Dim iDC As Integer
Dim iRetVal As Integer

'Skip if sprite is in use
If gVBSpr(riId).iInUse = False Then

    'Allocate resources for background saves
    iDC = CreateCompatibleDC(miGfxDC)
    If iDC Then
        'Store DC
        gVBSpr(riId).iSaveDC = iDC

        'Create a BitMap for the background
        iBmp = CreateCompatibleBitmap(miGfxDC, mVBGfx(riGfx).iW, mVBGfx(riGfx).iH)

        'Only proceed if BitMap allocated
        If iBmp Then

            'Store the BitMap
            gVBSpr(riId).iSaveBmp = iBmp
                
            'Swap the BitMap into the DC
            gVBSpr(riId).iSaveSav = SelectObject(iDC, iBmp)

            'Copy details of initial gfx
            gVBSpr(riId).iInUse = True
            gVBSpr(riId).iGfxX = mVBGfx(riGfx).iX
            gVBSpr(riId).iGfxY = mVBGfx(riGfx).iY
            gVBSpr(riId).iW = mVBGfx(riGfx).iW
            gVBSpr(riId).iH = mVBGfx(riGfx).iH
            gVBSpr(riId).iTrans = riTrans

            'Indicate success
            iRetVal = True

        End If

    End If

    'Free resources if allocation failed
    If iRetVal = False Then FreeSprite riId

End If

'Return success of the operation
iGetSprite = iRetVal

End Function

Sub InitGame ()

'Init global vars for the game
giScore = 0
giLives = START_LIVES
giLevel = 1
GetHiScore giHiScore, gsHiName

'Load and initialise the graphics used in the game
LoadGfx frmMain.picLoader

End Sub

Sub InitL1 (riDown As Integer)

'Purpose    To set up sprites etc for level 1 of game

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim iX As Integer
Dim iY As Integer
Dim iDC As Integer

'Make sure game is stopped
'giGameStatus = GAME_STOPPED

'Free sprites in use
FreeAllSprites

'Get DC to work with
iDC = frmMain.picGame.hDC

'Build the players ship
i = iGetSprite(PLAYER_ID, 10, 0)
If i Then
    iX = (frmMain.picGame.Width \ Screen.TwipsPerPixelX) \ 2
    iY = frmMain.picGame.Height \ Screen.TwipsPerPixelY - 27
    VBSprActivateSprite iDC, 0, iX, iY
End If

'Build players bullet, leave inactive, this uses transparent blitting!
i = iGetSprite(BULLET_ID, 12, 1)

'Build the invader sprites
giInvaders = 0
For j = FIRST_INVADER_ID To LAST_INVADER_ID
    k = j - FIRST_INVADER_ID
    i = iGetSprite(j, 2 * (k \ 6), 0)
    If i Then
        iX = (k Mod 6) * GamePrefs.iIGap + 10
        iY = (k \ 6) * 30 + 16 + riDown
        gVBSpr(j).iUser1 = 1
        VBSprActivateSprite iDC, j, iX, iY
        giInvaders = giInvaders + 1
    End If
Next j

'Build the invaders bullet sprites
For j = FIRST_INVADER_BULLET_ID To LAST_INVADER_BULLET_ID
    i = iGetSprite(j, 17, 0)
Next j

'Build the explosion sprite
i = iGetSprite(EXPLOSION_ID, 14, 0)

'Build the bonus ships sprite
i = iGetSprite(BONUS_SHIP_ID, 18, 0)

'Build the Boss
i = iGetSprite(BOSS_ID, 20, 0)

'Configure game variables
giFiring = False

End Sub

Sub InitL2 (ByVal viInc As Integer)

'Purpose    To set up sprites etc for level 1 of game

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim iX As Integer
Dim iY As Integer
Dim iDC As Integer

'Make sure game is stopped
'giGameStatus = GAME_STOPPED

'Free sprites in use
FreeAllSprites

'Get DC to work with
iDC = frmMain.picGame.hDC

'Build the players ship
i = iGetSprite(PLAYER_ID, 10, 0)
If i Then
    iX = (frmMain.picGame.Width \ Screen.TwipsPerPixelX) \ 2
    iY = frmMain.picGame.Height \ Screen.TwipsPerPixelY - 27
    VBSprActivateSprite iDC, 0, iX, iY
End If

'Build players bullet, leave inactive, this uses transparent blitting!
i = iGetSprite(BULLET_ID, 12, 1)

giInvaders = 0
For i = 0 To 2
    For j = 0 To 2
        k = iGetSprite(FIRST_INVADER_ID + giInvaders, 20, 0)
        If k Then
            iX = j * gVBSpr(FIRST_INVADER_ID + giInvaders).iW + GamePrefs.iIGap + 10
            iY = i * gVBSpr(FIRST_INVADER_ID + giInvaders).iH + 5
            gVBSpr(FIRST_INVADER_ID + giInvaders).iUser1 = 4 + viInc
            VBSprActivateSprite iDC, FIRST_INVADER_ID + giInvaders, iX, iY
            giInvaders = giInvaders + 1
        End If
    Next j
Next i
            
'Build the invaders bullet sprites
For j = FIRST_INVADER_BULLET_ID To LAST_INVADER_BULLET_ID
    i = iGetSprite(j, 17, 0)
Next j

'Build the explosion sprite
i = iGetSprite(EXPLOSION_ID, 14, 0)

'Build the bonus ships sprite
i = iGetSprite(BONUS_SHIP_ID, 18, 0)

'Build the Boss
i = iGetSprite(BOSS_ID, 20, 0)

'Configure game variables
giFiring = False

End Sub

Function iVBSprCollision (riId1 As Integer, riId2 As Integer) As Integer

Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim xx1 As Integer
Dim yy1 As Integer
Dim xx2 As Integer
Dim yy2 As Integer
Dim iRetVal As Integer

x1 = gVBSpr(riId1).iX
y1 = gVBSpr(riId1).iY
x2 = x1 + gVBSpr(riId1).iW - 1
y2 = y1 + gVBSpr(riId1).iH - 1

xx1 = gVBSpr(riId2).iX
yy1 = gVBSpr(riId2).iY
xx2 = xx1 + gVBSpr(riId2).iW - 1
yy2 = yy1 + gVBSpr(riId2).iH - 1

'Default to a collision
iRetVal = True

'Check if collision is impossible
If (xx2 < x1) Then iRetVal = False
If (yy2 < y1) Then iRetVal = False
If (xx1 > x2) Then iRetVal = False
If (yy1 > y2) Then iRetVal = False

iVBSprCollision = iRetVal

End Function

Sub LoadGfx (picGfx As PictureBox)

'Purpose    To load a bitmap that contains all sprite gfx for game into a DC.
'           Also creates mask for all sprites ready for transparent BitBlt().
'Entry      picGfx - A PictureBox to use LoadPicture with.  Note that this is
'           also used to obtain a compatible DC for loaded file.
'           Autosize=True, AutoRedraw=True,Visible=False,ScaleMode=PIXELS
'Notes      Uses module level variables to store resource pointers:
'           miGfxDC, miGfxBmp, miGfxSav, miMaskDC, miMaskBmp, miMaskSav

Dim iSuccess As Integer
Dim iDC As Integer
Dim iBmp As Integer
Dim iTBmp As Integer
Dim iW As Integer
Dim iH As Integer
Dim iScaleMode As Integer
Dim i As Integer
Dim lColour As Long
Dim sFName As String

'Fix name of data file
sFName = App.Path & "\mvaders.bmp"

'Default to failure
iSuccess = False

'Load the gfx file into the picturebox
picGfx.Picture = LoadPicture(sFName)

'Make scale mode for PictureBox pixels as required by GDI
iScaleMode = picGfx.ScaleMode
picGfx.ScaleMode = PIXELS

'Get dimensions of this graphic
iW = picGfx.ScaleWidth
iH = picGfx.ScaleHeight

'Create a DC for the sprite
iDC = CreateCompatibleDC(picGfx.hDC)

'Only proceed if DC allocated
If iDC Then
            
    'Store DC
    miGfxDC = iDC

    'Create a BitMap for this sprite
    iBmp = CreateCompatibleBitmap(picGfx.hDC, iW, iH)

    'Only proceed if BitMap allocated
    If iBmp Then

        'Store the BitMap
        miGfxBmp = iBmp
                
        'Swap the BitMap into the DC
        miGfxSav = SelectObject(iDC, iBmp)

        'Copy graphics into the DC
        i = BitBlt(iDC, 0, 0, iW, iH, picGfx.hDC, 0, 0, SRCCOPY)

        'Create a DC for the mask
        iDC = CreateCompatibleDC(picGfx.hDC)

        'Only proceed if DC allocated
        If iDC Then
            
            'Store DC
            miMaskDC = iDC

            'Create a BitMap for this mask
            iBmp = CreateBitmap(iW, iH, 1, 1, ByVal 0&)

            'Only proceed if BitMap allocated
            If iBmp Then

                'Store the BitMap
                miMaskBmp = iBmp
                
                'Swap the BitMap into the DC
                miMaskSav = SelectObject(iDC, iBmp)

                'Generate the mask, uses QBColor(0) as transparent
                lColour = SetBkColor(picGfx.hDC, QBColor(0))
                i = BitBlt(iDC, 0, 0, iW, iH, picGfx.hDC, 0, 0, SRCCOPY)
                lColour = SetBkColor(picGfx.hDC, lColour)

                'Flag success
                iSuccess = True

                'Define images in this bitmap
                
                mVBGfx(0).iX = 0        'Alien 1, Frame 1
                mVBGfx(0).iY = 0
                mVBGfx(0).iW = 32
                mVBGfx(0).iH = 20
                
                mVBGfx(1).iX = 34       'Alien 1, Frame 2
                mVBGfx(1).iY = 0
                mVBGfx(1).iW = 32
                mVBGfx(1).iH = 20

                mVBGfx(2).iX = 0        'Alien 2, Frame 1
                mVBGfx(2).iY = 23
                mVBGfx(2).iW = 32
                mVBGfx(2).iH = 20
                
                mVBGfx(3).iX = 34       'Alien 2, Frame 2
                mVBGfx(3).iY = 23
                mVBGfx(3).iW = 32
                mVBGfx(3).iH = 20
                
                mVBGfx(4).iX = 0        'Alien 3, Frame 1
                mVBGfx(4).iY = 45
                mVBGfx(4).iW = 32
                mVBGfx(4).iH = 20
                
                mVBGfx(5).iX = 34       'Alien 3, Frame 2
                mVBGfx(5).iY = 45
                mVBGfx(5).iW = 32
                mVBGfx(5).iH = 20
                
                mVBGfx(6).iX = 0        'Alien 4, Frame 1
                mVBGfx(6).iY = 69
                mVBGfx(6).iW = 32
                mVBGfx(6).iH = 18
                
                mVBGfx(7).iX = 34       'Alien 4, Frame 2
                mVBGfx(7).iY = 69
                mVBGfx(7).iW = 32
                mVBGfx(7).iH = 18
                
                mVBGfx(8).iX = 0        'Alien 5, Frame 1
                mVBGfx(8).iY = 89
                mVBGfx(8).iW = 32
                mVBGfx(8).iH = 20
                
                mVBGfx(9).iX = 34       'Alien 5, Frame 2
                mVBGfx(9).iY = 89
                mVBGfx(9).iW = 32
                mVBGfx(9).iH = 20

                mVBGfx(10).iX = 0       'Players ship Frame 1
                mVBGfx(10).iY = 113
                mVBGfx(10).iW = 32
                mVBGfx(10).iH = 17
                
                mVBGfx(11).iX = 34      'Players ship Frame 2
                mVBGfx(11).iY = 113
                mVBGfx(11).iW = 32
                mVBGfx(11).iH = 17
                
                mVBGfx(12).iX = 1       'Players bullet frame 1
                mVBGfx(12).iY = 135
                mVBGfx(12).iW = 10
                mVBGfx(12).iH = 18

                mVBGfx(13).iX = 14      'Players bullet frame 2
                mVBGfx(13).iY = 135
                mVBGfx(13).iW = 10
                mVBGfx(13).iH = 18

                mVBGfx(14).iX = 39      'Explosion frame 1
                mVBGfx(14).iY = 136
                mVBGfx(14).iW = 24
                mVBGfx(14).iH = 17

                mVBGfx(15).iX = 3       'Explosion frame 2
                mVBGfx(15).iY = 157
                mVBGfx(15).iW = 24
                mVBGfx(15).iH = 17

                mVBGfx(16).iX = 30      'Explosion frame 3
                mVBGfx(16).iY = 157
                mVBGfx(16).iW = 24
                mVBGfx(16).iH = 17

                mVBGfx(17).iX = 27      'Invaders bullet
                mVBGfx(17).iY = 136
                mVBGfx(17).iW = 10
                mVBGfx(17).iH = 16

                mVBGfx(18).iX = 2       'Bonus ship frame 1
                mVBGfx(18).iY = 177
                mVBGfx(18).iW = 24
                mVBGfx(18).iH = 14

                mVBGfx(19).iX = 38      'Bonus ship frame 2
                mVBGfx(19).iY = 177
                mVBGfx(19).iW = 24
                mVBGfx(19).iH = 14

                mVBGfx(20).iX = 76      'Boss frame 1
                mVBGfx(20).iY = 0
                mVBGfx(20).iW = 132
                mVBGfx(20).iH = 85

                mVBGfx(21).iX = 76      'Boss frame 2
                mVBGfx(21).iY = 90
                mVBGfx(21).iW = 132
                mVBGfx(21).iH = 85

            End If
        End If
    End If
End If
    
'If we failed to load the file, free resources
If iSuccess = False Then FreeGfx

'Reset Scale Mode of Picture Box
picGfx.ScaleMode = iScaleMode
picGfx.Picture = LoadPicture()

End Sub

Sub Main ()

'Initialise variables
giKeyStatus = 0
giGameStatus = GAME_STOPPED

'Load the game window
Load frmMain

'Initialise the game
InitGame

'Show form as modal
frmMain.Show VBModal

'Make sure form is unloaded
Unload frmMain

'Free gfx resources
FreeGfx

'Free sprite resources
FreeAllSprites

'Save high score info
SaveHiScore giHiScore, gsHiName

End Sub

Sub PlayHitMe ()

Dim i As Integer
Dim sFName As String

i = Int(Rnd * 10) + 1
sFName = App.Path & "\dead" & Format$(i, "") & ".wav"
i = sndPlaySound(ByVal CStr(sFName), SND_ASYNC)

End Sub

Sub SaveHiScore (riVal As Integer, rsName As String)
Dim sFName As String
Dim iFNum As Integer

'Trap error if file not accessible
On Error GoTo SaveHiScore_Err

'Name of ini file
sFName = App.Path & "\MVaders.dat"
iFNum = FreeFile

'Open the file
Open sFName For Output As #iFNum

'Read data from the file
Write #iFNum, riVal, rsName
Write #iFNum, GamePrefs.iTimer, GamePrefs.iIGap, GamePrefs.iISpeed, GamePrefs.iIBSpeed, GamePrefs.fIBFreq, GamePrefs.iIDrop, GamePrefs.iPSpeed, GamePrefs.iPBSpeed

'Close the file
Close #iFNum

Exit Sub
SaveHiScore_Err:

riVal = 1000
rsName = "Mark Meany"
Exit Sub
End Sub

Sub ShowGameOver ()

Dim i As Integer
Dim iX As Integer
Dim iY As Integer

iX = ((frmMain.picGame.Width \ Screen.TwipsPerPixelX) - 89) \ 2
iY = ((frmMain.picGame.Height \ Screen.TwipsPerPixelY) - 57) \ 2

i = BitBlt(frmMain.picGame.hDC, iX, iY, 89, 57, miGfxDC, 0, 205, SRCCOPY)

End Sub

Sub SplatGfx (iSX As Integer, iSY As Integer, iW As Integer, iH As Integer, picDst As PictureBox, iDX As Integer, iDY As Integer)

'Purpose    Copies gfx data to display for debug

Dim i As Integer

i = BitBlt(picDst.hDC, iSX, iSY, iW, iH, miGfxDC, iDX, iDY, SRCCOPY)

End Sub

Sub VBSprActivateSprite (hDC As Integer, riId As Integer, riX As Integer, riY As Integer)

'Purpose    Turn a sprite on so that it will be displayed
'Entry      riId - The sprite to activate
'           riX, riY - where to position the sprite
'Notes      Sprite must be in use and inactive

Dim i As Integer

If riId < UBound(gVBSpr) Then
    If gVBSpr(riId).iInUse Then
        If gVBSpr(riId).iActive = False Then
            gVBSpr(riId).iActive = True
            gVBSpr(riId).iX = riX
            gVBSpr(riId).iY = riY
        End If
    End If
End If

End Sub

Sub VBSprAnimateSprite (riId As Integer, riGfx As Integer)

'Purpose    Change anim frame of a sprite

'Notes      Only call if sprite has had background restored

'Do relative move
gVBSpr(riId).iGfxX = mVBGfx(riGfx).iX
gVBSpr(riId).iGfxY = mVBGfx(riGfx).iY

End Sub

Sub VBSprDeactivateSprite (riId As Integer)

If riId < UBound(gVBSpr) Then
    If gVBSpr(riId).iInUse Then gVBSpr(riId).iActive = False
End If

End Sub

Sub VBSprDrawSprites (hDC As Integer)

'Purpose    To save background & draw sprites with transparent bgrnd

Dim i As Integer
Dim iMax As Integer
Dim j As Integer

'Get number of sprites
iMax = UBound(gVBSpr) - 1

'Check each sprite
For i = 0 To iMax

    'Sprite must be active
    If gVBSpr(i).iActive Then

        'If transparent, do funky thing
        If gVBSpr(i).iTrans Then

            'Copy the mask to screen
            j = BitBlt(hDC, gVBSpr(i).iX, gVBSpr(i).iY, gVBSpr(i).iW, gVBSpr(i).iH, miMaskDC, gVBSpr(i).iGfxX, gVBSpr(i).iGfxY, SRCAND)

            'Copy the sprite to screen
            j = BitBlt(hDC, gVBSpr(i).iX, gVBSpr(i).iY, gVBSpr(i).iW, gVBSpr(i).iH, miGfxDC, gVBSpr(i).iGfxX, gVBSpr(i).iGfxY, SRCINVERT)

        'Otherwise do a straight forward copy
        Else

            j = BitBlt(hDC, gVBSpr(i).iX, gVBSpr(i).iY, gVBSpr(i).iW, gVBSpr(i).iH, miGfxDC, gVBSpr(i).iGfxX, gVBSpr(i).iGfxY, SRCCOPY)

        End If
    End If
Next i

End Sub

Sub VBSprExtent (riXMin As Integer, riXMax As Integer, riYMax As Integer, ByVal viStart As Integer, ByVal viStop As Integer)

'Purpose    To determine extremes of a sprite pack (invaders)
'Entry      riXMin -- container for smallest X value
'           riXMax -- container for largest X value
'           riStart -- Sprite Id to start search from
'           riStop -- Sprite Id to stop search at

Dim i As Integer
Dim x As Integer

riXMin = gVBSpr(viStart).iX
riXMax = riXMin
riYMax = -1

For i = viStart + 1 To viStop
    If gVBSpr(i).iActive Then
        x = gVBSpr(i).iX
        If x < riXMin Then riXMin = x
        If x > riXMax Then riXMax = x
        If gVBSpr(i).iY > riYMax Then riYMax = gVBSpr(i).iY
    End If
Next i

End Sub

Sub VBSprMoveSpriteRel (riId As Integer, riX As Integer, riY As Integer, riGfx As Integer)

'Purpose    Move a sprite relative to its current position

'Notes      Only call if sprite has had background restored

'Spriute must be active to move it
If gVBSpr(riId).iActive Then

    'Do relative move
    gVBSpr(riId).iX = gVBSpr(riId).iX + riX
    gVBSpr(riId).iY = gVBSpr(riId).iY + riY
    gVBSpr(riId).iGfxX = mVBGfx(riGfx).iX
    gVBSpr(riId).iGfxY = mVBGfx(riGfx).iY

End If

End Sub

Sub VBSprRestoreBgrnd (hDC As Integer)

'Purpose    Restores all saved backgrounds

Dim i As Integer
Dim iMax As Integer
Dim j As Integer

iMax = UBound(gVBSpr) - 1

'Go backwards through the array to restore in reverse order to save
For i = iMax To 0 Step -1

    'Sprite must be on
    If gVBSpr(i).iActive Then

        'Clear background
        j = BitBlt(hDC, gVBSpr(i).iX, gVBSpr(i).iY, gVBSpr(i).iW, gVBSpr(i).iH, hDC, gVBSpr(i).iX, gVBSpr(i).iY, SRCERASE)

    End If
Next i

End Sub

