Attribute VB_Name = "modMain"
Option Explicit
' ---------------------------------------------------------------------
' Global constants, variables, and others used within the game.
'
' *********************************************
' | @ Written by Pranay Uppuluri. @           |
' | @ Copyright (c) 1997-98 Pranay Uppuluri @ |
' *********************************************
'
' VB game example Break-Thru! by Mark Pruett ported
' to Visual Basic DirectX.
'
' Thanks for Patrice Scribe's DirectX.TLB for DirectX 3.0 or Higher,
' his dixuSprite Class, and his dixu module, this game looks
' to be easy to code.
'
' You can visit Patrice's home page at:
'
'           http://www.chez.com/scribe/  *OR*
'           http://ourworld.compuserve.com/homepages/pscribe/
'
' If it wasn't for his effort, I would have had to do a lot
' more coding than this!
' ---------------------------------------------------------------------

#If Win32 Then   ' Compile if the program is running in Windows 32 bit

' Game over's surface
Public ddsGameOver As DirectDrawSurface2

' Ball Information ---------------------------------------
Public bmpBall As dixuSprite

' Width and Height of the Ball...
Public Const ballW = 13
Public Const ballH = 13

' The current ball speed
Public XSpeed As Long
Public YSpeed As Long

' The slowest allowable ball speed & the fastest
Public MinXSpeed As Long
Public MinYSpeed As Long
Public MaxXSpeed As Long
Public MaxYSpeed As Long

' The units at which the ball speed can change
Public SpeedUnit As Long

' Either +1 or -1, determines the direction
' that the ball is moving.
Public Xdir As Long
Public Ydir As Long

' The starting position of the ball.
Public XStartBall As Long
Public YStartball As Long

' Number of balls left
Public NumBalls As Long  ' Number of balls left (Starts with MAX_BALLS)
Public Const MAX_BALLS = 4

' Paddle Information --------------------------------------
Public bmpPaddle As dixuSprite

' Width and Height of the Paddle...
Public Const paddleW = 30
Public Const paddleH = 12

' The starting position of the paddle
Public XStartPaddle As Long
Public YStartPaddle As Long

' The current amount of "english" that the paddle
' will apply to the ball.
Public PaddleEnglish As Long

' The amount that the paddle will move.
Public PaddleIncrement As Long

' Zrick Information ---------------------------------------
Public bmpZrick() As dixuSprite
Public zrickSurface As DirectDrawSurface2

Public Const zrickW = 31
Public Const zrickH = 17

Public Const BLOCKS_IN_ROW = 17
Public Const NUM_ROWS = 3
Public Const ZRICK_GAP = 4

' Bitmap files
Public Const tmpSplash = "tmpB0.tmp"
Public Const tmpBall = "tmpB1.tmp"
Public Const tmpPaddle = "tmpB2.tmp"
Public Const tmpZrick = "tmpB3.tmp"
Public Const tmpNum = "tmpB4.tmp"
Public Const tmpGameOver = "tmpB5.tmp"

' Surface to hold the Font.
Public ddsNum As DirectDrawSurface2

' The current score...
Public Score As String
Public lngScore As Long
Public xScore As Long ' x position that the score will be blitted to.
Public yScore As Long ' y position that the score will be blitted to.

' Non-digit score characters
Public Const ASC_0 = 48
Public Const SC_L = ASC_0 + 10
Public Const SC_E = ASC_0 + 11
Public Const SC_V = ASC_0 + 12
Public Const SC_SPEAKER = ASC_0 + 13
Public Const SC_SPACE = ASC_0 + 14

' Current Level
Public Level As Integer

' Boolean (True/False) value that indicates if the game
' has been paused.
Public Paused As Boolean
Public InGameOver As Boolean

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

' API Call that return any value other than 0 if the lpSrc1 & lpSrc2
' RECT's meet (for Collision detection purposes)
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Public Sub Zricks_Init()
' ----------------------------------------------------------
' Initializes DirectSound, DirectDraw, and loads resources.
' ----------------------------------------------------------
Dim tmpStr As String

    ChDir App.Path
    
    'ShowCrusor False
    
    ' Load resources...
    'frmMain.MousePointer = vbCustom
    'frmMain.MouseIcon = LoadResPicture(22, vbResCursor)
    
    ' Important!!! On Windows 32 bit platform, Binary
    ' RESOURCE DATA (i.e.. not any pics or icons or cursors or menus etc.)
    ' ARE TO BE CONVERTED TO UNICODE.
    tmpStr = StrConv(LoadResData(11, "BMP"), vbUnicode)
    
    Open tmpNum For Output As #1
        Print #1, tmpStr
    Close #1
    
    tmpStr = StrConv(LoadResData(12, "BMP"), vbUnicode)
    
    Open tmpPaddle For Output As #1
        Print #1, tmpStr
    Close #1
    
    tmpStr = StrConv(LoadResData(13, "BMP"), vbUnicode)
    
    Open tmpZrick For Output As #1
        Print #1, tmpStr
    Close #1
    
    tmpStr = StrConv(LoadResData(14, "BMP"), vbUnicode)
    
    Open tmpBall For Output As #1
        Print #1, tmpStr
    Close #1
    
    tmpStr = StrConv(LoadResData(15, "BMP"), vbUnicode)
    
    Open tmpGameOver For Output As #1
        Print #1, tmpStr
    Close #1
    
    ' Takes quite a bit of memory, so, set it to nothing (Null)
    tmpStr = vbNullString
    dixuInit 0, frmMain, 640, 480, 8
    
    Set ddsGameOver = dixuCreateSurface(210, 60, App.Path & "\" & tmpGameOver)
    
    ' Set the color key for Transperent Blts.
    Dim ddck As DDCOLORKEY
    ddck.dwColorSpaceLowValue = RGB(0, 0, 0)
    ddck.dwColorSpaceHighValue = ddck.dwColorSpaceLowValue
    ddsGameOver.SetColorKey DDCKEY_SRCBLT, ddck
    
    ' First Initialize sound, and then initialize DirectDraw Game stuff.
    InitSound
    InitGeneralGameData
End Sub

Public Sub Zricks_Done()
' -------------------------------------------------------
' Call this procedure when the Game is ending.
' -------------------------------------------------------
Dim i As Long

Set bmpPaddle = Nothing
Set bmpBall = Nothing
Set zrickSurface = Nothing
Set ddsNum = Nothing
Set ddsGameOver = Nothing
Set reg = Nothing

    For i = 1 To (NUM_ROWS * BLOCKS_IN_ROW)
        Set bmpZrick(i) = Nothing
    Next i
    
    On Error Resume Next
    
    For i = 0 To 5
        Kill App.Path & "\" & "tmpB" & i & ".tmp"
    Next i
    
    'ShowCursor True
    
    DSoundDone
    dixuDone
    End
End Sub

Public Sub InitGeneralGameData()
' --------------------------------------------------------
' Setup variables that don't change during game play.
' --------------------------------------------------------
Dim i As Long
Set bmpPaddle = New dixuSprite
Set bmpBall = New dixuSprite

    ReDim bmpZrick(1 To NUM_ROWS * BLOCKS_IN_ROW)
    
    Set zrickSurface = dixuCreateSurface(zrickW, zrickH, App.Path & "\" & tmpZrick)
    
    ' Initialize the Zricks font.
    Set ddsNum = dixuCreateSurface(0, 0, App.Path & "\" & tmpNum)
    Dim ddck As DDCOLORKEY
    ddck.dwColorSpaceLowValue = 0
    ddck.dwColorSpaceHighValue = 0
    ddsNum.SetColorKey DDCKEY_SRCBLT, ddck
   
    For i = 1 To (NUM_ROWS * BLOCKS_IN_ROW)
        Set bmpZrick(i) = New dixuSprite
        
        ' Initialize the Zricks class
        Set bmpZrick(i).Surface = zrickSurface
        bmpZrick(i).Key = "Zrick"
        
        bmpZrick(i).y = ZRICK_GAP + ZRICK_GAP - 1
        bmpZrick(i).X = ZRICK_GAP
        
        bmpZrick(i).Height = zrickH
        bmpZrick(i).Width = zrickW
        
        bmpZrick(i).UseColorKey = True
        bmpZrick(i).ColorKey = RGB(0, 0, 0)
        bmpZrick(i).Visible = False
        
        ' Give up some cycles
        DoEvents
    Next i

    NumBalls = MAX_BALLS
    Score = "0000"
    xScore = 15
    yScore = frmMain.ScaleHeight - 26
    
    Set bmpBall.Surface = dixuCreateSurface(ballW, ballH, App.Path & "\" & tmpBall)
    
    ' fill the tRect Struct
    bmpBall.Width = ballW
    bmpBall.Height = ballH
    
    bmpBall.UseColorKey = True
    bmpBall.ColorKey = RGB(0, 0, 0)
    bmpBall.Visible = False
    
    ' Determine the ball's start position based on the
    ' game's dimmensions...
    XStartBall = (frmMain.ScaleWidth - ballW) / 2
    YStartball = (frmMain.ScaleHeight) / 4
    
    ' Determine the paddle's start position
    XStartPaddle = (frmMain.ScaleWidth - paddleW) / 2
    YStartPaddle = frmMain.ScaleHeight - paddleH
    
    ' The slowest speed increment is one pixel.
    SpeedUnit = 1
    
    ' Set the minimum speed.
    MinXSpeed = SpeedUnit * 3
    MinYSpeed = MinXSpeed
    
    ' Set maximum speed
    MaxXSpeed = 5
    MaxYSpeed = MaxXSpeed
    
    ' Initial Speed is the slowest allowable.
    XSpeed = MinXSpeed
    YSpeed = MinYSpeed
    
    bmpBall.VelocityX = XSpeed
    bmpBall.VelocityY = YSpeed
    
    ' Initialize Paddle Class
    ' Setup the initial state of the paddle.
    PaddleEnglish = 0
    PaddleIncrement = 5
    
    Set bmpPaddle.Surface = dixuCreateSurface(paddleW, paddleH, App.Path & "\" & tmpPaddle)
    bmpPaddle.Key = "Paddle"
    
    bmpPaddle.Height = paddleH
    bmpPaddle.Width = paddleW
    
    bmpPaddle.UseColorKey = True
    bmpPaddle.ColorKey = RGB(0, 0, 0)
    bmpPaddle.VelocityX = MaxXSpeed - 1
    
    bmpPaddle.Visible = False
    
    ' Setup the CRegSetting Class
    Set reg = New CRegSettings
    
    ' Get the settings.
    GetScores
    
    ' Make sure the Primary Surface is clear.
    Dim fx As DDBLTFX
    
    With fx
        .dwSize = Len(fx)
        .dwFillColor = RGB(0, 0, 0)
    End With

    dixuPrimarySurface.Blt ByVal 0&, Nothing, ByVal 0&, DDBLT_COLORFILL, fx
    
    ' Setup a new level
    SetupNextLevel
    
    ' Set up the blocks...
    SetupBlocks
    
    ' Reset the ball and paddle
    ResetBall
    ResetPaddle
End Sub

Public Sub ResetBall()
' --------------------------------------------------------
' Move the ball back to its starting position.
' --------------------------------------------------------
    ' The ball always starts out going down and right.
    Xdir = 1
    Ydir = 1
    
    ' The ball starts with the minumum speed.
    bmpBall.VelocityX = MinXSpeed
    bmpBall.VelocityY = MinYSpeed
    
    ' move the ball to the starting position.
    bmpBall.X = XStartBall
    bmpBall.y = YStartball
    
    bmpBall.VelocityX = XSpeed
    bmpBall.VelocityY = YSpeed
    
    bmpBall.Visible = True
End Sub

Public Sub ResetPaddle()
' --------------------------------------------------------
' Move the paddle back to its starting position
' --------------------------------------------------------
    
    bmpPaddle.y = frmMain.ScaleHeight - paddleH - 10
    bmpPaddle.X = (frmMain.ScaleWidth - paddleW) / 2
    bmpPaddle.VelocityX = 0
    bmpPaddle.Visible = True
End Sub

Public Sub SetupNextLevel(Optional ByVal FreshStart As Boolean)
' --------------------------------------------------------
' Each time the user moves to a new level (after clearing
' all the blocks at the current level) the blocks must be
' replaced and the balls return
' --------------------------------------------------------
    
    ' Suspend the game play.
    If Paused = False Then Paused = True
    
    ' Clear the field
    Dim dixuClientRect As RECT
    
    Dim fx As DDBLTFX
    
    fx.dwSize = Len(fx)
    fx.dwRop = SRCCOPY
    GetClientRect Screen.ActiveForm.hwnd, dixuClientRect
    ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Left
    ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Right
    
    dixuBackBufferClear
    
    ' Just to make sure...
    dixuPrimarySurface.Blt dixuClientRect, dixuBackBuffer, ByVal 0&, DDBLT_ROP Or DDBLT_WAIT, fx
    dixuPrimarySurface.Blt dixuClientRect, dixuBackBuffer, ByVal 0&, DDBLT_ROP Or DDBLT_WAIT, fx

    If FreshStart = True Then
        Level = 1
        NumBalls = MAX_BALLS
        InGameOver = False
        Score = "0000"
        
        ' Now test if this is a High Score or not
        If IsAHiScore(lngScore) = True Then
            If lngScore <> 0 Then
                ' The score is a High Score. Ask the user to enter their
                ' name.
                frmMain.Visible = False
                frmNewScore.Show vbModal, frmMain
                frmMain.Visible = True
            End If
        End If
        
        lngScore = 0
    Else
        Level = Level + 1
    End If
    
    Dim buf As String
    
    buf = Chr$(SC_L) + Chr$(SC_E) + Chr$(SC_V) + Chr$(SC_E) + Chr$(SC_L)
    
    Call ScoreBlt(buf, frmMain.ScaleWidth / 2 - 64, frmMain.ScaleHeight / 2 - 8)
    Call ScoreBlt((Level), frmMain.ScaleWidth / 2 + 22, frmMain.ScaleHeight / 2 - 8)
    
    Dim f2x As DDBLTFX
    
    f2x.dwSize = Len(f2x)
    f2x.dwRop = SRCCOPY
    
    GetClientRect Screen.ActiveForm.hwnd, dixuClientRect
    ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Left
    ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Right
    dixuPrimarySurface.Blt dixuClientRect, dixuBackBuffer, ByVal 0&, DDBLT_ROP Or DDBLT_WAIT, f2x
    
    NoisePlay dsbNewLevel
    
    Sleep 2000
    
    Paused = False
    
    ' Setup the blocks
    SetupBlocks
    
    ' Reset the paddle and ball
    ResetPaddle
    ResetBall
    
    ' Call the Main Game sub
    Zricks_Game_Main
End Sub

Public Sub Zricks_Game_Main()
' ------------------------------------------------------------
' Game's Main Sub
' Here is where most of the game takes place.
' ------------------------------------------------------------
Dim rc1 As RECT  ' ball
Dim rc2(1 To NUM_ROWS * BLOCKS_IN_ROW) As RECT  ' zrick
Dim rc3 As RECT  ' paddle
Dim ln As Long
Dim ret As Long
Dim Xinc As Long
Dim Yinc As Long
Dim i As Long
Dim PaddleCollision As Integer

Static PrevPaddleCollision As Integer
Static MoreBlocks As Boolean

Do While Not dixuAppEnd And Paused = False
    DoEvents
    
    ' Draw the back buffer onto the primary surface
    dixuBackBufferDraw
    
    ' Determine how much, and in which direction, to move the ball.
    Xinc = Xdir * XSpeed
    Yinc = Ydir * YSpeed
    
    ' Ball will hit the right wall
    If (bmpBall.X + ballW + Xinc) >= frmMain.ScaleWidth Then
        ' Change the direction to the opposite
        Xdir = -Xdir
        Xinc = Xdir * XSpeed
        
        ' Play the wall hit sound
        NoisePlay dsbWallHit, GetPanValue(bmpBall.X, ballW)
    End If
    
    ' Ball will hit the left wall
    If (bmpBall.X + Xinc) <= 0 Then
        ' Change the direction to left
        Xdir = -Xdir
        Xinc = Xdir * XSpeed
        
        bmpBall.VelocityX = -(bmpBall.VelocityX)
        
        ' Play the wall hit sound
        NoisePlay dsbWallHit, GetPanValue(bmpBall.X, ballW)
    End If

    ' Ball will hit the top
    If (bmpBall.y + Yinc) <= 0 Then
        Ydir = -Ydir
        Yinc = Ydir * YSpeed
        
        bmpBall.VelocityY = -(bmpBall.VelocityY)
        
        ' Play the wall hit sound
        NoisePlay dsbWallHit, GetPanValue(bmpBall.X, ballW)
    End If
    
    ' Set the Rect structs using the current game pieces.
    rc1.Left = bmpBall.X
    rc1.Top = bmpBall.y
    rc1.Right = bmpBall.X + ballW - 1
    rc1.bottom = bmpBall.y + ballH - 1
    
    For i = 1 To (NUM_ROWS * BLOCKS_IN_ROW)
        rc2(i).Left = bmpZrick(i).X
        rc2(i).Top = bmpZrick(i).y
        rc2(i).Right = bmpZrick(i).X + bmpZrick(i).Width - 1
        rc2(i).bottom = bmpZrick(i).y + bmpZrick(i).Height - 1
    Next i
    rc3.Left = bmpPaddle.X
    rc3.Top = bmpPaddle.y
    rc3.Right = bmpPaddle.X + bmpPaddle.Width - 1
    rc3.bottom = bmpPaddle.y + bmpPaddle.Height - 1
    
    ' Check to see if the ball got past the padddle.
    If (bmpBall.y) >= rc3.Top Then
        MissedBall
    End If
    
    ' Check if the zrick and ball collided.
    PaddleCollision = Collided(rc1, rc3)
    
    ' Ball got past paddle (at the bottom of the field)
    If PaddleCollision Then
       Ydir = -Abs(Ydir)
       bmpBall.VelocityY = -Abs(bmpBall.VelocityY)
       
       ' Repaint the paddle
       bmpPaddle.Paint
       
       ' The x and y co-ordinates can't be negative
       bmpBall.X = Abs(bmpBall.X)
       bmpBall.y = Abs(bmpBall.y)
       bmpPaddle.X = Abs(bmpPaddle.X)
       bmpPaddle.y = Abs(bmpPaddle.y)
       
       ' Adjust ball dynamics for paddle english
       If Abs(PaddleEnglish) > 0 Then
        If PaddleEnglish > 0 Then
            If Xdir > 0 Then
                ' Speed it up
                XSpeed = XSpeed + SpeedUnit
                bmpBall.VelocityX = XSpeed
            Else
                ' Slow it down.
                XSpeed = XSpeed - SpeedUnit
                
                ' Reverse the ball's X direction
                Xdir = -Xdir
                bmpBall.VelocityX = -(bmpBall.VelocityX)
            End If
        ElseIf PaddleEnglish < 0 Then
            If Xdir < 0 Then
                ' Speed it up.
                XSpeed = -(XSpeed + SpeedUnit)
                bmpBall.VelocityX = XSpeed
            Else
                ' Slow it down.
                XSpeed = XSpeed - SpeedUnit
                
                ' Reverse the ball's X direction.
                Xdir = -Xdir
                bmpBall.VelocityX = -(bmpBall.VelocityX)
            End If
        End If
        
        ' Don't let the ball go too slow or too fast
        If XSpeed < MinXSpeed Then XSpeed = MinXSpeed
        If XSpeed > MaxXSpeed Then XSpeed = MaxXSpeed
       End If
      ' Play the paddle hit sound.
      NoisePlay dsbPaddleHit, GetPanValue(bmpBall.X, ballW)
      
      ' See if the ball collided with the blocks.
      ElseIf bmpBall.y < ((NUM_ROWS + 1) * bmpZrick(1).Height) Then
        MoreBlocks = False
        For i = 1 To (NUM_ROWS * BLOCKS_IN_ROW)
            If bmpZrick(i).Visible Then
                MoreBlocks = True
                If BlockCollided(rc1, rc2(i)) Then
                    ' Hide the block
                    bmpZrick(i).Visible = False
                    
                    ' If we hit a zrick, send the ball down
                    Ydir = Abs(Ydir)
                    bmpBall.VelocityY = Abs(bmpBall.VelocityY)
                    
                    ' Play the block hit sound
                    NoisePlay dsbZrickHit, GetPanValue(bmpZrick(i).X, bmpZrick(i).Width)
                    
                    ' The player gets a point for each block hit.
                    Score = Format(Val(Score) + 1, "0000")
                    lngScore = lngScore + 1
                    
                    ' BltFast the score onto the back buffer.
                    ScoreBlt Score, 15, frmMain.ScaleHeight - 22
                End If
            End If
        Next i
        
        ' Out of blocks and we still have balls left, so
        ' sack 'em up again!
        If (Not MoreBlocks) And (NumBalls < 0) Then
            SetupNextLevel
        End If
    End If
    
    ' This is used to avoid multiple collision detections
    ' for one single hit
    PrevPaddleCollision = PaddleCollision
    DoEvents
Loop
End Sub

Public Sub Zricks_Pause()
Paused = True
End Sub

Public Sub Zricks_Resume()
Paused = False
Zricks_Game_Main
End Sub

Public Sub SetupBlocks()
' -----------------------------------------------------------
' Setup the blocks between each round of game play.
' -----------------------------------------------------------
Dim i As Long
Dim XIncr As Long
Dim j As Long
Dim ArrPos As Long
    
    XIncr = zrickW + ZRICK_GAP
    bmpZrick(1).y = 6
    bmpZrick(1).X = 6
    
    For j = 1 To NUM_ROWS
        For i = 2 To BLOCKS_IN_ROW
            ArrPos = ((j - 1) * BLOCKS_IN_ROW) + i
            
            ' Place the block...
            bmpZrick(ArrPos).X = ZRICK_GAP + ((i - 1) * XIncr)
            bmpZrick(ArrPos).y = bmpZrick(1).y
            
            ' and paint it...
            bmpZrick(ArrPos).Visible = True
            bmpZrick(ArrPos).Paint
            
            ' Make the setup sound for each Zrick...
            NoisePlay dsbSetup
            
            ' Yield to OS
            DoEvents
        Next
        
        ' Calculate the new row position
        bmpZrick(1).y = bmpZrick(1).y + zrickH + ZRICK_GAP
        bmpZrick(1).Visible = False
    Next
End Sub

Public Function BlockCollided(A As RECT, B As RECT)
' --------------------------------------------------------
' Check to see if the RECT A and RECT B overlap
' each other.
' --------------------------------------------------------
Dim rc As RECT

    BlockCollided = IntersectRect(rc, A, B)
End Function

Public Function Collided(rc1 As RECT, rc2 As RECT) As Long
' -----------------------------------------------------------
' See if the two rectangles collide using the API function.
' -----------------------------------------------------------
Dim rc As RECT

    Collided = IntersectRect(rc, rc1, rc2)
End Function

Public Sub MissedBall()
' --------------------------------------------------------------
' Move the ball back to its starting position and create a new
' set of blocks and restart the level.
' --------------------------------------------------------------

    ' Play the missed ball sound
    NoisePlay dsbMissed, 0
    
    If NumBalls > 0 Then ' If the balls left over are greater than 0, then
        NumBalls = NumBalls - 1 ' reduce the balls left by 1
        ResetBall ' Reset the ball...
    Else
        GameOver
    End If
End Sub

Public Sub GameOver()
' ---------------------------------------------------------
' Call this procedure when there are no more balls left.
' ---------------------------------------------------------
Dim printX As Long, printY As Long
Dim rc As RECT
Dim i As Long
        
        ' Pause the game and set the GameOver flag
        Paused = True
        InGameOver = True
        
        rc.bottom = 60
        rc.Right = 210
        
        printX = (ScreenRect.Right - 210) / 2
        printY = (ScreenRect.bottom - 60) / 2

        dixuBackBufferClear
        dixuBackBuffer.BltFast printX, printY, ddsGameOver, rc, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        
        Dim fx As DDBLTFX
    
        fx.dwSize = Len(fx)
        fx.dwRop = SRCCOPY
        
        Dim dixuClientRect As RECT
        
        DoEvents
        
        GetClientRect Screen.ActiveForm.hwnd, dixuClientRect
        ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Left
        ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Right
            
        dixuPrimarySurface.Blt dixuClientRect, dixuBackBuffer, ByVal 0&, DDBLT_ROP Or DDBLT_WAIT, fx
        
        ' Sleep for a sec...
        Sleep 1000
End Sub


Public Sub ScoreBlt(Score As String, ByVal destX As Long, ByVal destY As Long)
' --------------------------------------------------------------
' Blts the score from the Score buffer to the back buffer
' --------------------------------------------------------------
Dim c As Integer
Dim i As Integer
Dim rc1 As RECT
Dim fx As DDBLTFX
Dim posX As Long, posY As Long

    posX = destX
    posY = destY
    
    For c = 1 To Len(Score)
        i = Asc(Mid$(Score, c)) - ASC_0
        rc1.Left = i * 16
        rc1.Top = 0
        rc1.Right = rc1.Left + 16
        rc1.bottom = rc1.Top + 16
        
        Call dixuBackBuffer.BltFast(posX, posY, ddsNum, rc1, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        posX = posX + 16
    Next c
End Sub

Public Function StrEncode(ByVal s As String, Key As Long, salt As Boolean) As String
' ------------------------------------------------------------------------------
' Encodes/Encrypts a string using key number.
' ------------------------------------------------------------------------------
Dim n As Long, i As Long, ss As String
Dim k1 As Long, k2 As Long, k3 As Long, k4 As Long, t As Long
Static saltvalue As String * 4
    
    If salt Then
        For i = 1 To 4
            t = 100 * (1 + Asc(Mid(saltvalue, i, 1))) * Rnd() * (Timer + 1)
            Mid(saltvalue, i, 1) = Chr(t Mod 256)
        Next
        s = Mid(saltvalue, 1, 2) & s & Mid(saltvalue, 3, 2)
    End If
    
    n = Len(s)
    ss = Space(n)
    ReDim sn(n) As Long
    
    k1 = 11 + (Key Mod 233): k2 = 7 + (Key Mod 239)
    k3 = 5 + (Key Mod 241): k4 = 3 + (Key Mod 251)
    
    For i = 1 To n: sn(i) = Asc(Mid(s, i, 1)): Next i
    
    For i = 2 To n: sn(i) = sn(i) Xor sn(i - 1) Xor ((k1 * sn(i - 1)) Mod 256): Next
    For i = n - 1 To 1 Step -1: sn(i) = sn(i) Xor sn(i + 1) Xor (k2 * sn(i + 1)) Mod 256: Next
    For i = 3 To n: sn(i) = sn(i) Xor sn(i - 2) Xor (k3 * sn(i - 1)) Mod 256: Next
    For i = n - 2 To 1 Step -1: sn(i) = sn(i) Xor sn(i + 2) Xor (k4 * sn(i + 1)) Mod 256: Next
    
    For i = 1 To n: Mid(ss, i, 1) = Chr(sn(i)): Next i
    
    StrEncode = ss
    saltvalue = Mid(ss, Len(ss) / 2, 4)

End Function

Public Function StrDecode(ByVal s As String, Key As Long, salt As Boolean) As String
' ---------------------------------------------------------------------------
' Decodes/Decrypts a string using the key number.
' ---------------------------------------------------------------------------
Dim n As Long, i As Long, ss As String
Dim k1 As Long, k2 As Long, k3 As Long, k4 As Long

    n = Len(s)
    ss = Space(n)
    ReDim sn(n) As Long
    
    k1 = 11 + (Key Mod 233): k2 = 7 + (Key Mod 239)
    k3 = 5 + (Key Mod 241): k4 = 3 + (Key Mod 251)
    
    For i = 1 To n: sn(i) = Asc(Mid(s, i, 1)): Next
    
    For i = 1 To n - 2: sn(i) = sn(i) Xor sn(i + 2) Xor (k4 * sn(i + 1)) Mod 256: Next
    For i = n To 3 Step -1: sn(i) = sn(i) Xor sn(i - 2) Xor (k3 * sn(i - 1)) Mod 256: Next
    For i = 1 To n - 1: sn(i) = sn(i) Xor sn(i + 1) Xor (k2 * sn(i + 1)) Mod 256: Next
    For i = n To 2 Step -1: sn(i) = sn(i) Xor sn(i - 1) Xor (k1 * sn(i - 1)) Mod 256: Next
    
    For i = 1 To n: Mid(ss, i, 1) = Chr(sn(i)): Next i
    
    If salt Then StrDecode = Mid(ss, 3, Len(ss) - 4) Else StrDecode = ss

End Function

#Else   ' Display a message box if this is not Windows 32 bit platform.
MsgBox "Zricks can only run under Windows 32 bit platform (i.e.. Windows '95 or Windows NT v3.5 or higher", vbInformation, "Wrong Platform"
#End If
