VERSION 2.00
Begin Form frmMain 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "MVaders -- Can You Save The World?  A Game by Mark Meany."
   ClientHeight    =   5460
   ClientLeft      =   990
   ClientTop       =   1770
   ClientWidth     =   7395
   Height          =   6150
   Icon            =   FRMMAIN.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7395
   Top             =   1140
   Width           =   7515
   Begin PictureBox picLoader 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1155
      Left            =   480
      ScaleHeight     =   1125
      ScaleWidth      =   1545
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin PictureBox picStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   60
      ScaleHeight     =   345
      ScaleWidth      =   7245
      TabIndex        =   1
      Top             =   60
      Width           =   7275
      Begin Label lblDebug 
         BackStyle       =   0  'Transparent
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "Courier New"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   420
         TabIndex        =   2
         Top             =   60
         Width           =   6555
      End
   End
   Begin PictureBox picGame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4995
      Left            =   60
      Picture         =   FRMMAIN.FRX:0302
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   483
      TabIndex        =   0
      Top             =   420
      Width           =   7275
   End
   Begin Timer tmrGameLoop 
      Interval        =   50
      Left            =   3120
      Top             =   2520
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFileAbout 
         Caption         =   "&About"
      End
      Begin Menu mnuFileSpacer 
         Caption         =   "-"
      End
      Begin Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin Menu mnuGame 
      Caption         =   "&Game"
      Begin Menu mnuGameNew 
         Caption         =   "&New"
      End
      Begin Menu mnuGamePause 
         Caption         =   "&Pause"
      End
      Begin Menu mnuGameAbort 
         Caption         =   "&Abort"
      End
      Begin Menu mnuGameSpacer 
         Caption         =   "-"
      End
      Begin Menu mnuGameOptions 
         Caption         =   "&Options"
      End
   End
End
Option Explicit

Dim miBoss As Integer

Sub Form_Activate ()

'Give initial instructions
ShowGameOver
picGame.Refresh

End Sub

Sub Form_KeyDown (KeyCode As Integer, Shift As Integer)

Debug.Print KeyCode

'Act on keys we need to monitor
Select Case KeyCode

Case KEY_CUR_LEFT   'Moving left
    giKeyStatus = giKeyStatus Or KEY_CUR_LEFT_FLAG

Case KEY_CUR_RIGHT  'Moving right
    giKeyStatus = giKeyStatus Or KEY_CUR_RIGHT_FLAG

Case KEY_FIRE       'Firing
    'Fire button is disabled
    If giFireLock Then Exit Sub
    
    'If game is not in progress, start it
    If giGameStatus Then
        If giGameStatus = GAME_STOPPED Then
            picGame.Picture = LoadPicture("")
            giLevel = 1
            giScore = 0
            giLives = 3
            InitL1 0
            tmrGameLoop.Interval = GamePrefs.iTimer
        End If
        giGameStatus = GAME_PLAYING
    'Otherwise flag player is firing
    Else
        giKeyStatus = giKeyStatus Or KEY_FIRE_FLAG
    End If

Case KEY_PAUSE      'Pausing the game
    If Not giGameStatus Then giGameStatus = GAME_PAUSED

Case KEY_ABORT
    giGameStatus = GAME_STOPPED

Case KEY_QUIT       'Quitting the game
    tmrGameLoop.Enabled = False
    frmMain.Hide

End Select

End Sub

Sub Form_KeyUp (KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case KEY_CUR_LEFT   'Moving left
    giKeyStatus = giKeyStatus And (Not KEY_CUR_LEFT_FLAG)

Case KEY_CUR_RIGHT  'Moving right
    giKeyStatus = giKeyStatus And (Not KEY_CUR_RIGHT_FLAG)

Case KEY_FIRE       'Firing
    giKeyStatus = giKeyStatus And (Not KEY_FIRE_FLAG)
    giFireLock = False

End Select

End Sub

Sub Form_Load ()

'Center form on the screen
CenterForm Me

End Sub

Sub mnuFileAbout_Click ()

'If game is in progress then pause it
If giGameStatus = GAME_PLAYING Then giGameStatus = GAME_PAUSED

'Show the about window
frmAbout.Show VBModal

End Sub

Sub mnuFileExit_Click ()

'Just hide the form to quit
tmrGameLoop.Enabled = False
frmMain.Hide

End Sub

Sub mnuGameAbort_Click ()

'Abort no matter what status we are in!
giGameStatus = GAME_STOPPED

End Sub

Sub mnuGameNew_Click ()

'Abort current game if in progress!
giGameStatus = GAME_STOPPED
DoEvents

'And start a new one
picGame.Picture = LoadPicture("")
giLevel = 1
giScore = 0
giLives = 3
InitL1 0
tmrGameLoop.Interval = GamePrefs.iTimer

End Sub

Sub mnuGameOptions_Click ()

'Pause game if in progress
If giGameStatus = GAME_PLAYING Then
    giGameStatus = GAME_PAUSED
    giFireLock = True
End If

'Allow user to make changes
frmOptions.Show VBModal

End Sub

Sub mnuGamePause_Click ()

'Pause the game if in progress
If giGameStatus = GAME_PLAYING Then giGameStatus = GAME_PAUSED

End Sub

Sub tmrGameLoop_Timer ()

Dim i As Integer
Dim j As Integer
Dim iDC As Integer
Dim iX As Integer
Dim iY As Integer
Dim iXMin As Integer
Dim iXMax As Integer
Dim iYMax As Integer
Static iDy As Integer
Static iDx As Integer
Static iToggle As Integer
Static iDown As Integer
Static iBonus As Integer
Static iMod As Integer

Dim sDebug As String    'For debug only

'Initialise invaders speed
If iDx = 0 Then iDx = GamePrefs.iISpeed
If iMod = 0 Then iMod = 10

'Toggle is used for animation
If iToggle Then iToggle = 0 Else iToggle = 1

'Build status display
lblDebug = "Lives: " & Format$(giLives, "") & "           HIGH: " & Format$(giHiScore, "00000") & "    Level: " & Format$(giLevel, "00") & "      Score: " & Format$(giScore, "00000")

'Only process if game running
If giGameStatus = GAME_PLAYING Then

    'Get working DC
    iDC = picGame.hDC

    'Remove sprites
    VBSprRestoreBgrnd iDC

    'Handle bonus ships
    If iBonus = 0 Then
        
        'Reset counter for next bonus ship
        iBonus = Int(Rnd * 40) + 20

        'Display bonus ship
        If gVBSpr(BONUS_SHIP_ID).iActive = False Then VBSprActivateSprite iDC, BONUS_SHIP_ID, 0, 1

    Else
        'Else just dec counter
        iBonus = iBonus - 1
    End If

    'If bonus ship is active, move it
    If gVBSpr(BONUS_SHIP_ID).iActive Then
        VBSprMoveSpriteRel BONUS_SHIP_ID, 8, 0, 18 + iToggle
        If gVBSpr(BONUS_SHIP_ID).iX > ((picGame.Width \ Screen.TwipsPerPixelX) - 24) Then VBSprDeactivateSprite BONUS_SHIP_ID
    End If

    'If explosion is active, animate it until all 3 frames shown
    If gVBSpr(EXPLOSION_ID).iActive Then
        gVBSpr(EXPLOSION_ID).iUser1 = gVBSpr(EXPLOSION_ID).iUser1 + 1
        If gVBSpr(EXPLOSION_ID).iUser1 = 3 Then
            gVBSpr(EXPLOSION_ID).iActive = False
        Else
            VBSprAnimateSprite EXPLOSION_ID, 14 + gVBSpr(EXPLOSION_ID).iUser1
        End If
    End If
    
    'Animate invaders
    For i = FIRST_INVADER_ID To LAST_INVADER_ID
        If gVBSpr(i).iW > 50 Then
            VBSprAnimateSprite i, 20 + iToggle
        Else
            VBSprAnimateSprite i, 2 * ((i - FIRST_INVADER_ID) \ 6) + iToggle
        End If
    Next i

    'Move the invaders
    If iDy Then
        'Move invaders down, they are at the edge of the screen!
        iYMax = -1
        For i = FIRST_INVADER_ID To LAST_INVADER_ID
            If gVBSpr(i).iActive Then
                gVBSpr(i).iY = gVBSpr(i).iY + iDy
                If (gVBSpr(i).iY + gVBSpr(i).iH) > iYMax Then iYMax = gVBSpr(i).iY + gVBSpr(i).iH
            End If
        Next i
        iDy = 0
    Else
        'Normal moving
        iXMin = 9999
        iXMax = -1
        For i = FIRST_INVADER_ID To LAST_INVADER_ID
            If gVBSpr(i).iActive Then
                gVBSpr(i).iX = gVBSpr(i).iX + iDx
                If (gVBSpr(i).iX + gVBSpr(i).iW) > iXMax Then iXMax = gVBSpr(i).iX + gVBSpr(i).iW
                If gVBSpr(i).iX < iXMin Then iXMin = gVBSpr(i).iX
            
                'Random invader shooting
                If Rnd > GamePrefs.fIBFreq Then
                    For j = FIRST_INVADER_BULLET_ID To LAST_INVADER_BULLET_ID
                        If gVBSpr(j).iActive = False Then
                            VBSprActivateSprite iDC, j, gVBSpr(i).iX, gVBSpr(i).iY
                            Exit For
                        End If
                    Next j
                End If

            End If
        Next i
        'Test if we need to move down on next frame, sets iDy if so
        If (iXMin <= 0) Or (iXMax >= (picGame.Width \ Screen.TwipsPerPixelX - 10)) Then
            iDx = iDx * -1
            iDy = GamePrefs.iIDrop
        End If
    End If

    'End game if vaders have landed
    If iYMax >= gVBSpr(PLAYER_ID).iY Then
        i = sndPlaySound(ByVal CStr(APP.Path & "\landed.wav"), SND_ASYNC)
        giGameStatus = GAME_STOPPED
        If giScore > giHiScore Then giHiScore = giScore
        iDx = GamePrefs.iISpeed
        iDy = 0
        iDown = 0
        iMod = 10
        giFireLock = True
    End If
    
    'Moves players ship
    If giKeyStatus And KEY_CUR_LEFT_FLAG Then iX = -1 * GamePrefs.iPSpeed
    If giKeyStatus And KEY_CUR_RIGHT_FLAG Then iX = GamePrefs.iPSpeed
    VBSprMoveSpriteRel PLAYER_ID, iX, 0, 10 + iToggle

    'If there is a bullet process it
    If giFiring Then

        'If bullet has reached top of display deactivate it
        If gVBSpr(BULLET_ID).iY <= 0 Then
            VBSprDeactivateSprite BULLET_ID
            giFiring = False
        End If

        'Move the bullet
        VBSprMoveSpriteRel BULLET_ID, 0, -1 * GamePrefs.iPBSpeed, 12 + iToggle
        
        'Check for bullet/invaders collision
        For i = FIRST_INVADER_ID To LAST_INVADER_ID
            If gVBSpr(i).iActive Then
                'Is there a hit
                If iVBSprCollision(BULLET_ID, i) Then
                    
                    'Only kill invader if hit enough times
                    gVBSpr(i).iUser1 = gVBSpr(i).iUser1 - 1
                    If gVBSpr(i).iUser1 = 0 Then
                        VBSprDeactivateSprite i
                        giScore = giScore + 5
                        giInvaders = giInvaders - 1
                        
                        'Start explosion
                        VBSprAnimateSprite EXPLOSION_ID, 14
                        VBSprActivateSprite iDC, EXPLOSION_ID, gVBSpr(i).iX + gVBSpr(i).iW \ 2, gVBSpr(i).iY + gVBSpr(i).iH \ 2
                        gVBSpr(EXPLOSION_ID).iUser1 = 0

                        'Speed up invaders as they get killed
                        If (giInvaders Mod iMod) = 0 Then iDx = iDx + iDx \ 2
                        If giInvaders = 1 Then iDx = iDx + iDx \ 2
                    End If

                    'Deactivate the bullet sprite
                    VBSprDeactivateSprite BULLET_ID
                    
                    giFiring = False
            
                    'Sound fx
                    PlayHitMe

                    'Prepare next level if all invaders killed
                    If giInvaders = 0 Then
                        If iDown < 100 Then iDown = iDown + 10
                        iDx = GamePrefs.iISpeed
                        giLevel = giLevel + 1
                        If giLevel Mod 5 = 0 Then
                            iMod = 3
                            InitL2 giLevel \ 5
                        Else
                            iMod = 10
                            InitL1 iDown
                        End If
                        giFireLock = True
                    End If
                    
                    'Dont need to check rest of invaders!
                    Exit For
                End If
            End If
        Next i
        
        'Check for bonus ship/bullet collision
        If gVBSpr(BONUS_SHIP_ID).iActive Then
            If iVBSprCollision(BULLET_ID, BONUS_SHIP_ID) Then
                'Deactivate the bullet sprite
                VBSprDeactivateSprite BULLET_ID
                VBSprDeactivateSprite BONUS_SHIP_ID
                giFiring = False
            
                'Start explosion
                VBSprAnimateSprite EXPLOSION_ID, 14
                VBSprActivateSprite iDC, EXPLOSION_ID, gVBSpr(BONUS_SHIP_ID).iX, gVBSpr(BONUS_SHIP_ID).iY
                gVBSpr(EXPLOSION_ID).iUser1 = 0

                'Sound Fx and scoring, note score increases with level
                PlayHitMe
                giScore = giScore + 10 * giLevel
            End If
        End If

    'See if we user wants to fire
    Else

        If giKeyStatus And KEY_FIRE_FLAG Then
        
            'Activate the players bullet
            VBSprActivateSprite iDC, BULLET_ID, gVBSpr(0).iX + 12, gVBSpr(0).iY - 18
            giFiring = True
        End If

    End If

    'Move the invaders bullets
    For j = FIRST_INVADER_BULLET_ID To LAST_INVADER_BULLET_ID
        If gVBSpr(j).iActive Then
            gVBSpr(j).iY = gVBSpr(j).iY + GamePrefs.iIBSpeed

            'Check for off bottom
            If gVBSpr(j).iY >= (picGame.Height \ Screen.TwipsPerPixelY) Then VBSprDeactivateSprite j

            'Check for hit player
            If iVBSprCollision(PLAYER_ID, j) Then
                i = sndPlaySound(ByVal CStr(APP.Path & "\hitship.wav"), SND_ASYNC)
                giLives = giLives - 1

                'All lives gone then game finishes
                If giLives = 0 Then
                    If giScore > giHiScore Then giHiScore = giScore
                    iDx = GamePrefs.iISpeed
                    iDy = 0
                    iDown = 0
                    giGameStatus = GAME_STOPPED
                    giFireLock = True
                Else
                    'Clear all alien bullet to give player a chance
                    VBSprDeactivateSprite FIRST_INVADER_BULLET_ID
                    VBSprDeactivateSprite FIRST_INVADER_BULLET_ID + 1
                    VBSprDeactivateSprite FIRST_INVADER_BULLET_ID + 2
                    giGameStatus = GAME_PAUSED
                    giFireLock = True
                End If
                Exit For
            End If
        End If

    Next j
    
    'Redraw sprites
    VBSprDrawSprites iDC

    'Update the display
    picGame.Refresh

ElseIf giGameStatus = GAME_STOPPED Then
    'Game is stopped so tell user how to start it!!!
    ShowGameOver
    picGame.Refresh
End If

End Sub

