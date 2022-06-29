VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   1050
   ClientWidth     =   10365
   ControlBox      =   0   'False
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDisplay.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   8595
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   23595
      Left            =   -22800
      ScaleHeight     =   23535
      ScaleWidth      =   27420
      TabIndex        =   0
      Top             =   -19320
      Visible         =   0   'False
      Width           =   27480
   End
   Begin VB.PictureBox picBuf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   -3840
      ScaleHeight     =   6735
      ScaleWidth      =   8850
      TabIndex        =   2
      Top             =   -2040
      Width           =   8850
   End
   Begin VB.PictureBox picSpr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6915
      Left            =   360
      Picture         =   "frmDisplay.frx":0A1C
      ScaleHeight     =   6855
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   -600
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variables must be declared before use
Option Explicit

'========================================================='
'========================================================='
'=============== RPG Game Version 0.0.4 =================='
'============== Written by Matthew Eagar ================='
'============ Compiled in Visual Basic 6.0 ==============='
'========================================================='
'========================================================='
'
'   This program is a work in progress.  As of yet it has no
'   actual plot, so it's not really a playable game.
'   I'm thinking towards making it almost like a MUD (you know
'   those text based RPG's?), except graphical.  That's something
'   fore the future though...
'
'   This isn't ment to be a full game, just a working engine.
'   there is no actual objective.  I havn't yet got doors
'   working, because that would require me to draw some more
'   textures for the insides of houses, which takes FOREVER!
'   Also, the textures could REALLY use some work,
'   as they were drawn in MS Paint.
'
'   This program may not run well on some computers.
'   The method used, bitblt, works well, but isn't designed for games.
'   It runs fine on a Pentium 233, but slow on a P75.  I havn't tested
'   it on anything in between those.
'
'   I'm still working on this, so look for me to post newer versions
'   of it.  It'll remain free, and it's really ment for educational purposes.
'
'   Check on www.planet-source-code.com for my latest entries.  I'll eventually
'   be opening my own web page, but for now I'm posting source code there.
'
'   Please contact me with ANY questions, comments, suggestions, or problems,
'   ANY input is welcome:
'
'   email:  meagar@home.com
'   ICQ:    45058462
'
'   Also, I havn't tested this on any computer running anything less then VB6.
'   I did run it in vb5, but it took some work.
'   You will need the VB6 runtime files the use this.
'
'   ====================================================================
'
'   Updates and Improvements over various versions:
'
'   Updates to Version 0.0.2:
'   =========================
'   -Added side scrolling and top scrolling
'   -Rechanged the map size from 13x11 to 30x30 tiles to accomidate side scrolling
'   -Added Bridge Tiles for bridge construction
'   -Added sound effects
'   -re-wrote most of movement code
'
'
'   Updates to Version 0.0.3:
'   =========================
'   -Texture tiles redrawn in greater detail
'   -Game speed increased by using While loop instead of timer
'   -Options menu added, for adjusting game speed, and walking speed for slower/
'       faster systems.
'   -Removed the usage of the Plus / Minus keys for game speed, because there
'       is now 2 kinds of speed adjustments, game and walking
'
'
'   Updates to Version 0.0.4:
'   =========================
'   -Added resolution code for changing the resolution and color depth
'   -Added neet startup menu
'   -Began implimenting character classes
'   -Added many new graphics for character classes
'   -Redrew character sprite in much greater detail.
'

Dim animX As Integer    'holds the current x location of the animation frame
Dim animY As Integer    'holds the current y location of the animation frame

Dim direction As Integer    'the direction the characters facing
Dim charX As Integer       'holds the character's x coords
Dim charY As Integer       'holds the character's y coords
Dim lastX As Integer    'holds the character's last y coords
Dim lastY As Integer    'holds the character's last x coords
Dim BackBuilt As Integer 'determines if the back ground needs to be built

'map variables
Dim mapX As Integer     'holds the current map x number
Dim mapY As Integer     'holds the current map y number
Dim mapArea As String      'the current map mapArea
Dim MapName As String   'holds the name of the map

'the location of the screen
Dim screenX As Integer  'holds the current location of the screen on the map
Dim screenY As Integer  'holds the current location of the screen on the map
Dim charPosX As Integer 'holds the coords to center the character on the screen
Dim charPosY As Integer 'holds the coords to center the character on the screen

Dim sound As Boolean     'holds whether to play sounds or not
Dim moving As Integer   'holds whether the character is moving or not
Dim changeFrame As Integer  'holds when the frame should be changed (animation is much to fast with out this)
Dim sndStep1 As String  'holds the path of the "step sound"

'symbolic constants
'directions
Const dLEFT As Integer = 1    'left direction
Const dUP As Integer = 2      'up direction
Const dRIGHT As Integer = 3   'right direction
Const dDOWN  As Integer = 4   'down direction

'animation frames
Const aLEFT As Integer = 133    'left animation
Const aUP As Integer = 67    'up animation
Const aRIGHT As Integer = 199 'right animation
Const aDOWN As Integer = 1  'down animation

'when the user presses a key
Private Sub picBuf_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim X As Integer 'counting variable

    'if movement, turn the mouse cursor into the invisible icon.
    'simply making a mouse cursor that was invisible is easier
    'then using API calls.
    frmDisplay.MouseIcon = frmTextures.picInvisible.Picture
    If moving <> 1 Then
        'determine how to act, based on which key the user presses.
        Select Case KeyCode
        Case Is = 37    'left arrow key
            animX = aLEFT   'set the animation frame to the proper direction
            direction = dLEFT 'set the direction
        Case Is = 38    'up arrow key
            animX = aUP 'set the animation frame to the proper direction
            direction = dUP
        Case Is = 39    'right arrow key
            animX = aRIGHT
            direction = dRIGHT
        Case Is = 40    'down arrow key
            animX = aDOWN
            direction = dDOWN
        Case Is = 27    'escape key
            'play the button sound
            Call sndPlaySound(sndButton, &H1)
            
            'show the startup window, and hide this one
            frmStartup.Show
            frmDisplay.Hide
            frmStartup.SetFocus
        Case Is = 83    'the S key
            'turn sound on or off
            If sound = True Then
                sound = False
            Else
                sound = True
            End If
        End Select
        redrawPic
        'moving <> 1 makes it so that the moveChar sub can't be called more then
        'once by holding the keydown
        If KeyCode >= 37 And KeyCode <= 40 And moving <> 1 Then  'if a direction key's been pressed
            'indicate that the character is moving
            moving = 1
        End If
    End If
End Sub
    
Private Sub picBuf_KeyUp(KeyCode As Integer, Shift As Integer)
    'indicate that the character isn't moving
    moving = 0
    'return the character to the intital frame, so they don't look like
    'they've taken half a step
    animY = 1
    Call redrawPic
End Sub

Private Sub Form_Load()

    'initialize the variables
    animX = 1
    animY = 1
    
    sound = True
    BackBuilt = False
    
    If Right(App.Path, 1) = "\" Then
        sndStep1 = App.Path & "1.wav"
    Else
        sndStep1 = App.Path & "\1.wav"
    End If
    
    'maps are loaded in the following way:
    'take the mapX, then add the letter 'a' then take the mapY, then add ".map"
    'so, the first map is called 0a0.map, the map beside it is called
    '1a0.map, and the map above the first is called 0a1.map
    'eventually the middle letter will stand for the mapArea, eg a = lev 1, b = lev 2
    
    mapX = 0    'the current map
    mapY = 0
    mapArea = "a"  'the initial mapArea
    Speed = 15  'set the initial walking speed
    
    picBuf.Height = 7100
    picBuf.Width = 9500
    
    picBuf.Left = 50 'center the main picture box
    picBuf.Top = 50
    
    'set the origional location of the character, which is held in charX and charY, but
    'calculated from screenX and screenY
    screenX = 0
    screenY = 0
    
    'charPosX is the distance of the character from the left side of the screen
    'charPosY is the distance from the top
    charPosX = picBuf.Width * 0.03  'center the character on the screen
    charPosY = picBuf.Height * 0.03
    
    'calculate the origional location of the character
    charX = screenX + charPosX
    charY = screenY + charPosY
    
    Call BuildBack  'build the back ground
    Call redrawPic  'load the pic into the main pic box
    
    While (1)   'main program loop
    
        If moving = 1 Then
            If touching() <> 1 Then
                'move the character in the proper direction
                If direction = dLEFT Then
                    screenX = screenX - Speed
                ElseIf direction = dUP Then
                    screenY = screenY - Speed
                ElseIf direction = dRIGHT Then
                    screenX = screenX + Speed
                ElseIf direction = dDOWN Then
                    screenY = screenY + Speed
                End If
                        
                charX = screenX + charPosX
                charY = screenY + charPosY
            
                If changeFrame = 1 Then 'this causes the frame to be updated once every 2 loops
                    animY = animY + 57    'advance the frame, each frame is 50 pixels wide, + a 1 pixel border
                    changeFrame = 0
                        
                    'there are 8 frames in the character's animation: this sees if the last frame has
                    'been shown. if it has, it resets it to the first.
                    If animY >= 408 Then
                        animY = 1  'goes to first frame
                        'play the foot step sound
                        If sound = True Then Call sndPlaySound(sndStep1, &H1)
                    ElseIf animY >= 204 And animY <= 255 Then
                        'play the foot step sound
                        If sound = True Then Call sndPlaySound(sndStep1, &H1)
                    End If
                Else
                    'if the frame isn't changed this time, change it next time.
                    changeFrame = 1
                End If
            
            Else
                'if the char is touching something in the current direction,
                'indicate the character isn't moving
                moving = 0
                'return the character to the first frame, so that they don't look
                'like they've taken half a step
                animY = 1
                Call redrawPic
                'stop this sub
            
            End If
        
            'see if the back ground has been built
            If BackBuilt = False Then
                'build the background
                Call BuildBack
                BackBuilt = True
            End If
            
            
            'see if the character has left the screen, by checking if the character's
            'x or y position is greater then the total amount of tiles
            If screenX + 25 >= 1200 - charPosX Then 'if the character has left the right side of the screen
                mapX = mapX + 1 'set the current map name to the next map name
                screenX = 10 - charPosX - 25 'set the character's position back to the left side of the screen
                Call BuildBack  'redraw the back ground
                Call redrawPic  'reload the screen
            ElseIf screenX + 25 <= 0 - charPosX Then 'see if the character has left the left side of the screen
                mapX = mapX - 1 'set the current map name to the next map name
                screenX = 1190 - charPosX - 25 'set the character position to the right side of the screen
                Call BuildBack  'redraw the back ground
                Call redrawPic  'reload the screen
            ElseIf screenY + 25 <= 0 - charPosY Then  'see if the character has left the top of the screen
                mapY = mapY + 1 'set the current map name to the next map name
                screenY = 1190 - charPosY - 25 'set the characters position to the bottom of the screen
                Call BuildBack  'redraw the back ground
                Call redrawPic  'reload the screen
            ElseIf screenY + 25 >= 1200 - charPosY Then 'see if the character has left the bottom of the screen
                mapY = mapY - 1 'set the current map name to the next map name
                screenY = 10 - charPosY - 25  'move the character to the top of the screen
                Call BuildBack  'redraw the back ground
                Call redrawPic  'reload the screen
            End If
        End If
        
        Call redrawPic 'redraws the form
        delay
        DoEvents    'allow for keyup
    Wend
    
End Sub
    
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, animX As Single, animY As Single)
    'turn the mouse icon into the visible icon
    frmDisplay.MouseIcon = frmTextures.picVisible.Picture
End Sub

'assembles the back ground
Sub BuildBack()

    'this sub builds the back ground.  It is called only once per map,
    'as the map is built in a hidden pic box, and kept untill the next map is needed.
    
    Dim g As Integer    'counting variable
    Dim X As Integer    'holds x coords of tile
    Dim Y As Integer    'holds y coords of tile
    On Error GoTo errHandler
    
    'set the name of the map
    If Right(App.Path, 1) = "\" Then
        MapName = App.Path & mapX & mapArea & mapY & ".map"
    Else
        MapName = App.Path & "\" & mapX & mapArea & mapY & ".map"
    End If
    
    'read the textures and the walkable values from the map file
    Open MapName For Input As #1
        For g = 0 To 899
            Input #1, Texture(g), Walkable(g), mapXStored(g), mapYStored(g), mapAreaStored(g)
        Next g
    Close
    
    'clear the picture box which will hold the back ground
    picBack.Cls
    
    X = 0
    Y = 0
    
    'loop through each tile, getting it with bitblt from frmTextures, and putting it into
    'the picBack pic box.
    For g = 0 To 899
        tileLeft(g) = X
        tileTop(g) = Y
        Call BitBlt(picBack.hdc, X, Y, 40, 40, frmTextures.picTextures(Texture(g)).hdc, 0, 0, SRCCOPY)
        Y = Y + 40
        
        'if a column has been finished, goto the next one
        If Y >= 1200 Then
            Y = 0
            X = X + 40
        End If
    Next g
    
    'by-pass error handler
    GoTo endsub
    
    'for errors
errHandler:
    
    MsgBox "Error number " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Dragon Lore"
    MsgBox MapName & " was not found or was corrupted.  Please re-install this program."
    End

endsub:
End Sub

Sub redrawPic()

    'this function draws the picture to the screen.

    'black out the old picture
    picBuf.Cls
    'Copy the back ground to the buffer pic box
    Call BitBlt(picBuf.hdc, 0, 0, 2900, 9500, picBack.hdc, screenX, screenY, SRCCOPY)
    'Copy the first layer of the sprite to the buffer
    'this mask is like a negative, it is a black shadow of the character,
    'sourouneded by white (see picSpr). when added using SRCAND, every other color
    'except black becomes transparent.  So only the black figure is left
    Call BitBlt(picBuf.hdc, charPosX, charPosY, 32, 56, picSpr.hdc, animX + 33, animY, SRCAND)
    'Copy the second layer of the sprite to the buffer, for transparent effect.
    'this copys the color onto the black, using SRCINVERT, which makes it like copying
    'the colors onto white, so the colors stay the same.
    Call BitBlt(picBuf.hdc, charPosX, charPosY, 32, 56, picSpr.hdc, animX, animY, SRCINVERT)
    'refresh the picture
    picBuf.Refresh

End Sub

Private Function touching() As Integer
    Dim g As Integer ' counting variable
    Dim tmpX As Integer
    Dim tmpY As Integer
    
    
    'this looks at the direction the character is moving, and sees if the next step
    'will put the character onto a tile which has a walkable value of 1, which is
    'either water trees or a building.  If it is, it returns 1. if not, it returns 0.
    
    tmpX = 0
    tmpY = 0
    
    'check each tile
    'I'm looking for ways to OPTIMIZE this!! Email me with suggestions!
    For g = 0 To 899
        'only proceed to check a tile if it is within a certain radius of the character,
        'and if it is a tree/water/wall
        If Abs((charX + 16) - (tileLeft(g) + 20)) < 160 And Abs((charY + 16) - (tileTop(g) + 20)) < 160 And (Walkable(g) = 1 Or Walkable(g) = 2) Then
            If direction = dLEFT Then   'if the character is walking left
                'check the left side of the character
                If charX - Speed > tileLeft(g) And charX - Speed < tileLeft(g) + 40 Then
                    'check the lower left corner
                    If charY + 50 > tileTop(g) And charY + 50 < tileTop(g) + 40 Then
                        GoTo endsub
                    'check the top left corner
                    ElseIf charY + 35 > tileTop(g) And charY + 35 < tileTop(g) + 40 Then
                        GoTo endsub
                    End If
                End If
            ElseIf direction = dUP Then 'if the character is walking up
                'check the top side of the character
                If charY + 30 - Speed > tileTop(g) And charY + 30 - 10 < tileTop(g) + 40 Then
                    'check the top right corner
                    If charX + 30 > tileLeft(g) And charX + 30 < tileLeft(g) + 40 Then
                        GoTo endsub
                    'check to top left corner
                    ElseIf charX > tileLeft(g) And charX < tileLeft(g) + 40 Then
                        GoTo endsub
                    End If
                End If
            ElseIf direction = dRIGHT Then  'if the character is walking right
                'check the right side of the character
                If charX + 30 + Speed > tileLeft(g) And charX + 30 + Speed < tileLeft(g) + 40 Then
                    'check the right top corner
                   'check the lower left corner
                    If charY + 50 > tileTop(g) And charY + 50 < tileTop(g) + 40 Then
                        GoTo endsub
                    'check the top left corner
                    ElseIf charY + 35 > tileTop(g) And charY + 35 < tileTop(g) + 40 Then
                        GoTo endsub
                    End If
                End If
            ElseIf direction = dDOWN Then   'if the character is walking down
                'check the bottom side of the character
                If charY + 50 + Speed > tileTop(g) And charY + 50 + Speed < tileTop(g) + 40 Then
                    'check the bottom right side
                    If charX + 30 > tileLeft(g) And charX + 30 < tileLeft(g) + 40 Then
                        GoTo endsub
                    'check to bottom left corner
                    ElseIf charX > tileLeft(g) And charX < tileLeft(g) + 40 Then
                        GoTo endsub
                    End If
                End If
            End If
        End If
    Next g

    touching = 0

    GoTo endFunct

endsub:

    If Walkable(g) = 1 Then
        touching = 1
    ElseIf Walkable(g) = 2 Then
        moving = 0
        mapX = mapXStored(g)
        mapY = mapYStored(g)
        mapArea = mapAreaStored(g)
        BuildBack
    End If

endFunct:

End Function



'delays the system for 'wait' amount of time
Sub delay()
    Dim a As Double, b As Double
    a = Timer
    b = Timer
    While a < b + wait
        a = Timer
    Wend
End Sub

Private Sub picBuf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'turn the mouse icon into the visible icon
    frmDisplay.MouseIcon = frmTextures.picVisible.Picture
End Sub

