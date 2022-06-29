VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dixuSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' dixuSprite class v 0.3, Copyright Patrice Scribe, 1997
' http://ourworld.compuserve.com/homepages/pscribe or
' http://www.chez.com/scribe
' Contributions :
' ? - Zoom & clipping feature (not fully implemented yet)

Option Explicit

Public Visible As Boolean               ' Visible property
Public Bounce As Long                   ' Bounce behavior
Public Key As String                    ' Identification key
Public X As Long                        ' x position
Public y As Long                        ' y position
Public VelocityX As Long                ' Velocity along the X axis (not yet pixels/s)
Public VelocityY As Long                ' Velocity along the Y axis (not yet pixels/s)
Public Surface As DirectDrawSurface2
Public Width As Long                    ' Width of sprite in pixels
Public Height As Long                   ' Height of sprite in pixels
Public Zoom As Single                   ' Zoom sprite

'Private ddsSurface As DirectDrawSurface2 ' Source for sprite bitmap
Private SurfaceHeight As Long
Private SurfaceWidth As Long
Private Top As Long                     ' Top of the source image
Private Left As Long                    ' Left of the source image
Private BltFlag As Long
Private BltFastFlag As Long
Private fx As DDBLTFX

Private Sub Class_Initialize()
    ' Register the new sprite in a sprites collection
    fx.dwSize = Len(fx)
    Bounce = 1
    Zoom = 1
    Key = CStr(Timer)
2:
    On Error GoTo 1
    dixuSprites.Add Me, Key
    Visible = True
Exit Sub
1:
    Key = Val(Key) + (Rnd * 6)
    GoTo 2
End Sub

Private Sub Class_Terminate()
    Set Surface = Nothing
End Sub

Public Sub Move()
    ' Move according to elapsed time
    X = X + VelocityX '/ (1 / (dixuTime - dixuLastTime))
    y = y + VelocityY '/ (1 / (dixuTime - dixuLastTime))
    ' Boucing tests
    dixuBounceTest X, VelocityX, CLng(Zoom * Width), ScreenRect.Right
    dixuBounceTest y, VelocityY, CLng(Zoom * Height), ScreenRect.bottom
End Sub

' Paint the sprite to the back buffer
Public Sub Paint()
    Dim r As RECT
    Dim d As RECT
    On Error Resume Next
    If SurfaceWidth = 0 Then
        Dim ddsd As DDSURFACEDESC
        ddsd.dwSize = Len(ddsd)
        Surface.GetSurfaceDesc ddsd
        SurfaceWidth = ddsd.dwWidth
        SurfaceHeight = ddsd.dwHeight
    End If
    r.Top = Top
    r.Left = Left
    r.Right = r.Left + Width
    r.bottom = r.Top + Height
    d.Top = y
    d.Left = X
    d.Right = d.Left + Width * Zoom
    d.bottom = d.Top + Height * Zoom
    If d.Right > ScreenRect.Right Then
        r.Right = r.Left + Width * (ScreenRect.Right / d.Right)
        d.Right = ScreenRect.Right - 1
    End If
    If d.bottom > ScreenRect.bottom Then
        r.bottom = r.Top + Height * (ScreenRect.bottom / d.bottom)
        d.bottom = ScreenRect.bottom - 1
    End If
    If X < 0 Then X = 0
    If X + Width > ScreenRect.Right Then X = ScreenRect.Right - Width
    If y < 0 Then
        VelocityY = -VelocityY
        Move
    End If
    If y + Height > ScreenRect.bottom Then
        VelocityY = -VelocityY
        Move
    End If
    If Visible = True Then
            ' Blit image to the back buffer
            If Zoom = 1 Then
                dixuBackBuffer.BltFast X, y, Surface, r, BltFastFlag Or DDBLTFAST_WAIT
            Else
                Dim fx As DDBLTFX
                fx.dwSize = Len(fx)
                fx.dwRop = &HCC0020
                dixuBackBuffer.Blt d, Surface, r, BltFlag Or DDBLT_WAIT Or DDBLT_ROP, fx
            End If
            ' Next image position
            Left = (Left + Width) Mod SurfaceWidth
            ' If row was done, next line and back to top if done
            If Left = 0 Then Top = (Top + Height) Mod SurfaceHeight
    End If
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & ". " & dixuErrorString(Err.Number)
End Sub

Public Property Let ColorKey(ByVal Color As Long)
    Dim ddck As DDCOLORKEY
    ddck.dwColorSpaceLowValue = Color
    ddck.dwColorSpaceHighValue = ddck.dwColorSpaceLowValue
    Surface.SetColorKey DDCKEY_SRCBLT, ddck
End Property

Public Property Let UseColorKey(ByVal bln As Boolean)
    If bln Then
        BltFlag = DDBLT_KEYSRCOVERRIDE
        BltFastFlag = DDBLTFAST_SRCCOLORKEY
    Else
        BltFlag = 0
        BltFastFlag = 0
    End If
End Property

' ********** Private **********

' Boucing test
Private Sub dixuBounceTest(ByRef Where As Long, ByRef Velocity As Long, ByVal Size As Long, ByVal Limit As Long)
    ' If sprite goes right or down
    If Velocity > 0 Then
        ' If bottom right corner out of screen
        If Where > Limit - Size Then
            ' Bounce
            Where = Limit - Size - (Where - (Limit - Size))
            ' Reverse speed
            Velocity = -Bounce * Velocity
        End If
    ' Else sprite goes top or left
        ' If top left corner out of screen
        If Where < 0 Then
            ' Bounce
            Where = -Where
            ' Reverse speed
            Velocity = -Bounce * Velocity
        End If
    End If
End Sub
