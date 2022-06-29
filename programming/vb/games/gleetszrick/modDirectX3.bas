Attribute VB_Name = "modDirectX3"
' dixu v 0.3, Copyright Patrice Scribe, 1997
' http://ourworld.compuserve.com/homepages/pscribe or
' http://www.chez.com/scribe
' Changes from v 0.2 :
' - windowed mode support
' - support for sprite transparency and zoom property
' Contributions :
' ? - zoom support and clipping (not fully implemented yet)
' Note that there are few changes to make for use with DirectX.tlb
' (designed for DirectX3.tlb or DirectX5.tlb)
' Note also that stdole2.tlb must be referenced

Option Explicit

' ***** dixu, public
Global Const dixuFaceDown = 0   ' Down face (floor)
Global Const dixuFaceTop = 1    ' Top face (ceiling)
Global Const dixuFaceFront = 5  ' Front face
Global Const dixuFaceLeft = 3   ' Left face
Global Const dixuFaceRight = 4  ' Right face
Global Const dixuFaceBack = 2   ' Back face (related to the subjective view)

' Flags for dixuInit
Global Const dixuInit3DDevice = 1       ' 3D enabled
Global Const dixuInitFullScreen = 2     ' Full screen mode
Global Const dixuInitRGB = 4            ' RGB driver
Global Const dixuInitMono = 8           ' Mono driver
Global Const dixuInitWindowProc = 16    ' Handles Windows messages

Public dixuAppEnd As Boolean    ' Esc detected in dixuCameraMove

' DirectDraw public objects
Public dixuDDraw As DirectDraw2
Public dixuPrimarySurface As DirectDrawSurface2
Public dixuBackBuffer As DirectDrawSurface2
Public dixuClipper As DirectDrawClipper
Public ScreenRect As RECT   ' Use MyFrm.ScaleWidth, MyFrm.ScaleHeight instead...

' Direct3D Retained Mode public objects
Public dixuD3DRM As Direct3DRM
Public dixuD3DRMDevice As Direct3DRMDevice
Public dixuD3DRMViewport As Direct3DRMViewPort

' High-level 3D objects
Public dixuCamera As Direct3DRMFrame
Public dixuScene As Direct3DRMFrame

' Time values
Public dixuTime As Single
Public dixuLastTime As Single

' Don't compile if dixuSprite not needed
#If DIXU_NOSPRITE = 0 Then
Public dixuSprites As New Collection
#End If

' ***** Win32, public
Public Const SRCCOPY = &HCC0020
Public Const TRANSPARENT = 1    ' For SetBkMode

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

' ***** Win32, private
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)

Private Const IMAGE_BITMAP = 0
Private Const LR_LOADFROMFILE = &H10
Private Const LR_CREATEDIBSECTION = &H2000

Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33

' SetWindowPos
Private Const SWP_NOSIZE = &H1  ' Don't size
Private Const SWP_NOMOVE = &H2  ' Don't move
Private Const HWND_TOPMOST = -1 ' Topmost
Private Const HWND_NOTOPMOST = -2


' Windows messages
Private Const WM_ACTIVATE = &H6
Private Const WM_MOVE = &H3
Private Const WM_PAINT = &HF

' Windows style
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
'Public Const WS_CHILD = &H40000000
'Public Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
'Public Const WS_DISABLED = &H8000000
'Public Const WS_CLIPSIBLINGS = &H4000000
'Public Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
'Public Const WS_DLGFRAME = &H400000
'Public Const WS_VSCROLL = &H200000
'Public Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
'Public Const WS_GROUP = &H20000
'Public Const WS_TABSTOP = &H10000
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
'Private Const GWL_STYLE = -16
'Public Const WS_TILED = WS_OVERLAPPED
'Public Const WS_ICONIC = WS_MINIMIZE
'Public Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
'Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

'
'   Common Window Styles
'  /


'Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

'Public Const WS_CHILDWINDOW = (WS_CHILD)

Private Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type

Private Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved As Byte
End Type

Private Type POINTAPI
        X As Long
        y As Long
End Type

' GDI32
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' KERNEL32
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' USER32
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal First As Long, ByVal Size As Long, ptr As Any) As Long
Private Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' dixu, private
Public Const Pi = 3.14159265358

Private bln3DDevice As Boolean      ' dixuInit3DDevice specified ?
Private blnFullScreen As Boolean    ' dixuInitFullScreen specified ?
Private SpritesRect As RECT         ' Not used yet
Public blnBackBufferClear As Boolean ' Back buffer to clear ?

' Camera values
Private sngCameraStep As Single ' Moving forward (or backward)
Private sngCameraCos As Single  ' For camera rotation
Private sngCameraSin As Single  ' For camera rotation

Private lpPrevWndProc As Long   ' Window proc
'Private dixuFrm As Form         ' Form used
Private PrevWindowStyle As Long   ' Window previous style

' Initializes DirectX
Sub dixuInit(ByVal Flags As Long, frm As Form, ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long)
    Dim ddsd As DDSURFACEDESC
    Dim ddc As DDSCAPS
    ' Camera default values
    'Set dixuFrm = frm
    If PrevWindowStyle = 0 Then
        PrevWindowStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    End If
    sngCameraStep = 3
    sngCameraCos = Cos(10 * Pi / 180)
    sngCameraSin = Sin(10 * Pi / 180)
    ' 3D enabled ?
    bln3DDevice = (Flags And dixuInit3DDevice) <> 0
    ' Full screen mode ?
    blnFullScreen = (Flags And dixuInitFullScreen) <> 0
    ' Initializes DirectDraw if not already done
    If dixuDDraw Is Nothing Then
        DirectDrawCreate ByVal 0&, dixuDDraw, Nothing
    End If
    ' Clear other DirectDraw objects if needed
    Set dixuD3DRMViewport = Nothing
    Set dixuD3DRMDevice = Nothing
    Set dixuBackBuffer = Nothing
    Set dixuPrimarySurface = Nothing
    ' DirectDraw part
    If blnFullScreen Then
        ' Full screen without borders, captions, boxes...
        SetWindowLong frm.hwnd, GWL_STYLE, WS_POPUP Or WS_VISIBLE
        ' Topmost level
        SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        ' Full screen mode : change display mode
        dixuDDraw.SetCooperativeLevel frm.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
        dixuDDraw.SetDisplayMode Width, Height, BitsPerPixel, 0, 0
        With ddsd
            ' Structure size
            .dwSize = Len(ddsd)
            ' Use DDSD_CAPS and BackBufferCount
            .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
            With .DDSCAPS
                .dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_SYSTEMMEMORY
                ' If 3D enabled
                If bln3DDevice Then .dwCaps = .dwCaps Or DDSCAPS_3DDEVICE
            End With
            ' One back buffer
            .dwBackBufferCount = 1
        End With
        ' Creates buffers
        dixuDDraw.CreateSurface ddsd, dixuPrimarySurface, Nothing
        ' Retrieve back buffer
        ddc.dwCaps = DDSCAPS_BACKBUFFER
        dixuPrimarySurface.GetAttachedSurface ddc, dixuBackBuffer
        ' Keep screen rect
        ScreenRect.Left = 0
        ScreenRect.Top = 0
        ScreenRect.Right = Width - 1
        ScreenRect.bottom = Height - 1
    Else
        ' Windowed mode
        ' Restore original style (in case return back from full screen)
        SetWindowLong frm.hwnd, GWL_STYLE, PrevWindowStyle
        ' Restore Z order
        SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
        dixuDDraw.RestoreDisplayMode
        'SetWindowLong
        dixuDDraw.SetCooperativeLevel frm.hwnd, DDSCL_NORMAL
        ' Create the front buffer
        With ddsd
            .dwSize = Len(ddsd)
            .dwFlags = DDSD_CAPS
            .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
        End With
        dixuDDraw.CreateSurface ddsd, dixuPrimarySurface, Nothing
        ' Create back buffer
        With ddsd
            .dwSize = Len(ddsd)
            .dwFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
            .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            If bln3DDevice Then .DDSCAPS.dwCaps = .DDSCAPS.dwCaps Or DDSCAPS_3DDEVICE
            .dwWidth = frm.ScaleWidth
            .dwHeight = frm.ScaleHeight
        End With
        dixuDDraw.CreateSurface ddsd, dixuBackBuffer, Nothing
        ' Create and attach clipper (allow proper operation when window is covered)
        ' Just try without the next three lines, you'll understand !
        'dixuPrimarySurface.AddAttachedSurface dixuBackBuffer
        dixuDDraw.CreateClipper 0, dixuClipper, Nothing
        dixuClipper.SetHWnd 0, frm.hwnd
        dixuPrimarySurface.SetClipper dixuClipper
        
        ' Keep screen rect
        ScreenRect.Left = 0
        ScreenRect.Top = 0
        ScreenRect.Right = frm.ScaleWidth
        ScreenRect.bottom = frm.ScaleHeight
    End If
    ' Now work on Direct3D part
    If bln3DDevice Then
        Dim d3d As Direct3D
        Dim fds As D3DFINDDEVICESEARCH
        Dim fdr As D3DFINDDEVICERESULT
        Dim ddsdFront As DDSURFACEDESC
        If d3d Is Nothing Then
            ' Get D3D interface (QueryInterface stuff for C/C++)
            Set d3d = dixuDDraw
        End If
        ' Search for the color model driver
        fds.dwSize = Len(fds)
        fds.dwFlags = D3DFDS_COLORMODEL
        If (Flags And dixuInitRGB) <> 0 Then
            fds.dcmColorModel = D3DCOLOR_RGB
        Else
            fds.dcmColorModel = D3DCOLOR_MONO
        End If
        fdr.dwSize = Len(fdr)
        ' Find the driver
        d3d.FindDevice fds, fdr
        ' If 256 colors, set the palette
        If BitsPerPixel = 8 Then
            Dim ColorTable(0 To 255) As PALETTEENTRY
            Dim i As Long
            Dim Palette As DirectDrawPalette
            Debug.Print GetSystemPaletteEntries(frm.hdc, 0, 256, ColorTable(0))
            For i = 0 To 255
                ColorTable(i).peFlags = &H40
            Next
            dixuDDraw.CreatePalette 4, ColorTable(0), Palette, Nothing
            dixuBackBuffer.SetPalette Palette
        End If
        ' If needed, create top-level objects
        If dixuD3DRM Is Nothing Then
            Direct3DRMCreate dixuD3DRM
            dixuD3DRM.CreateFrame Nothing, dixuScene
            dixuD3DRM.CreateFrame dixuScene, dixuCamera
        End If
        ' Create the device from the existing DirectDraw back buffer
        ' fdr.mguid when used with DirectX.tlb
        ' fdr.GUID when used with DirectX3.tlb or DirectX5.tlb
        If blnFullScreen Then
            dixuD3DRM.CreateDeviceFromSurface ByVal 0&, dixuDDraw, dixuBackBuffer, dixuD3DRMDevice
        Else
            dixuD3DRM.CreateDeviceFromClipper dixuClipper, fdr.GUID, frm.ScaleWidth, frm.ScaleHeight, dixuD3DRMDevice
            ' Enables window messages processing
            If (Flags And dixuInitWindowProc) <> 0 Then
                lpPrevWndProc = GetWindowLong(frm.hwnd, GWL_WNDPROC)
                SetWindowLong frm.hwnd, GWL_WNDPROC, AddressOf WindowProc
            End If
        End If
        dixuD3DRM.CreateViewport dixuD3DRMDevice, dixuCamera, 0, 0, dixuD3DRMDevice.GetWidth - 1, dixuD3DRMDevice.GetHeight - 1, dixuD3DRMViewport
    End If
    frm.Show
    blnBackBufferClear = True
End Sub

' Clean up objects
Public Sub dixuDone()
    Dim i As Long
    ' Clear sprites (don't compile if dixuSprite not needed)
    #If DIXU_NOSPRITE = 0 Then
    ' Clear sprites
    'Dim Sprite As dixuSprite
    While dixuSprites.Count > 0
        dixuSprites.Remove 1
    Wend
    #End If
    ' Reset display mode
    If blnFullScreen Then
        Dim dd As DirectDraw
        'Set dd = dixuDDraw
        dixuDDraw.FlipToGDISurface
        'Set dd = Nothing
        dixuDDraw.RestoreDisplayMode
    End If
    dixuDDraw.SetCooperativeLevel 0, DDSCL_NORMAL
    If bln3DDevice Then
        ' Clear 3D objects
        Set dixuCamera = Nothing
        Set dixuScene = Nothing
        Set dixuD3DRMViewport = Nothing
        Set dixuD3DRMDevice = Nothing
        Set dixuD3DRM = Nothing
    End If
    ' Clear DirectDraw objects
    Set dixuClipper = Nothing
    Set dixuBackBuffer = Nothing
    Set dixuPrimarySurface = Nothing
    Set dixuDDraw = Nothing
End Sub

' Clears the back buffer
Public Sub dixuBackBufferClear()
    Dim fx As DDBLTFX
    With fx
        .dwSize = Len(fx)
        .dwFillColor = RGB(0, 0, 0)
    End With
    ' ScreenRect is necessary when used with DirectX.tlb
    dixuBackBuffer.Blt ByVal 0&, Nothing, ByVal 0&, DDBLT_COLORFILL, fx
    ' Buffer already clear
    blnBackBufferClear = False
End Sub
' Draws the back buffer
Public Sub dixuBackBufferDraw()
    On Error GoTo dixuBackBufferDraw_Error
    If bln3DDevice Then
        ' Render 3D scene
        dixuScene.Move 1
        dixuD3DRMViewport.Clear
        dixuD3DRMViewport.Render dixuScene
        dixuD3DRMDevice.Update
        'dixuD3DRM.Tick 1
        'While dixuD3DRMDevice.Get
        'Do
            'dixuBackBuffer.GetBltStatus 0
        'Loop Until Err.Number = 0
    End If
    dixuBackBufferClear
    ' Render sprites (don't compile if dixuSprite not needed)
    #If DIXU_NOSPRITE = 0 Then
    If dixuSprites.Count <> 0 Then dixuSpritesDraw
    #End If
    ' Render the 2D scene
    If blnFullScreen Then
        ' Workaround for DirectDrawSurface2.Flip bug
        Dim dds As DirectDrawSurface
        Set dds = dixuPrimarySurface
        dds.Flip Nothing, DDFLIP_WAIT
        Set dds = Nothing
    Else
        Dim fx As DDBLTFX
        fx.dwSize = Len(fx)
        fx.dwRop = SRCCOPY
        Dim dixuClientRect As RECT
        ' Top of the client area is calculated from the bottom by substracting the client area height and the window frame height
        'dixuClientRect.Top = (dixuFrm.Top + dixuFrm.Height) / Screen.TwipsPerPixelY - GetSystemMetrics(SM_CYFRAME) - dixuFrm.ScaleHeight
        'dixuClientRect.bottom = dixuClientRect.Top + dixuFrm.ScaleHeight
        'dixuClientRect.Left = dixuFrm.Left / Screen.TwipsPerPixelY + GetSystemMetrics(SM_CXFRAME)
        'dixuClientRect.Right = dixuClientRect.Left + dixuFrm.ScaleWidth
        If Not bln3DDevice Then
            GetClientRect Screen.ActiveForm.hwnd, dixuClientRect
            ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Left
            ClientToScreen Screen.ActiveForm.hwnd, dixuClientRect.Right
            dixuPrimarySurface.Blt dixuClientRect, dixuBackBuffer, ByVal 0&, DDBLT_ROP Or DDBLT_WAIT, fx
        End If
        
    End If
    blnBackBufferClear = True
    Exit Sub
dixuBackBufferDraw_Error:
    Debug.Print Now, Err.Description
    If Err.Number = DDERR_SURFACELOST Or Err.Number = DDERR_SURFACELOST - &H100000 Then
        dixuBackBuffer.Restore
        dixuPrimarySurface.Restore
    End If
    Err.Clear
    Resume
End Sub

Function dixuErrorString(ByVal AutomationError As Long) As String
    Select Case AutomationError
        ' DirectDraw errors
        Case DDERR_ALREADYINITIALIZED: dixuErrorString = "DDERR_ALREADYINITIALIZED"
        Case DDERR_BLTFASTCANTCLIP: dixuErrorString = "DDERR_BLTFASTCANTCLIP"
        Case DDERR_CANNOTATTACHSURFACE: dixuErrorString = "DDERR_CANNOTATTACHSURFACE"
        Case DDERR_CANTCREATEDC: dixuErrorString = "DDERR_CANTCREATEDC"
        Case DDERR_CANTDUPLICATE: dixuErrorString = "DDERR_CANTDUPLICATE"
        Case DDERR_CANTLOCKSURFACE: dixuErrorString = "DDERR_CANTLOCKSURFACE"
        Case DDERR_CANTPAGELOCK: dixuErrorString = "DDERR_CANTPAGELOCK"
        Case DDERR_CANTPAGEUNLOCK: dixuErrorString = "DDERR_CANTPAGEUNLOCK"
        Case DDERR_CLIPPERISUSINGHWND: dixuErrorString = "DDERR_CLIPPERISUSINGHWND"
        Case DDERR_COLORKEYNOTSET: dixuErrorString = "DDERR_COLORKEYNOTSET"
        Case DDERR_CURRENTLYNOTAVAIL: dixuErrorString = "DDERR_CURRENTLYNOTAVAIL"
        Case DDERR_DCALREADYCREATED: dixuErrorString = "DDERR_DCALREADYCREATED"
        Case DDERR_DEVICEDOESNTOWNSURFACE: dixuErrorString = "DDERR_DEVICEDOESNTOWNSURFACE"
        Case DDERR_DIRECTDRAWALREADYCREATED: dixuErrorString = "DDERR_DIRECTDRAWALREADYCREATED"
        Case DDERR_EXCEPTION: dixuErrorString = "DDERR_EXCEPTION"
        Case DDERR_EXCLUSIVEMODEALREADYSET: dixuErrorString = "DDERR_EXCLUSIVEMODEALREADYSET"
        'Case DDERR_GENERIC: dixuErrorString = "DDERR_GENERIC"
        Case DDERR_HEIGHTALIGN: dixuErrorString = "DDERR_HEIGHTALIGN"
        Case DDERR_HWNDALREADYSET: dixuErrorString = "DDERR_HWNDALREADYSET"
        Case DDERR_HWNDSUBCLASSED: dixuErrorString = "DDERR_HWNDSUBCLASSED"
        Case DDERR_IMPLICITLYCREATED: dixuErrorString = "DERR_IMPLICITYCREATED"
        Case DDERR_INCOMPATIBLEPRIMARY: dixuErrorString = "DERR_INCOMPATIBLEPRIMARY"
        Case DDERR_INVALIDCAPS: dixuErrorString = "DDERR_INVALIDCAPS"
        Case DDERR_INVALIDCLIPLIST: dixuErrorString = "DERR_INVALIDCLIPLIST"
        Case DDERR_INVALIDDIRECTDRAWGUID: dixuErrorString = "DDERR_INVALIDDIRECTDRAWGUID"
        Case DDERR_INVALIDMODE: dixuErrorString = "DERR_INVALIDMODE"
        Case DDERR_INVALIDOBJECT: dixuErrorString = "DDERR_INVALIDOBJECT"
        'Case DDERR_INVALIDPARAMS: dixuErrorString = "DDERR_INVALIDPARAMS"
        Case DDERR_INVALIDPIXELFORMAT: dixuErrorString = "DDERR_INVALIDPIXELFORMAT"
        Case DDERR_INVALIDPOSITION: dixuErrorString = "DDERR_INVALIDPOSITION"
        Case DDERR_INVALIDRECT: dixuErrorString = "DDERR_INVALIDRECT"
        Case DDERR_INVALIDSURFACETYPE: dixuErrorString = "DDERR_INVALIDSURFACETYPE"
        Case DDERR_LOCKEDSURFACES: dixuErrorString = "DDERR_LOCKEDSURFACES"
        Case DDERR_MOREDATA: dixuErrorString = "DDERR_MOREDATA"
        Case DDERR_NO3D: dixuErrorString = "DDERR_NO3D"
        Case DDERR_NOALPHAHW: dixuErrorString = "DDERR_NOALPHAHW"
        Case DDERR_NOBLTHW: dixuErrorString = "DDERR_NOBLTHW"
        Case DDERR_NOCLIPLIST: dixuErrorString = "DDERR_NOCLIPLIST"
        Case DDERR_NOCLIPPERATTACHED: dixuErrorString = "DDERR_NOCLIPPERATTACHED"
        Case DDERR_NOCOLORCONVHW: dixuErrorString = "DDERR_NOCOLORCONVHW"
        Case DDERR_NOCOLORKEY: dixuErrorString = "DDERR_NOCOLORKEY"
        Case DDERR_NOCOLORKEYHW: dixuErrorString = "DDERR_NOCOLORKEYHW"
        Case DDERR_NOCOOPERATIVELEVELSET: dixuErrorString = "DDERR_NOCOOPERATIVELEVELSET"
        Case DDERR_NODC: dixuErrorString = "DDERR_NODC"
        Case DDERR_NODDROPSHW: dixuErrorString = "DDERR_NODDROPSHW"
        Case DDERR_NODIRECTDRAWHW: dixuErrorString = "DDERR_NODIRECTDRAWHW"
        Case DDERR_NODIRECTDRAWSUPPORT: dixuErrorString = "DDERR_NODIRECTDRAWSUPPORT"
        Case DDERR_NOEMULATION: dixuErrorString = "DDERR_NOEMULATION"
        Case DDERR_NOEXCLUSIVEMODE: dixuErrorString = "DDERR_NOEXCLUSIVEMODE"
        Case DDERR_NOFLIPHW: dixuErrorString = "DDERR_NOFLIPHW"
        Case DDERR_NOHWND: dixuErrorString = "DDERR_NOHWND"
        Case DDERR_NOMIPMAPHW: dixuErrorString = "DDERR_NOMIPMAPHW"
        Case DDERR_NOMIRRORHW: dixuErrorString = "DDERR_NOMIRRORHW"
        Case DDERR_NONONLOCALVIDMEM: dixuErrorString = "DDERR_NONONLOCALVIDMEM"
        Case DDERR_NOOPTIMIZEHW: dixuErrorString = "DDERR_NOOPTIMIZEHW"
        Case DDERR_NOOVERLAYDEST: dixuErrorString = "DDERR_NOOVERLAYDEST"
        Case DDERR_NOOVERLAYHW: dixuErrorString = "DDERR_NOOVERLAYHW"
        Case DDERR_NOPALETTEATTACHED: dixuErrorString = "DDERR_NOPALETTEATTACHED"
        Case DDERR_NOPALETTEHW: dixuErrorString = "DDERR_NOPALETTEHW"
        Case DDERR_NORASTEROPHW: dixuErrorString = "DDERR_NORASTEROPHW"
        Case DDERR_NOROTATIONHW: dixuErrorString = "DDERR_NOROTATIONHW"
        Case DDERR_NOSTRETCHHW: dixuErrorString = "DDERR_NOSTRETCHHW"
        Case DDERR_NOT4BITCOLOR: dixuErrorString = "DDERR_NOT4BITCOLOR"
        Case DDERR_NOT4BITCOLORINDEX: dixuErrorString = "DDERR_NOT4BITCOLORINDEX"
        Case DDERR_NOT8BITCOLOR: dixuErrorString = "DDERR_NOT8BITCOLOR"
        Case DDERR_NOTAOVERLAYSURFACE: dixuErrorString = "DDERR_NOTAOVERLAYSURFACE"
        Case DDERR_NOTEXTUREHW: dixuErrorString = "DDERR_NOTEXTUREHW"
        Case DDERR_NOTFLIPPABLE: dixuErrorString = "DDERR_NOTFLIPPABLE"
        Case DDERR_NOTFOUND: dixuErrorString = "DDERR_NOTFOUND"
        'Case DDERR_NOTINITIALIZED: dixuErrorString = "DDERR_NOTINITIALIZED"
        Case DDERR_NOTLOADED: dixuErrorString = "DDERR_NOTLOADED"
        Case DDERR_NOTLOCKED: dixuErrorString = "DDERR_NOTLOCKED"
        Case DDERR_NOTPAGELOCKED: dixuErrorString = "DDERR_NOTPAGELOCKED"
        Case DDERR_NOTPALETTIZED: dixuErrorString = "DDERR_NOTPALETTIZED"
        Case DDERR_NOVSYNCHW: dixuErrorString = "DDERR_NOVSYNCHW"
        Case DDERR_NOZBUFFERHW: dixuErrorString = "DDERR_NOZBUFFERHW"
        Case DDERR_NOZOVERLAYHW: dixuErrorString = "DDERR_NOZOVERLAYHW"
        Case DDERR_OUTOFCAPS: dixuErrorString = "DDERR_OUTOFCAPS"
        'Case DDERR_OUTOFMEMORY: dixuErrorString = "DDERR_OUTOFMEMORY"
        Case DDERR_OUTOFVIDEOMEMORY: dixuErrorString = "DDERR_OUTOFVIDEOMEMORY"
        Case DDERR_OVERLAYCANTCLIP: dixuErrorString = "DDERR_OVERLAYCANTCLIP"
        Case DDERR_OVERLAYCOLORKEYONLYONEACTIVE: dixuErrorString = "DDERR_OVERLAYCOLORKEYONLYONEACTIVE"
        Case DDERR_OVERLAYNOTVISIBLE: dixuErrorString = "DDERR_OVERLAYNOTVISIBLE"
        Case DDERR_PALETTEBUSY: dixuErrorString = "DDERR_PALETTEBUSY"
        Case DDERR_PRIMARYSURFACEALREADYEXISTS: dixuErrorString = "DDERR_PRIMARYSURFACEALREADYEXISTS"
        Case DDERR_REGIONTOOSMALL: dixuErrorString = "DDERR_REGIONTOOSMALL"
        Case DDERR_SURFACEALREADYATTACHED: dixuErrorString = "DDERR_SURFACEALREADYATTACHED"
        Case DDERR_SURFACEALREADYDEPENDENT: dixuErrorString = "DDERR_SURFACEALREADYDEPENDENT"
        Case DDERR_SURFACEBUSY: dixuErrorString = "DDERR_SURFACEBUSY"
        Case DDERR_SURFACEISOBSCURED: dixuErrorString = "DDERR_SURFACEISOBSCURED"
        Case DDERR_SURFACELOST: dixuErrorString = "DDERR_SURFACELOST"
        Case DDERR_SURFACENOTATTACHED: dixuErrorString = "DDERR_SURFACENOTATTACHED"
        Case DDERR_TOOBIGHEIGHT: dixuErrorString = "DDERR_TOOBIGHEIGHT"
        Case DDERR_TOOBIGSIZE: dixuErrorString = "DDERR_TOOBIGSIZE"
        Case DDERR_TOOBIGWIDTH: dixuErrorString = "DDERR_TOOBIGWIDTH"
        'Case DDERR_UNSUPPORTED: dixuErrorString = "DDERR_UNSUPPORTED"
        Case DDERR_UNSUPPORTEDFORMAT: dixuErrorString = "DDERR_UNSUPPORTEDFORMAT"
        Case DDERR_UNSUPPORTEDMASK: dixuErrorString = "DDERR_UNSUPPORTEDMASK"
        Case DDERR_UNSUPPORTEDMODE: dixuErrorString = "DDERR_UNSUPPORTEDMODE"
        Case DDERR_VERTICALBLANKINPROGRESS: dixuErrorString = "DDERR_VERTICALBLANKINPROGRESS"
        Case DDERR_VIDEONOTACTIVE: dixuErrorString = "DDERR_VIDEONOTACTIVE"
        Case DDERR_WASSTILLDRAWING: dixuErrorString = "DDERR_WASSTILLDRAWING"
        Case DDERR_WRONGMODE: dixuErrorString = "DDERR_WRONGMODE"
        Case DDERR_XALIGN: dixuErrorString = "DDERR_XALIGN"
        ' Direct3D
        Case D3DRMERR_BADALLOC: dixuErrorString = "D3DRMERR_BADALLOC"
        Case D3DRMERR_BADDEVICE: dixuErrorString = "D3DRMERR_BADDEVICE"
        Case D3DRMERR_BADFILE: dixuErrorString = "D3DRMERR_BADFILE"
        Case D3DRMERR_BADMAJORVERSION: dixuErrorString = "D3DRMERR_BADMAJORVERSION"
        Case D3DRMERR_BADMINORVERSION: dixuErrorString = "D3DRMERR_BADMINORVERSION"
        Case D3DRMERR_BADOBJECT: dixuErrorString = "D3DRMERR_BADOBJECT"
        Case D3DRMERR_BADPMDATA: dixuErrorString = "D3DRMERR_BADPMDATA"
        Case D3DRMERR_BADTYPE: dixuErrorString = "D3DRMERR_BADTYPE"
        Case D3DRMERR_BADVALUE: dixuErrorString = "D3DRMERR_BADVALUE"
        Case D3DRMERR_BOXNOTSET: dixuErrorString = "D3DRMERR_BOXNOTSET"
        Case D3DRMERR_CONNECTIONLOST: dixuErrorString = "D3DRMERR_CONNECTIONLOST"
        Case D3DRMERR_FACEUSED: dixuErrorString = "D3DRMERR_FACEUSED"
        Case D3DRMERR_FILENOTFOUND: dixuErrorString = "D3DRMERR_FILENOTFOUND"
        'Case D3DRMERR_INVALIDDATA: dixuErrorString = "D3DRMERR_INVALIDDATA"
        'Case D3DRMERR_INVALIDOBJECT: dixuErrorString = "D3DRMERR_INVALIDOBJECT"
        'Case D3DRMERR_INVALIDPARAMS: dixuErrorString = "D3DRMERR_INVALIDPARAMS"
        Case D3DRMERR_NOTDONEYET: dixuErrorString = "D3DRMERR_NOTDONEYET"
        Case D3DRMERR_NOTENOUGHDATA: dixuErrorString = "D3DRMERR_NOTENOUGHDATA"
        Case D3DRMERR_NOTFOUND: dixuErrorString = "D3DRMERR_NOTFOUND"
        Case D3DRMERR_PENDING: dixuErrorString = "D3DRMERR_PENDING"
        Case D3DRMERR_REQUESTTOOLARGE: dixuErrorString = "D3DRMERR_REQUESTTOOLARGE"
        Case D3DRMERR_REQUESTTOOSMALL: dixuErrorString = "D3DRMERR_REQUESTTOOSMALL"
        Case D3DRMERR_UNABLETOEXECUTE: dixuErrorString = "D3DRMERR_UNABLETOEXECUTE"
        Case Else: dixuErrorString = Hex$(AutomationError)
    End Select
End Function

'********** DirectDraw support **********

' Creates a DirectDraw surface from a file
Public Function dixuCreateSurface(ByVal Width As Long, ByVal Height As Long, ByVal strFile As String) As DirectDrawSurface2
    Dim frm As Form
    Dim Picture As StdPicture
    Dim PictureWidth As Long
    Dim PictureHeight As Long
    Dim ddsd As DDSURFACEDESC       ' Surface description
    Dim dds As DirectDrawSurface2   ' DirectDraw surface
    Dim hdcPicture As Long          ' Picture device context
    Dim hdcSurface As Long          ' Surface device context
    Set frm = Screen.ActiveForm
    ' Load picture
    Set Picture = LoadPicture(strFile)
    PictureWidth = frm.ScaleX(Picture.Width, vbHimetric, vbPixels)
    PictureHeight = frm.ScaleY(Picture.Height, vbHimetric, vbPixels)
    If Width = 0 Then Width = PictureWidth
    If Height = 0 Then Height = PictureHeight
    ' Fill surface description
    With ddsd
        .dwSize = Len(ddsd)
        .dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .dwWidth = Width
        .dwHeight = Height
    End With
    ' Create surface
    dixuDDraw.CreateSurface ddsd, dds, Nothing
    ' Create memory device
    hdcPicture = CreateCompatibleDC(ByVal 0&)
    ' Select the bitmap in this memory device
    SelectObject hdcPicture, Picture.Handle
    ' Restore the surface
    dds.Restore
    ' Get the surface's DC
    dds.GetDC hdcSurface
    ' Copy from the memory device to the DirectDrawSurface
    StretchBlt hdcSurface, 0, 0, Width, Height, hdcPicture, 0, 0, PictureWidth, PictureHeight, SRCCOPY
    ' Release the surface's DC
    dds.ReleaseDC hdcSurface
    ' Release the memory device and the bitmap
    DeleteDC hdcPicture
    'Set Picture = Nothing
    Set dixuCreateSurface = dds
End Function

' ********** Direct3D Retained Mode support **********

' >>>>>>>>>> Camera

' Set camera steps values
Public Sub dixuSetCameraMoves(ByVal Step As Single, ByVal Angle As Single)
    sngCameraStep = Step
    sngCameraCos = Cos(Angle * Pi / 180)
    sngCameraSin = Sin(Angle * Pi / 180)
End Sub

' Move camera according to KeyCode (features dectection collision code for moving forward - can still pass through walls when moving backward !)
Public Sub dixuCameraMove(ByVal KeyCode As Long)
    On Error GoTo dixuCameraMove_Error
    Select Case KeyCode
        Case vbKeyDown      ' Move backward
            dixuCamera.SetPosition dixuCamera, 0, 0, -sngCameraStep
        Case vbKeyEscape    ' Ends app
            dixuAppEnd = True
        Case vbKeyLeft      ' Turn left
            dixuCamera.SetOrientation dixuCamera, -sngCameraSin, 0, sngCameraCos, 0, 1, 0
        Case vbKeyRight     ' Turn right
            dixuCamera.SetOrientation dixuCamera, sngCameraSin, 0, sngCameraCos, 0, 1, 0
        Case vbKeyUp        ' Move forward
            Dim PickedArray As Direct3DRMPickedArray
            Dim MeshBuilder As Direct3DRMMeshBuilder
            Dim FrameArray As Direct3DRMFrameArray
            Dim PickDesc As D3DRMPICKDESC
            Dim Distance As Single
            Dim CameraPosition As D3DVECTOR
            Dim Screen As D3DRMVECTOR4D ' Screen coordinates
            Dim World As D3DVECTOR      ' World coordinates
            ' Retrieves intersected visuals
            dixuD3DRMViewport.Pick ScreenRect.Right \ 2, ScreenRect.bottom \ 2, PickedArray
            If PickedArray.GetSize <> 0 Then
                ' Retrieve the first visual, parent frame (?), and details
                PickedArray.GetPick 0, MeshBuilder, FrameArray, PickDesc
                ' Copy screen coordinates
                With PickDesc.vPosition
                    Screen.X = .X
                    Screen.y = .y
                    Screen.z = .z
                    Screen.w = 1
                End With
                ' Transform to world coordinates
                dixuD3DRMViewport.InverseTransform World, Screen
                ' Get camera position
                dixuCamera.GetPosition dixuScene, CameraPosition
                ' Compute distance between intersection and camera
                Distance = Sqr((CameraPosition.X - World.X) ^ 2 + (CameraPosition.z - World.z) ^ 2)
                ' Enough distance : move
                If Distance > 4 * sngCameraStep Then dixuCamera.SetPosition dixuCamera, 0, 0, sngCameraStep
            Else
                ' No visual : move
                dixuCamera.SetPosition dixuCamera, 0, 0, sngCameraStep
            End If
            ' Collision detection defeated if key pressed repeatidly and scene not rendered ?!
            'dixuBackBufferDraw
    End Select
    Exit Sub
' Trap transient errors (division by zero)
dixuCameraMove_Error:
    Debug.Print Hex$(Err.Number)
    Resume
End Sub

' Creates a room (a cube whose faces are visible from inside)
Public Function dixuCreateRoom() As Direct3DRMMeshBuilder
    Dim aVertices(0 To 8) As D3DVECTOR  ' Vertices array
    Dim aNormals(0) As D3DVECTOR        ' Normal array (not used)
    Dim aFaces(1 To 31) As Long         ' Faces array
    Dim MeshBuilder As Direct3DRMMeshBuilder
    Dim i As Integer
    ' Coordinates for floor vertices
    aVertices(0).X = -0.5
    aVertices(0).y = 0
    aVertices(0).z = -0.5
    aVertices(1).X = -0.5
    aVertices(1).y = 0
    aVertices(1).z = 0.5
    aVertices(2).X = 0.5
    aVertices(2).y = 0
    aVertices(2).z = 0.5
    aVertices(3).X = 0.5
    aVertices(3).y = 0
    aVertices(3).z = -0.5
    ' Copy floor vertices to ceiling vertices
    For i = 0 To 3
        ' Copy vertex
        aVertices(4 + i) = aVertices(i)
        ' Change height
        aVertices(4 + i).y = 1
    Next
    ' Fill faces array (number of vertices and vertex index for each face)
    ' Faces are described clockwise
    ' Floor
    aFaces(1) = 4 ' 4 vertices
    aFaces(2) = 0
    aFaces(3) = 1
    aFaces(4) = 2
    aFaces(5) = 3
    ' Ceiling
    aFaces(6) = 4
    aFaces(7) = 7
    aFaces(8) = 6
    aFaces(9) = 5
    aFaces(10) = 4
    ' Front wall
    aFaces(11) = 4
    aFaces(12) = 1
    aFaces(13) = 5
    aFaces(14) = 6
    aFaces(15) = 2
    ' Left wall
    aFaces(16) = 4
    aFaces(17) = 0
    aFaces(18) = 4
    aFaces(19) = 5
    aFaces(20) = 1
    ' Right wall
    aFaces(21) = 4
    aFaces(22) = 2
    aFaces(23) = 6
    aFaces(24) = 7
    aFaces(25) = 3
    ' Back wall
    aFaces(26) = 4
    aFaces(27) = 3
    aFaces(28) = 7
    aFaces(29) = 4
    aFaces(30) = 0
    ' Terminator
    aFaces(31) = 0
    ' Create and return object
    dixuD3DRM.CreateMeshBuilder MeshBuilder
    MeshBuilder.AddFaces 8, aVertices(0), 0, aNormals(0), aFaces(1), Nothing
    Set dixuCreateRoom = MeshBuilder
End Function

' DirectSound support
' Creates a DirectSoundBuffer from a wave file
Public Function dixuCreateSoundBuffer(ByVal Flags As Long, ByVal strFile As String, ByVal ds As DirectSound) As DirectSoundBuffer
    Dim hWave As Long
    Dim pcmwave As WAVEFORMATEX
    Dim lngSize As Long
    Dim lngPosition As Long
    Dim ptr1 As Long, ptr2 As Long, lng1 As Long, lng2 As Long
    Dim aByte() As Byte
    Dim dsb As DirectSoundBuffer
    ' Byte array to load the whole file
    ReDim aByte(1 To FileLen(strFile))
    hWave = FreeFile
    Open strFile For Binary As hWave
    ' Load the whole file in the byte array
    Get hWave, , aByte
    Close hWave
    ' Search "fmt" tag
    lngPosition = 1
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) <> "fmt"
        lngPosition = lngPosition + 1
    Wend
    ' Copy wave header to structure
    CopyMemory ByVal VarPtr(pcmwave), ByVal VarPtr(aByte(lngPosition + 8)), Len(pcmwave)
    ' Search "data" tag
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) + Chr$(aByte(lngPosition + 3)) <> "data"
        lngPosition = lngPosition + 1
    Wend
    ' Get the data size
    CopyMemory ByVal VarPtr(lngSize), ByVal VarPtr(aByte(lngPosition + 4)), Len(lngSize)
    ' Fill buffer description
    Dim dsbd As DSBUFFERDESC
    With dsbd
        .dwSize = Len(dsbd)
        '.dwFlags = DSBCAPS_CTRLDEFAULT 'Or DSBCAPS_STATIC Or DSBCAPS_LOCSOFTWARE
        .dwFlags = IIf(Flags = 0, DSBCAPS_CTRLDEFAULT, Flags)
        .dwBufferBytes = lngSize
        .lpwfxFormat = VarPtr(pcmwave)
    End With
    ' Create the sound buffer
    ds.CreateSoundBuffer dsbd, dsb, Nothing
    ' Lock
    dsb.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&
    ' Copy data to buffer
    CopyMemory ByVal ptr1, ByVal VarPtr(aByte(lngPosition + 4 + 4)), lng1
    ' Copy second part if needed
    If lng2 <> 0 Then
        CopyMemory ByVal ptr2, ByVal VarPtr(aByte(lngPosition + 4 + 4 + lng1)), lng2
    End If
    ' Works only with DirectX3.tlb or DirectX5.tlb...
    dsb.Unlock ptr1, lng1, ptr2, lng2
    Set dixuCreateSoundBuffer = dsb
End Function

' ********** Sprites support **********
' (don't compile if not needed)

#If DIXU_NOSPRITE = 0 Then
Private Sub dixuSpritesDraw()
    Dim Sprite As dixuSprite
    dixuLastTime = dixuTime
    dixuTime = Timer
    ' If runs at midnight !
    While dixuTime < dixuLastTime
        dixuTime = dixuTime + 86400 ' 1 day
    Wend
    
    ' Blt the score and the number of balls left and the sound mode.
    ScoreBlt Score, xScore, yScore
    ScoreBlt strSound, 112, yScore
    
    ' Move and paint all sprites
    For Each Sprite In dixuSprites
        Sprite.Move
        Sprite.Paint
    Next
    
    If Sgn(NumBalls) <> -1 Then
        ScoreBlt Chr(SC_SPACE) & NumBalls, 80, yScore
    End If
End Sub

Public Sub dixuSpriteRemove(Sprite As dixuSprite)
    dixuSprites.Remove Sprite.Key
    Set Sprite = Nothing
End Sub

#End If

' Windows messages support (not used yet, for an upcoming release...)
Private Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim D3DRMWinDEvice As Direct3DRMWinDevice
    Dim r As RECT
    Dim ps As PAINTSTRUCT
    Dim hdc As Long
    Select Case Msg
        Case WM_ACTIVATE
            If Not dixuD3DRMDevice Is Nothing Then
                Set D3DRMWinDEvice = dixuD3DRMDevice
                D3DRMWinDEvice.HandleActivate wParam
                Set D3DRMWinDEvice = Nothing
            Else
                WindowProc = CallWindowProc(lpPrevWndProc, hwnd, Msg, wParam, lParam)
            End If
            WindowProc = 0
        Case WM_PAINT
            If Not dixuD3DRMDevice Is Nothing Then
                If GetUpdateRect(hwnd, r, 0&) Then
                    Set D3DRMWinDEvice = dixuD3DRMDevice
                    BeginPaint hwnd, ps
                    D3DRMWinDEvice.HandlePaint ps.hdc
                    EndPaint hwnd, ps
                    Set D3DRMWinDEvice = Nothing
                End If
                WindowProc = 0
            Else
                WindowProc = CallWindowProc(lpPrevWndProc, hwnd, Msg, wParam, lParam)
            End If
            WindowProc = 0
        Case Else
           WindowProc = CallWindowProc(lpPrevWndProc, hwnd, Msg, wParam, lParam)
    End Select
End Function

Private Sub dixuHandlePaint(ByVal hwnd As Long)
    Dim D3DRMWinDEvice As Direct3DRMWinDevice
    Dim r As RECT
    Dim ps As PAINTSTRUCT
    If Not dixuD3DRMDevice Is Nothing Then
        If GetUpdateRect(hwnd, r, 0) Then
            BeginPaint hwnd, ps
            Set D3DRMWinDEvice = dixuD3DRMDevice
            D3DRMWinDEvice.HandlePaint ps.hdc
            Set D3DRMWinDEvice = Nothing
            EndPaint hwnd, ps
        End If
    End If
End Sub
