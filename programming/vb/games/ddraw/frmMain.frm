VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DirectDraw Animation-1"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' DirectDraw Stuff
' If you referenced the Patrice Scribe's
' TLB and you get an error, try changing the variable
' to DirectDrawxxxxx instead of IDirectDrawxxxx
Dim lpDD As IDirectDraw
Dim lpDDSFront As IDirectDrawSurface
Dim lpDDSBack  As IDirectDrawSurface
Dim lpDDSPic   As IDirectDrawSurface

' Some other vars
Dim bEnd As Boolean ' True  = App is ending
Dim x As Long, y As Long

' API Declarations
' Win32
Const IMAGE_BITMAP = 0
Const LR_LOADFROMFILE = &H10
Const LR_CREATEDIBSECTION = &H2000
Const SRCCOPY = &HCC0020

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' GDI32
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' USER32
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' Loads a bitmap in a DirectDraw surface
Public Function CreateDDSFromBitmap(dd As IDirectDraw, ByVal strFile As String) As IDirectDrawSurface
Dim hbm As Long ' Handle on bitmap
Dim bm As BITMAP ' Bitmap header
Dim ddsd As DDSURFACEDESC ' Surface description
Dim dds As IDirectDrawSurface ' Created surface
Dim hdcImage As Long ' Handle on image
Dim mhdc As Long ' Handle on surface context

' Load bitmap
hbm = LoadImage(ByVal 0&, strFile, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

' Get bitmap info
GetObject hbm, Len(bm), bm

' Fill surface description
With ddsd
.dwSize = Len(ddsd)
.dwFlags = DDSD_CAPS + DDSD_HEIGHT + DDSD_WIDTH
.DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN
.ddpfPixelFormat.dwSize = 8
.dwWidth = bm.bmWidth
.dwHeight = bm.bmHeight
End With

' Create surface
dd.CreateSurface ddsd, dds, Nothing

' Create memory device
hdcImage = CreateCompatibleDC(ByVal 0&)

' Select the bitmap in this memory device
SelectObject hdcImage, hbm

' Restore the surface
dds.Restore

' Get the surface's DC
dds.GetDC mhdc

' Copy from the memory device to the DirectDrawSurface
StretchBlt mhdc, 0, 0, ddsd.dwWidth, ddsd.dwHeight, hdcImage, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY

' Release the surface's DC
dds.ReleaseDC mhdc

' Release the memory device and the bitmap
DeleteDC hdcImage
DeleteObject hbm

' Returns the new surface
Set CreateDDSFromBitmap = dds
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:
            bEnd = True
        Case vbKeyLeft:  ' Move the fire object left/right respectively
            x = x - 5
            If x < 0 Then x = 0
        Case vbKeyRight:
            x = x + 5
            If x + 84 > 640 Then x = 640 - 84
        Case vbKeyUp:
            y = y - 5
            If y < 0 Then y = 0
        Case vbKeyDown:
            y = y + 5
            If y + 88 > 480 Then y = 480 - 88
    End Select
End Sub

Private Sub Form_Load()
Dim ddsd As DDSURFACEDESC
Dim ddc  As DDSCAPS
Dim ddck As DDCOLORKEY
Dim rc   As RECT
Dim i    As Long

    ' Initialize DirectDraw
    Call DirectDrawCreate(ByVal 0&, lpDD, Nothing)
    Call lpDD.SetCooperativeLevel(Me.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT)
    Call lpDD.SetDisplayMode(640, 480, 8)
    
    ' Create a front and a bitmap surfaces
    With ddsd
        .dwSize = Len(ddsd)
        .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
        .dwBackBufferCount = 1
    End With
    Call lpDD.CreateSurface(ddsd, lpDDSFront, Nothing)
    
    ' Retrieve the back buffer
    With ddc
        .dwCaps = DDSCAPS_BACKBUFFER
    End With
    Call lpDDSFront.GetAttachedSurface(ddc, lpDDSBack)
    
    ' Create a surface with the bitmap
    Set lpDDSPic = CreateDDSFromBitmap(lpDD, App.Path & "\Fir0.bmp")
    
    ' Set the color key
    With ddck
        .dwColorSpaceHighValue = RGB(0, 0, 0)
        .dwColorSpaceLowValue = .dwColorSpaceHighValue
    End With
    Call lpDDSPic.SetColorKey(DDCKEY_SRCBLT, ddck)
    
    Me.Show
    
    ' Now, start blitting until the end
    While Not bEnd
        DoEvents ' be nice
        
        ' Clear the front buffer
        Call ClearBuffer(lpDDSBack)
        
        ' See what we need to blt
        If i = 0 Then rc.Left = 1
        If i = 1 Then rc.Left = 86
        If i = 2 Then rc.Left = 171
        If i = 3 Then rc.Left = 1
        If i = 4 Then rc.Left = 86
        If i = 5 Then rc.Left = 171
        
        rc.Right = rc.Left + 84
        
        If i > 2 Then
            rc.Top = 90
        Else
            rc.Top = 1
        End If
        
        rc.Bottom = rc.Top + 88
        
        ' Copy the needed region
        Call lpDDSBack.BltFast(x, y, lpDDSPic, rc, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        
        ' Wait for vertical blank to end
        Call WaitForVerticalBlank
        
        ' Copy the work area surface onto the front buffer
        Call lpDDSFront.Flip(Nothing, DDFLIP_WAIT)
        
        ' Synchronize to reasonable speed
        Call Sleep(80)
        
        ' Increment and see the i value
        i = i + 1
        If i > 5 Then i = 0
    Wend
    
    Unload Me
End Sub

' Clears the buffer
Sub ClearBuffer(ByRef lpDDS As IDirectDrawSurface)
Dim fx As DDBLTFX

    ' Fill out the blt operation description
    With fx
        .dwSize = Len(fx)
        .dwFillColor = RGB(0, 0, 0)
    End With
    
    ' Color fill the surface
    Call lpDDS.Blt(ByVal 0&, Nothing, ByVal 0&, DDBLT_WAIT Or DDBLT_COLORFILL, fx)
End Sub

' Copies a whole buffer onto another
Sub CopyBuffer(ByRef lpDDSSrc As IDirectDrawSurface, ByRef lpDDSDest As IDirectDrawSurface)
Dim ddsd As DDSURFACEDESC
Dim rc   As RECT

    ' Get the surface desc for the source surface
    With ddsd
        .dwSize = Len(ddsd)
        .dwFlags = DDSD_WIDTH Or DDSD_HEIGHT
    End With
    Call lpDDSSrc.GetSurfaceDesc(ddsd)
    
    ' Now, copy the whole source onto the dest buffer
    rc.Left = 0
    rc.Top = 0
    rc.Right = ddsd.dwWidth
    rc.Bottom = ddsd.dwHeight
    
    ' BltFast the surface
    Call lpDDSDest.BltFast(0, 0, lpDDSSrc, rc, DDBLTFAST_WAIT) ' Set this flag if that surface has a source key Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub WaitForVerticalBlank()
    ' This waits for the veritcal blank to end
    Call lpDD.WaitForVerticalBlank(DDWAITVB_BLOCKEND, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call lpDD.RestoreDisplayMode
    Set lpDDSPic = Nothing
    Set lpDDSBack = Nothing
    Set lpDDSFront = Nothing
    Set lpDD = Nothing
End Sub
