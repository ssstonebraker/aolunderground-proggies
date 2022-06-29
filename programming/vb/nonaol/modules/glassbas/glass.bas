Attribute VB_Name = "glassBAS"
' Glass.bas
' By Tracy Blanton
' AKA "Pikode"
' Id Like To Thank TokyoMagnm@aol.com For This Great Code
' And Thanx TO KNK 2000 (www.knk2000.com/knk) Cause That Is Where I Found The Form With This Great Codeing In It
'
'
'
'How To Use Glass.bas
'
' TO Make The From Transparent Enter "GlassifyForm Me"
' To Make Non Transparent Enter "UnglassifyForm Me"
'
'
'
' Now Start Programming
'
'
'
' Note: Do NOT Edit This Code In Any Way Or Your Form Will Not Work

Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Sub GlassifyForm(frm As Form)

Const RGN_DIFF = 4
Const RGN_OR = 2

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

    If WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - ScaleWidth) / 2
    title_height = hgt - border_width - ScaleHeight
    inner_rgn = CreateRectRgn( _
        border_width, _
        title_height, _
        wid - border_width, _
        hgt - border_width)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF

    ' Create the control regions.
    For Each ctl In Controls
        If ctl.Container Is frm Then
            ctl_left = ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = ScaleX(ctl.Top, frm.ScaleMode, vbPixels) _
                + title_height
            ctl_right = ScaleX(ctl.Width, frm.ScaleMode, vbPixels) _
                + ctl_left
            ctl_bottom = ScaleX(ctl.Height, frm.ScaleMode, vbPixels) _
                + ctl_top
            control_rgn = CreateRectRgn( _
                ctl_left, ctl_top, _
                ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, _
                control_rgn, RGN_OR
        End If
    Next ctl

    ' Restrict the window to the region.
    SetWindowRgn hWnd, combined_rgn, True
End Sub

Private Sub UnglassifyForm(frm As Form)
Const RGN_DIFF = 4
Const RGN_OR = 2

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

    If WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - ScaleWidth) / 2
    title_height = hgt - border_width - ScaleHeight
    inner_rgn = CreateRectRgn( _
        border_width, _
        title_height, _
        wid - border_width, _
        hgt - border_width)

    

    ' Create the control regions.
    For Each ctl In Controls
        If ctl.Container Is frm Then
            ctl_left = ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = ScaleX(ctl.Top, frm.ScaleMode, vbPixels) _
                + title_height
            ctl_right = ScaleX(ctl.Width, frm.ScaleMode, vbPixels) _
                + ctl_left
            ctl_bottom = ScaleX(ctl.Height, frm.ScaleMode, vbPixels) _
                + ctl_top
            control_rgn = CreateRectRgn( _
                ctl_left, ctl_top, _
                ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, _
                control_rgn, RGN_OR
        End If
    Next ctl

    ' Restrict the window to the region.
    SetWindowRgn hWnd, combined_rgn, True

