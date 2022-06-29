VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Minimize"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Glassify"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UnGlassify"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
End Sub

Private Sub Command1_Click()
UnglassifyForm Me
End Sub


Private Sub Command2_Click()
GlassifyForm Me
End Sub

Private Sub Command3_Click()
WindowState = vbMinimized
End Sub

Private Sub Command4_Click()
Unload Me
End Sub


Private Sub Label1_Click()
Unload Me

End Sub
