VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl EarthBlow1 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   5310
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   2220
      TabIndex        =   0
      Top             =   0
      Width           =   2220
   End
   Begin PicClip.PictureClip PClip 
      Left            =   810
      Top             =   1080
      _ExtentX        =   18891
      _ExtentY        =   7303
      _Version        =   393216
      Rows            =   3
      Cols            =   7
   End
End
Attribute VB_Name = "EarthBlow1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Control is Used to Show the Earth       '
' Dissolving If it Is Not Saved By The Player  '
'______________________________________________'
Dim ExitLoop As Boolean, T, TT
'Sets The Control To a Round Shape
Sub SetShape(Optional Round As Boolean = True, Optional Square As Boolean = False)
If Round = True Then
 SetWindowRgn UserControl.hWnd, CreateEllipticRgn(16, 6, (UserControl.Width / Screen.TwipsPerPixelX) - 13, (UserControl.Height / Screen.TwipsPerPixelY) - 8), True
ElseIf Square = True Then
 SetWindowRgn UserControl.hWnd, CreateRectRgn(0, 0, (UserControl.Width / Screen.TwipsPerPixelX), (UserControl.Height / Screen.TwipsPerPixelY)), True
End If
End Sub
Sub UserControl_Initialize()
 ' load the animation picture
 ThisDir
 PClip.Picture = LoadPicture("EarthBlow.img")
 If PClip.Picture = 0 Then LoadPicture (App.Path & "\" & "EarthBlow.img")
 Picture1.AutoSize = False
 Picture1.Picture = PClip.GraphicCell(0)
 Picture1.AutoSize = True
End Sub

'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' Calling "Animate" By Itself Will Display The Automated         '
' Animation, But The Others Enable Slowing Down of the dissolve, '
' Times To Loop The Animation and To Show Only Certain Frames    '
'________________________________________________________________'
Public Sub Animate(Optional Delay = 0, Optional TimesToLoop As Integer = 1, Optional Frame As Integer = 21)
Pause Delay ' if the user wants to wait a few secs before the animation starts
If Frame < 21 Then
 Picture1.Picture = PClip.GraphicCell(Frame)
 Picture1.Refresh
 Pause 0.1
 Exit Sub
End If
Dim i, II
For II = 1 To TimesToLoop
 For i = 0 To 20 ' go through each frame and display it
  Picture1.Picture = PClip.GraphicCell(i)
  Pause 0.1
  If ExitLoop = True Then Exit Sub
 Next i
Next II
End Sub
Private Sub UserControl_Resize()
 UserControl.Width = Picture1.Width
 UserControl.Height = Picture1.Height
End Sub
Sub Pause(T)
 TT = Timer
 Do
  Picture1.Refresh
 Loop Until Timer > T + TT
End Sub
Private Sub UserControl_Terminate()
 ExitLoop = True
End Sub
