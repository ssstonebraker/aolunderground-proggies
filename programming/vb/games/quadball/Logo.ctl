VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl Logo 
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ScaleHeight     =   2235
   ScaleWidth      =   2820
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox MainPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   1950
      TabIndex        =   0
      Top             =   0
      Width           =   1950
   End
   Begin PicClip.PictureClip Animate 
      Left            =   225
      Top             =   225
      _ExtentX        =   18521
      _ExtentY        =   35719
      _Version        =   393216
      Rows            =   18
      Cols            =   7
   End
End
Attribute VB_Name = "Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' this is used for the spinning Quad-Ball Logo.
Dim Cell As Integer
Private Sub Timer1_Timer()
 MainPic.Picture = Animate.GraphicCell(Cell)
 DoEvents
 Cell = Cell + 1
 If Cell > Int((7 * 18) - 1) Then Cell = 0
End Sub
Private Sub UserControl_Initialize()
 ExitIt = False
 ThisDir
 Animate.Picture = LoadPicture("MadeBy.img")
 MainPic.Picture = Animate.GraphicCell(0)
 Resize
End Sub
' if logo spins too fast then increase the value of speed
Public Sub Start(Optional Speed As Integer = 50)
 Timer1.Interval = Speed
 Timer1.Enabled = True
End Sub
Public Sub StopAnimation()
 Timer1.Enabled = False
End Sub
' I used this sub so that the resize could be called outside of the control
Public Sub Resize()
 Width = MainPic.Width
 Height = MainPic.Height
End Sub
Private Sub UserControl_Resize()
 Resize
End Sub
