VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   2760
   End
   Begin VB.PictureBox picBackward 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   1200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picForward 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   120
      Picture         =   "Form1.frx":1F9A
      ScaleHeight     =   1050
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'number of times to redraw image per rotation
'in other words 360/NUM_TURNS = # of degrees per each turn of the image
Dim dAngle As Double

Const NUM_TURNS = 36
Const PI = 3.14159265358979
Const CENTER_X = 4000

Const SRCCOPY = &HCC0020

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Sub Form_Load()
picForward.AutoSize = True
picBackward.AutoSize = True
picBuffer.Width = picForward.Width
picBuffer.Height = picForward.Height
picDisplay.Width = picForward.Width
picDisplay.Height = picForward.Height
picBuffer.Visible = False
picForward.Visible = False
picBackward.Visible = False
picBuffer.AutoRedraw = True
picForward.AutoRedraw = True
picBackward.AutoRedraw = True
picDisplay.AutoRedraw = True
picForward.BorderStyle = 0
picBackward.BorderStyle = 0
picBuffer.BorderStyle = 0
picDisplay.BorderStyle = 0
End Sub
Private Sub Timer1_Timer()
'assume that 0 degrees is when the picture is facing forward

picBuffer.Cls

If Cos(dAngle * PI / 180) >= 0 Then
    Call StretchBlt(picBuffer.hdc, (picForward.Width - Abs(Cos(dAngle * PI / 180) * picForward.Width)) / (2 * Screen.TwipsPerPixelX), 0, Abs(Cos(dAngle * PI / 180) * picForward.Width) / Screen.TwipsPerPixelX, picForward.Height / Screen.TwipsPerPixelY, picForward.hdc, 0, 0, picForward.Width / Screen.TwipsPerPixelX, picForward.Height / Screen.TwipsPerPixelY, SRCCOPY)
ElseIf Cos(dAngle * PI / 180) < 0 Then
    Call StretchBlt(picBuffer.hdc, (picBackward.Width - Abs(Cos(dAngle * PI / 180) * picBackward.Width)) / (2 * Screen.TwipsPerPixelX), 0, Abs(Cos(dAngle * PI / 180) * picBackward.Width) / Screen.TwipsPerPixelX, picBackward.Height / Screen.TwipsPerPixelY, picBackward.hdc, 0, 0, picBackward.Width / Screen.TwipsPerPixelX, picBackward.Height / Screen.TwipsPerPixelY, SRCCOPY)
End If

Call BitBlt(picDisplay.hdc, 0, 0, picBuffer.Width / Screen.TwipsPerPixelX, picBuffer.Height / Screen.TwipsPerPixelY, picBuffer.hdc, 0, 0, SRCCOPY)
picDisplay.Refresh

'increment angle and make sure it stays between 0 and 360
dAngle = dAngle + 360 / NUM_TURNS
dAngle = dAngle Mod 360
End Sub
