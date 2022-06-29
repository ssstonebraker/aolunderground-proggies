VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   6720
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   240
      Picture         =   "frmMain.frx":7F790
      ScaleHeight     =   5775
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   6000
      Width           =   6780
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "> exit <"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5760
      TabIndex        =   1
      Top             =   4320
      Width           =   570
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "dos - dos@hider.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "special thanks to chokehold for making this great image. visit his site at www.hider.com/chokehold."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":AA398
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'form to bitmap example - dos - 7/3/99
'email: dos@hider.com
'site:  www.hider.com/dos
'aim:   xdosx

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2

Private Const WM_MOVE = &HF012
Private Const WM_SYSCOMMAND = &H112

Private lngRegion As Long

Private Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hdc, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
End Function

Private Sub Form_Load()
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(picBox)
  lngRetr& = SetWindowRgn(Me.hWnd, lngRegion&, True)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(Me.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub lblExit_Click()
  End
End Sub
