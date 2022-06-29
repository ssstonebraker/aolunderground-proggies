VERSION 5.00
Begin VB.Form ZoomForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoom In Example by PAT or JK"
   ClientHeight    =   2685
   ClientLeft      =   2025
   ClientTop       =   1980
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   2880
      ScaleHeight     =   1995
      ScaleWidth      =   2610
      TabIndex        =   1
      Top             =   120
      Width           =   2640
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   120
      Picture         =   "ZoomForm.frx":0000
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   0
      Top             =   120
      Width           =   2640
      Begin VB.Shape Shape1 
         Height          =   600
         Left            =   1140
         Top             =   420
         Width           =   780
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Just drag the box around on the picture to see a zoomed in picture."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2220
      Width           =   5355
   End
End
Attribute VB_Name = "ZoomForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Zoom in example by PAT or JK
'email:patorjk@aol.com, webpage: www.patorjk.com
'This is an example on how to zoom in on a picture

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Private Sub Form_Load()
Me.ScaleMode = 3 ' Make the scalemode pixels
Picture1.ScaleMode = 3
Picture2.ScaleMode = 3
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ' make sure the mouse button is down
Shape1.Left = X - (Shape1.Width / 2)
Shape1.Top = Y - (Shape1.Height / 2) ' grab the center of the box
DoEvents
' take a snap shot of the boxed area and put it in picture2
Call StretchBlt(Picture2.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hdc, (Shape1.Left) + 1, (Shape1.Top) + 1, (Shape1.Width) - 2, (Shape1.Height) - 2, SRCCOPY)
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
Call StretchBlt(Picture2.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hdc, (Shape1.Left) + 1, (Shape1.Top) + 1, (Shape1.Width) - 2, (Shape1.Height) - 2, SRCCOPY)
End Sub
