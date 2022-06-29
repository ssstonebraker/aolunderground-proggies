VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Picmove sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   1035
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5745
   Height          =   1440
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   5745
   Top             =   1170
   Width           =   5865
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2400
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   600
      Width           =   360
   End
   Begin VB.Image imOff 
      Height          =   330
      Left            =   780
      Picture         =   "picmove.frx":0000
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imOn 
      Height          =   330
      Left            =   300
      Picture         =   "picmove.frx":018A
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Move the mouse cursor over the picture and watch it change!"
      Height          =   375
      Left            =   420
      TabIndex        =   0
      Top             =   60
      Width           =   4755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' Picmove sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' Shows how to change a picture as a mouse moves.

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim bOn As Boolean

Private Sub Form_Load()
    Picture1.Picture = imOff.Picture
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bOn Then
        bOn = True
        Picture1.Picture = imOn.Picture
        SetCapture Picture1.hwnd
    ElseIf X < 0 Or Y < 0 Or X > Picture1.Width Or Y > Picture1.Height Then
        bOn = False
        Picture1.Picture = imOff.Picture
        ReleaseCapture
    End If
End Sub
