VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "OnTop sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   3390
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6240
   Height          =   3795
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   6240
   Top             =   1170
   Width           =   6360
   Begin VB.CommandButton Command2 
      Caption         =   "Remove from Top"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Put on Top"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"ontop.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' OnTop Sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This app shows you how to make a "floating" window which will stay on top
' even when not activated.  Note that you MUST compile this and run as an EXE
' file before it will stay on top of all other running programs.  That is because
' when it is run in the VB environment, it is considered a "child" window of VB,
' and will thus only float on top of other VB windows.

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Sub Command1_Click()
    Dim i
    i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub Command2_Click()
    Dim i
    i = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
