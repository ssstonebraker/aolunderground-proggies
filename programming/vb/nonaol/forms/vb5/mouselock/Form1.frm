VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mouse Trapper By DeViL"
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   2370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   255
      Left            =   -120
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   255
      Left            =   -120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
'need this crap for it to work
Private Declare Function ClipCursor Lib "user32" _
(lpRect As Any) As Long

Private Sub DisableTrap(CurForm As Form)

    Dim erg As Long
    Dim NewRect As RECT
    CurForm.Caption = "Mouse released" 'makes the forms caption Mouse Released take this off if u want

    With NewRect
    .Left = 0&
    .Top = 0&
    .Right = Screen.Width / Screen.TwipsPerPixelX
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
    End With

    erg& = ClipCursor(NewRect)
    End Sub
Private Sub EnableTrap(CurForm As Form)

    Dim x As Long, y As Long, erg As Long
    Dim NewRect As RECT
    x& = Screen.TwipsPerPixelX
    y& = Screen.TwipsPerPixelY
    CurForm.Caption = "Mouse trapped" 'makes the forms caption Mouse Trapped take this off if u want

    With NewRect
    .Left = CurForm.Left / x&
    .Top = CurForm.Top / y&
    .Right = .Left + CurForm.Width / x&
    .Bottom = .Top + CurForm.Height / y&
    End With

    erg& = ClipCursor(NewRect)
    End Sub
Private Sub Command1_Click()
DisableTrap Me 'takes away the trap
End Sub

Private Sub Command2_Click()
'this starts the trap
EnableTrap Me
End Sub

'This show how to keep a mouse onto a form and cannot
'off!
Private Sub Form_Unload(Cancel As Integer)
DisableTrap Me 'do this or it will stay trapped un less
'they do alt+ctrl+del
End Sub
