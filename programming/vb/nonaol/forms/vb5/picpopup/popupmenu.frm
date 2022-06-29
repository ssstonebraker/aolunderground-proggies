VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "aim virii's Popup menu example"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   Icon            =   "popupmenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Caption         =   "Right Click"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Left Click"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2880
      Picture         =   "popupmenu.frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "popupmenu.frx":0614
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu FILE 
      Caption         =   "&File"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ABOUT 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
      Begin VB.Menu SPACE 
         Caption         =   "-"
      End
      Begin VB.Menu EXIT 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        PopupMenu FILE
    End If

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        PopupMenu FILE
    End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 3 Then
        PopupMenu FILE
    End If
End Sub
