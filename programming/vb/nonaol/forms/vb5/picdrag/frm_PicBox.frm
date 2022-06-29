VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5172
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGraphic 
      Height          =   372
      Left            =   120
      Picture         =   "frm_PicBox.frx":0000
      ScaleHeight     =   324
      ScaleWidth      =   444
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   $"frm_PicBox.frx":04BE
      Height          =   612
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lDragging As Boolean

'We want lDragging to be set to true when the MouseDown event
'occurs and we want it to be false when the MouseUp event occurs.
'When the form loads, we want to start by assuming that the dragging
'hasn't yet begun. So, we need to set this variable in three different
'handlers:

Private Sub Form_Load()

    lDragging = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lDragging = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lDragging Then
        Form1.PaintPicture picGraphic.Picture, X, Y, picGraphic.Width, _
            picGraphic.Height
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lDragging = False
    
End Sub
