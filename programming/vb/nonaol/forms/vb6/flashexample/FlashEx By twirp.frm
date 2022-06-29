VERSION 5.00
Begin VB.Form frmFlash 
   Caption         =   "Flasher Example by Twirp"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   Icon            =   "FlashEx By twirp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer tmrFlash 
      Left            =   3480
      Top             =   240
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Fast"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Flash"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Fast or slow:"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Property Let Rate(intPerSecond As Integer)
tmrFlash.Interval = 1000 / intPerSecond
End Property
Property Let Flash(blnState As Boolean)
tmrFlash.Enabled = blnState
End Property


Private Sub Command1_Click()
If txtSpeed.text = "Fast" Then
Call FlashFast
Else
If txtSpeed.text = "Slow" Then
Call FlashSlow
Else
MsgBox "Please enter a valid spee, Fast or Slow. Thank you", , "Thank you for shopping Twirp"
End If
End If
End Sub

Private Sub Command2_Click()
tmrFlash.Enabled = False
End Sub

Private Sub Form_Load()
tmrFlash.Enabled = False
End Sub

Private Sub tmrFlash_Timer()
Dim lngRtn As Long
lngRtn = FlashWindow(hwnd, CLng(True))
End Sub
