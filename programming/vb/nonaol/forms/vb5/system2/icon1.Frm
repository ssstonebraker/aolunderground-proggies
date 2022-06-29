VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2460
   Height          =   1860
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   2460
   Top             =   1170
   Width           =   2580
   Begin VB.CommandButton Command3 
      Caption         =   "Minimize"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "Example icon form"
      Top             =   840
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   1680
      Picture         =   "icon1.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim picture As Long
Dim message As Long
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = Picture1.hwnd
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = Picture1.picture
Tic.szTip = text1 & Chr$(0)
erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Private Sub Command2_Click()
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = Picture1.hwnd
Tic.uID = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub


Private Sub Command3_Click()
Me.WindowState = 1
Me.Hide
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / Screen.TwipsPerPixelX
Select Case X
Case WM_RBUTTONDOWN
Caption = "right"
Me.Show
Me.WindowState = 0

Case WM_MOUSEMOVE
Caption = "mouse"
End Select
End Sub


