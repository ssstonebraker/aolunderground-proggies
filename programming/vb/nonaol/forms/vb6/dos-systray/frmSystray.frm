VERSION 5.00
Begin VB.Form frmSystray 
   Caption         =   "systray example by dos - 2.8.99"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrayIt 
      Caption         =   "send to systray"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSysTray 
         Caption         =   "un tray it"
      End
   End
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vbTray As NOTIFYICONDATA

Private Sub TrayIt()
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = Me.hwnd
    vbTray.uId = vbNull
    vbTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    vbTray.ucallbackMessage = WM_MOUSEMOVE
    vbTray.hIcon = Me.Icon
    vbTray.szTip = Me.Caption & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    Me.Hide
End Sub

Private Sub UnTrayIt()
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = Me.hwnd
    vbTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub

Private Sub cmdTrayIt_Click()
    Call TrayIt
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lngMsg As Long
    Dim blnFlag As Boolean, lngResult As Long
    lngMsg = X / Screen.TwipsPerPixelX
    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
            Case WM_LBUTTONDBLCLICK
                Me.WindowState = 0
                Me.Show
            Case WM_RBUTTONUP
                lngResult = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu mnuFile
        End Select
        blnFlag = False
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Call TrayIt
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnTrayIt
End Sub

Private Sub mnuSysTray_Click()
    MsgBox "systray menu clicked"
    Call UnTrayIt
    Me.WindowState = 0
    Me.Show
End Sub
