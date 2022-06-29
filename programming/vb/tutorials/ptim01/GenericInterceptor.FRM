VERSION 5.00
Begin VB.Form frmGenericInterceptor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Generic Interceptor"
   ClientHeight    =   2415
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   7755
   Begin VB.CheckBox chkBinary 
      Caption         =   "&Binary Protocol"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtLogFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Text            =   "C:\Temp\t.txt"
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtListenClient 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtReceiveServer 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox txtReceiveClient 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "143"
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton btnListen 
      Caption         =   "&Listen"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Text            =   "143"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblLogFile 
      Caption         =   "Log File:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLocalPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblRemotePort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmGenericInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' Copyright 1999 - Digital Press, John Rhoton
'
' This program has been written to illustrate the Internet Mail protocols.
' It is provided free of charge and unconditionally.  However, it is not
' intended for production use, and therefore without warranty or any
' implication of support.
'
' You can find an explanation of the concepts behind this code in
' the book:  Programmer's Guide to Internet Mail by John Rhoton,
' Digital Press 1999.  ISBN: 1-55558-212-5.
'
' For ordering information please see http://www.amazon.com or
' you can order directly with http://www.bh.com/digitalpress.
'
'========================================================================

Option Explicit

Const WM_LBUTTONDOWN = &H201
Dim TransmitServerSocket As Long
Dim TransmitClientSocket As Long
Dim ListenClientSocket As Long
Dim Status As Integer

Private Sub Form_Load()
  Status = StartWinSock()
End Sub

Private Sub btnListen_Click()
  Dim ListenPort As Long
  
  ListenPort = val(txtLocalPort.Text)
  
  Status = CreateSocket(ListenClientSocket, ListenPort)
  
  frmVBWinSockDialog.Left = Me.Left
  frmVBWinSockDialog.Top = Me.Height
  
  If listen(ListenClientSocket, 5) Then
    MsgBox "Could not listen"
    End
  End If
  
  If WSAAsyncSelect(ListenClientSocket, txtListenClient.hwnd, WM_LBUTTONDOWN, FD_ACCEPT) Then
    MsgBox "Unable to set Asynch mode"
  End If

  Open txtLogFile.Text For Output As #1

End Sub

Private Sub txtListenClient_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim TransmitClientSocketAddress As OutputSocketDescriptor
  Dim TransmitClientSocketAddressSize As Integer
  Dim ServerAddress As Long
  Dim ServerName As String
  Dim SocketResponse As String
  Dim ConnectPort As Long
  
  TransmitClientSocketAddressSize = 20
  TransmitClientSocket = accept(ListenClientSocket, TransmitClientSocketAddress, TransmitClientSocketAddressSize)

  TransmitServerSocket = 0
  ConnectPort = val(txtRemotePort.Text)
  
  ServerName = txtServer.Text
  Status = GetIPAddress(ServerAddress, ServerName)
  Status = CreateSocket(TransmitServerSocket, 0)
  Status = ConnectSocket(TransmitServerSocket, ServerAddress, ConnectPort)

  If WSAAsyncSelect(TransmitClientSocket, txtReceiveClient.hwnd, WM_LBUTTONDOWN, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If
  
  If WSAAsyncSelect(TransmitServerSocket, txtReceiveServer.hwnd, WM_LBUTTONDOWN, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If
    
End Sub

Private Sub txtReceiveServer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim SocketResponse As String
    
  If chkBinary.Value = 0 Then
    Status = ReceiveSocket(TransmitServerSocket, SocketResponse, -1)
    Status = SendSocket(TransmitClientSocket, SocketResponse, False)
  Else
    Status = ReceiveSocketBinary(TransmitServerSocket, SocketResponse)
    Status = SendSocketBinary(TransmitClientSocket, SocketResponse)
  End If
  PrintLogFile "[-->" + Trim$(Str$(TransmitServerSocket)) + "] ", SocketResponse

End Sub

Private Sub txtReceiveClient_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim SocketResponse As String

  If chkBinary.Value = 0 Then
    Status = ReceiveSocket(TransmitClientSocket, SocketResponse, -1)
    Status = SendSocket(TransmitServerSocket, SocketResponse, False)
  Else
    Status = ReceiveSocketBinary(TransmitClientSocket, SocketResponse)
    Status = SendSocketBinary(TransmitServerSocket, SocketResponse)
  End If
  
  PrintLogFile "[-->" + Trim$(Str$(TransmitClientSocket)) + "] ", SocketResponse

End Sub

Private Sub txtLocalPort_Change()
    txtRemotePort.Text = txtLocalPort.Text
End Sub

Private Sub PrintLogFile(ByVal PrintPrefix As String, ByVal PrintText As String)
  Dim ScanPosition As Integer
  
  If chkBinary.Value = 0 Then
    Print #1, PrintPrefix + Chr$(13) + Chr$(10); PrintText
  Else
    Print #1, PrintPrefix + Chr$(13) + Chr$(10);
    For ScanPosition = 1 To Len(PrintText)
      Print #1, Right$(Space$(4) + Str$(Asc(Mid$(PrintText, ScanPosition, 1))), 4);
      If ScanPosition / 20 = Int(ScanPosition / 20) Then
        Print #1, ""
      End If
    Next
    Print #1, ""
    Print #1, ""
    Print #1, ""
  End If
    
End Sub

Private Sub btnQuit_Click()
    
  Unload Me
  End

End Sub

Private Sub Form_Unload(CANCEL As Integer)

  If TransmitServerSocket <> 0 Then
    Status = ReleaseSocket(TransmitServerSocket)
  End If
  
  If TransmitClientSocket <> 0 Then
    Status = ReleaseSocket(TransmitClientSocket)
  End If
  
  If ListenClientSocket <> 0 Then
    Status = ReleaseSocket(ListenClientSocket)
  End If
  
  Status = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(Status)
  Close #1

End Sub
