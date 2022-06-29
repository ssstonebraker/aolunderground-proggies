VERSION 5.00
Begin VB.Form frmGenericAsynchronousServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Generic Asynchronous Server"
   ClientHeight    =   5100
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   7470
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
   ScaleHeight     =   5100
   ScaleWidth      =   7470
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton btnListen 
      Caption         =   "&Listen"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtDialog 
      Height          =   3735
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   6975
   End
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtRequest 
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   3
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtSocket 
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblSocket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Socket:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmGenericAsynchronousServer"
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

Dim ListenSocket As Long
Dim TransmitSocket As Long
Dim Status As Integer

Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Dim CurrentConnection As Integer
Dim CRLF As String

Dim DialogLength(20) As Integer

Private Sub Form_Load()

  Status = StartWinSock()
  
  CRLF = Chr(13) + Chr(10)
  
  cboConnection.AddItem "0"
    
  txtRequest(0).Text = ""
  txtSocket(0).Text = ""
  txtDialog(0).Text = ""
  
End Sub

  ' Listen on specified port
Private Sub btnListen_Click()
  Dim ListenSocketAddress As InputSocketDescriptor
  Dim TCPPort As Integer
  
  TCPPort = val(txtPort.Text)
  
  Status = CreateSocket(ListenSocket, TCPPort)
  
  If listen(ListenSocket, 5) Then
      MsgBox "Could not listen on Port " + Str$(TCPPort)
      End
  End If
  
  If WSAAsyncSelect(ListenSocket, txtConnect.hwnd, WM_MBUTTONUP, FD_ACCEPT) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

  ' Accept incoming connection
Private Sub txtConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim TransmitSocketAddress As OutputSocketDescriptor
  Dim TransmitSocketAddressSize As Integer
  Dim CurrentConnection As Integer
  
  TransmitSocketAddressSize = LenB(TransmitSocketAddress)
  TransmitSocketAddressSize = 20
  
  TransmitSocket = accept(ListenSocket, TransmitSocketAddress, TransmitSocketAddressSize)

  CurrentConnection = cboConnection.ListCount
    
  cboConnection.AddItem Trim(Str(CurrentConnection))
  Load txtSocket(CurrentConnection)
  Load txtRequest(CurrentConnection)
  Load txtDialog(CurrentConnection)
  
  txtSocket(CurrentConnection) = Trim(Str(TransmitSocket))
  
  DisplayConnection CurrentConnection
  
  If WSAAsyncSelect(TransmitSocket, txtRequest(CurrentConnection).hwnd, WM_MBUTTONDBLCLK, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If
    
End Sub

  ' Process Client Request
Private Sub txtRequest_DblClick(Index As Integer)
  Dim SendLine As String
  Dim ClientRequest As String
  
  TransmitSocket = val(txtSocket(Index).Text)
  Status = ReceiveSocket(TransmitSocket, ClientRequest)
  DisplayConnection Index
  txtDialog(Index).Text = txtDialog(Index).Text + ClientRequest + CRLF
  DialogLength(Index) = Len(txtDialog(Index).Text)
  
  txtDialog(Index).SelStart = DialogLength(Index)
End Sub

Private Sub txtDialog_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim LastLine As String

  If KeyCode = 13 Then
    TransmitSocket = val(txtSocket(Index).Text)
  
    LastLine = Mid(txtDialog(Index).Text, DialogLength(Index) + 1)
    Status = SendSocket(TransmitSocket, LastLine)
    DialogLength(Index) = Len(txtDialog(Index).Text)
  End If

End Sub

Private Sub cboConnection_Click()
  DisplayConnection cboConnection.ListIndex
End Sub
  
Private Sub DisplayConnection(NewConnection As Integer)

  If NewConnection < 0 Or NewConnection >= cboConnection.ListCount Then
    NewConnection = 0
  End If
  
  If NewConnection = CurrentConnection Then Exit Sub
  
  txtSocket(CurrentConnection).Visible = False
  txtSocket(NewConnection).Visible = True
  txtDialog(CurrentConnection).Visible = False
  txtDialog(NewConnection).Visible = True
      
  cboConnection.ListIndex = NewConnection
  CurrentConnection = NewConnection
  
End Sub

Private Sub btnQuit_Click()
    
  Status = ReleaseSocket(ListenSocket)
  Status = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(Status)

  End

End Sub


Private Sub Form_Unload(CANCEL As Integer)
   
  If ListenSocket > 0 Then
    Status = ReleaseSocket(ListenSocket)
    Status = WSACleanup()
    Debug.Print "WSACleanup status " & SocketError(Status)
  End If

End Sub


