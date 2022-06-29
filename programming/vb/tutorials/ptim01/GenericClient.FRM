VERSION 5.00
Begin VB.Form frmGenericClient 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Generic Client"
   ClientHeight    =   2025
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
   ScaleHeight     =   2025
   ScaleWidth      =   7755
   Begin VB.CommandButton btnReceive 
      Caption         =   "&Receive"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtSendLine 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Text            =   "25"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblSendLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   240
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
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmGenericClient"
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

Dim TransmitSocket As Long

Private Sub Form_Load()
  Status = StartWinSock()
End Sub

Private Sub btnConnect_Click()
  Dim ServerAddress As Long
  Dim ServerName As String
  Dim SocketResponse As String
  Dim RemotePort As Long
  
  
  Dim CRLF As String
  CRLF = Chr$(13) + Chr$(10)
  
  TransmitSocket = 0
  RemotePort = val(txtPort.Text)
  
  ServerName = txtServer.Text
  Status = GetIPAddress(ServerAddress, ServerName)
  Status = CreateSocket(TransmitSocket, 0)
  Status = ConnectSocket(TransmitSocket, ServerAddress, RemotePort)

  btnConnect.DEFAULT = False
  btnConnect.Enabled = False
  btnSend.DEFAULT = True
    
End Sub

Private Sub btnReceive_Click()
  Dim SocketResponse As String
  
  Status = ReceiveSocket(TransmitSocket, SocketResponse)
  txtSendLine.SetFocus

End Sub

Private Sub btnSend_Click()
  Dim SendLine As String
  
  SendLine = txtSendLine.Text
  Status = SendSocket(TransmitSocket, SendLine)
  
  txtSendLine.SelStart = 0
  txtSendLine.SelLength = Len(txtSendLine.Text)
  txtSendLine.SetFocus


End Sub

Private Sub btnQuit_Click()
    
  Unload Me
  End

End Sub

Private Sub Form_Unload(Cancel As Integer)

  If TransmistSocket > 0 Then
    Status = ReleaseSocket(TransmitSocket)
  End If
  
  Status = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(Status)

End Sub


