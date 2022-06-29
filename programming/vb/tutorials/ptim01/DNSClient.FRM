VERSION 5.00
Begin VB.Form frmDNSClient 
   AutoRedraw      =   -1  'True
   Caption         =   "DNS Client"
   ClientHeight    =   4395
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   7065
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
   ScaleHeight     =   4395
   ScaleWidth      =   7065
   Begin VB.CommandButton btnLookup 
      Caption         =   "&Lookup"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox lstResourceRecords 
      Height          =   1815
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   5775
   End
   Begin VB.CommandButton btnQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblResourceRecords 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Resource Records::"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblDomain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Domain:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
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
Attribute VB_Name = "frmDNSClient"
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
Dim lngTransmitSocket As Long
Const DNSPort = 53
Dim lngStatus As Long

Private Sub Form_Load()

    lngStatus = StartWinSock()

End Sub

Private Sub btnLookup_Click()
    
  Dim lngServerAddress As Long
  Dim strServerName As String
  
  Dim strSocketResponse As String
  Dim strSendLine As String
  Dim bytLineLength As Byte
  Dim dmsgSendMessage As New DNSMessage
    
  Dim x As Integer
  Dim dmsgResponseMessage As DNSMessage
    
  lngTransmitSocket = 0
  
  strServerName = txtServer.Text
  lngStatus = GetIPAddress(lngServerAddress, strServerName)
  lngStatus = CreateSocket(lngTransmitSocket, 0)
  lngStatus = ConnectSocket(lngTransmitSocket, lngServerAddress, DNSPort)

  dmsgSendMessage.ComposeQuestion 15, txtDomain.Text
  strSendLine = dmsgSendMessage.TransferString
  
  bytLineLength = Len(strSendLine)
  
  strSendLine = Chr$(0) + Chr$(bytLineLength) + strSendLine + Chr$(0) + Chr$(0)
    
  Debug.Print strSendLine
  
  For x = 1 To Len(strSendLine) ' Len(strSocketResponse)
    Debug.Print Asc(Mid$(strSendLine, x, 1));
    If Int(x / 4) = x / 4 Then Debug.Print
  Next x
  Debug.Print
  Debug.Print "-"
  
  lngStatus = SendSocketBinary(lngTransmitSocket, strSendLine)
  lngStatus = ReceiveSocketBinary(lngTransmitSocket, strSocketResponse)
  lngStatus = ReleaseSocket(lngTransmitSocket)
  
  For x = 1 To Len(strSocketResponse)
    Debug.Print Asc(Mid$(strSocketResponse, x, 1)),
    If Int(x / 16) = x / 16 Then Debug.Print
  Next x
  
  Set dmsgResponseMessage = New DNSMessage
  dmsgResponseMessage.Parse (Mid(strSocketResponse, 3))
 
  Debug.Print dmsgResponseMessage.ID
  Debug.Print dmsgResponseMessage.Question(1).QName
  
  lstResourceRecords.Clear
  For x = 1 To dmsgResponseMessage.ANCOUNT
    lstResourceRecords.AddItem Right$(Space(5) + Str(dmsgResponseMessage.Answer(x).RDATA.Preference), 5) + _
                               " " + dmsgResponseMessage.Answer(x).RDATA.Exchange
  Next
    
  txtDomain.SelStart = 0
  txtDomain.SelLength = Len(txtDomain.Text)
  txtDomain.SetFocus

End Sub

Private Sub btnQuit_Click()
    
    lngStatus = WSACleanup()
    Debug.Print "WSACleanup lngStatus " & SocketError(lngStatus)

    End

End Sub
