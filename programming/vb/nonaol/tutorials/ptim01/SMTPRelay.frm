VERSION 5.00
Begin VB.Form frmSMTPRelay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "SMTP Relay"
   ClientHeight    =   7785
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
   ScaleHeight     =   7785
   ScaleWidth      =   7470
   Begin VB.TextBox txtDomain 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   1440
      Width           =   5775
   End
   Begin VB.ListBox lstResourceRecords 
      Height          =   1815
      Index           =   0
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Width           =   5655
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   10
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtMessage 
      Height          =   3735
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3960
      Width           =   7095
   End
   Begin VB.TextBox txtRequest 
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   7800
      Width           =   1095
   End
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSocket 
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblDomain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Domain:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblResourceRecords 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Resource Records:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sender:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSocket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Socket:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmSMTPRelay"
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

Dim lngListenSocket As Long
Dim lngTransmitSocket As Long
Dim lngLookupSocket As Long
Dim lngRelaySocket As Long
Dim intStatus As Integer

Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Const SMTPPort = 25
Const DNSPort = 53
Const DNSServer = "" ' To be completed

Dim intCurrentConnection As Integer
Dim CRLF As String

Private Sub Form_Load()
  Dim lngListenSocketAddress As InputSocketDescriptor

  intStatus = StartWinSock()
  
  CRLF = Chr(13) + Chr(10)
  
  cboConnection.AddItem ""
    
  txtRequest(0).Text = ""
  txtSocket(0).Text = ""
  txtMessage(0).Text = ""
  txtSender(0).Text = ""
  txtFrom(0).Text = ""
  txtTo(0).Text = ""
  
  intStatus = CreateSocket(lngListenSocket, SMTPPort)
  
  If listen(lngListenSocket, 5) Then
      MsgBox "Could not listen on Port " + Str$(SMTPPort)
      End
  End If
  
  If WSAAsyncSelect(lngListenSocket, txtConnect.hwnd, WM_MBUTTONUP, FD_ACCEPT) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

  ' Accept incoming connection
Private Sub txtConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim lngTransmitSocketAddress As OutputSocketDescriptor
  Dim lngTransmitSocketAddressSize As Integer
  Dim intCurrentConnection As Integer
  
  lngTransmitSocketAddressSize = LenB(lngTransmitSocketAddress)
  lngTransmitSocketAddressSize = 20
  
  lngTransmitSocket = accept(lngListenSocket, lngTransmitSocketAddress, lngTransmitSocketAddressSize)

  intCurrentConnection = cboConnection.ListCount
    
  cboConnection.AddItem Trim(Str(intCurrentConnection))
  Load txtSocket(intCurrentConnection)
  Load txtRequest(intCurrentConnection)
  Load txtMessage(intCurrentConnection)
  Load txtSender(intCurrentConnection)
  Load txtFrom(intCurrentConnection)
  Load txtTo(intCurrentConnection)
  Load lstResourceRecords(intCurrentConnection)
  Load txtDomain(intCurrentConnection)
  
  txtSocket(intCurrentConnection) = Trim(Str(lngTransmitSocket))
  
  DisplayConnection intCurrentConnection
  
  intStatus = SendSocket(lngTransmitSocket, "220 SMTP server here")
    
  If WSAAsyncSelect(lngTransmitSocket, txtRequest(intCurrentConnection).hwnd, WM_MBUTTONDBLCLK, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If

    
End Sub

  ' Process incoming request
Private Sub txtRequest_DblClick(Index As Integer)
  Dim strClientRequest As String
  
  intStatus = ReceiveSocket(lngTransmitSocket, strClientRequest, 1)
  DisplayConnection Index
  txtRequest(Index).Text = strClientRequest
  
  strClientRequest = Trim(strClientRequest)
  Select Case Trim(UCase(Left$(strClientRequest, 4)))
    Case "HELO"
      txtSender(Index).Text = Trim(Mid(strClientRequest, 5))
      intStatus = SendSocket(lngTransmitSocket, "250 olive.tree")
    Case "MAIL"
      txtFrom(Index).Text = FormatUser(Mid(strClientRequest, 11))
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
    Case "RCPT"
      txtTo(Index).Text = FormatUser(Mid$(strClientRequest, 9))
      txtDomain(Index).Text = Mid(txtTo(Index).Text, _
                                  InStr(1, txtTo(Index).Text, "@") + 1)
            
      LookupDomain Index
      ConnectNextRelay Index
      
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
    Case "DATA"
      RelayMessage Index
      Debug.Print
    Case "QUIT"
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
      intStatus = ReleaseSocket(lngTransmitSocket)
      intStatus = SendSocket(lngRelaySocket, strClientRequest)
      intStatus = ReceiveSocket(lngRelaySocket, strClientRequest, 2)
      intStatus = ReleaseSocket(lngRelaySocket)
      Debug.Print
    Case Else
      intStatus = SendSocket(lngTransmitSocket, "250 OK")
  End Select

End Sub

Private Function FormatUser(InputUser As String) As String

  FormatUser = Trim(InputUser)
  
  If Left(FormatUser, 1) = "<" Then FormatUser = Mid(FormatUser, 2)
  If Right(FormatUser, 1) = ">" Then FormatUser = Left(FormatUser, Len(FormatUser) - 1)

End Function


Private Sub LookupDomain(Index As Integer)
    
  Dim lngServerAddress As Long
  Dim strSocketResponse As String
  Dim strSendLine As String
  Dim bytLineLength As Byte
  Dim dmsgSendMessage As New DNSMessage
    
  Dim x As Integer
  Dim dmsgResponseMessage As New DNSMessage
    
  intStatus = GetIPAddress(lngServerAddress, DNSServer)
  intStatus = CreateSocket(lngLookupSocket, 0)
  intStatus = ConnectSocket(lngLookupSocket, lngServerAddress, DNSPort)

  dmsgSendMessage.ComposeQuestion 15, txtDomain(Index).Text
  strSendLine = dmsgSendMessage.TransferString
  
  bytLineLength = Len(strSendLine)
  
  strSendLine = Chr$(0) + Chr$(bytLineLength) + strSendLine + Chr$(0) + Chr$(0)
    
  intStatus = SendSocketBinary(lngLookupSocket, strSendLine)
  intStatus = ReceiveSocketBinary(lngLookupSocket, strSocketResponse)
  intStatus = ReleaseSocket(lngLookupSocket)
  
  dmsgResponseMessage.Parse Mid(strSocketResponse, 3)
 
  lstResourceRecords(Index).Clear
  For x = 1 To dmsgResponseMessage.ANCOUNT
    lstResourceRecords(Index).AddItem Right$(Space(5) + Str(dmsgResponseMessage.Answer(x).RDATA.Preference), 5) + _
                               " " + dmsgResponseMessage.Answer(x).RDATA.Exchange
  Next
    
End Sub

Private Sub ConnectNextRelay(Index As Integer)
  Dim lngServerAddress As Long
  Dim strNextRelay As String
  Dim strServerResponse As String
  Dim intRelayIndex As Integer
  
  intRelayIndex = 0
  
  Do
    strNextRelay = Mid(lstResourceRecords(Index).List(intRelayIndex), 7)
    
    intStatus = GetIPAddress(lngServerAddress, strNextRelay)
    intStatus = CreateSocket(lngRelaySocket, 0)
    intStatus = ConnectSocket(lngRelaySocket, lngServerAddress, SMTPPort)
    intRelayIndex = intRelayIndex + 1
  Loop Until intStatus

  intStatus = ReceiveSocket(lngRelaySocket, strServerResponse, 2)
  intStatus = SendSocket(lngRelaySocket, "mail from:<" + txtFrom(Index).Text + ">")
  intStatus = ReceiveSocket(lngRelaySocket, strServerResponse, 2)
  intStatus = SendSocket(lngRelaySocket, "rcpt to:<" + txtTo(Index).Text + ">")
  intStatus = ReceiveSocket(lngRelaySocket, strServerResponse, 2)
  intStatus = SendSocket(lngRelaySocket, "data")
  intStatus = ReceiveSocket(lngRelaySocket, strServerResponse, 2)

End Sub

Private Sub RelayMessage(Index As Integer)
  Dim strClientRequest As String
  Dim strFileName As String
  Dim intFileNumber As Integer
  
  intStatus = SendSocket(lngTransmitSocket, "354 Start sending...")
  
  txtMessage(Index).Text = ""
  Do
    intStatus = ReceiveSocket(lngTransmitSocket, strClientRequest, 1)
    If intStatus Then
      intStatus = SendSocket(lngRelaySocket, strClientRequest)
      If strClientRequest = "." Then Exit Do
      txtMessage(Index).Text = txtMessage(Index).Text + strClientRequest + CRLF
    End If
  Loop While strClientRequest <> "."
  
  intStatus = SendSocket(lngRelaySocket, strClientRequest)
  
  intStatus = SendSocket(lngTransmitSocket, "250 OK")
  intStatus = ReceiveSocket(lngRelaySocket, strClientRequest, 2)

End Sub

Private Function RandomName() As String
  RandomName = "M" + Format(Now, "yymmddhhnnss") + Trim(Str(Int(Rnd * 10000)))
End Function

Private Sub cboConnection_Click()
  DisplayConnection cboConnection.ListIndex
End Sub

Private Sub DisplayConnection(NewConnection As Integer)
  
  If NewConnection < 0 Or NewConnection >= cboConnection.ListCount Then
    NewConnection = 0
  End If
  
  If NewConnection = intCurrentConnection Then Exit Sub
  
  txtRequest(intCurrentConnection).Visible = False
  txtRequest(NewConnection).Visible = True
  txtSocket(intCurrentConnection).Visible = False
  txtSocket(NewConnection).Visible = True
  txtMessage(intCurrentConnection).Visible = False
  txtMessage(NewConnection).Visible = True
  txtSender(intCurrentConnection).Visible = False
  txtSender(NewConnection).Visible = True
  txtFrom(intCurrentConnection).Visible = False
  txtFrom(NewConnection).Visible = True
  txtTo(intCurrentConnection).Visible = False
  txtTo(NewConnection).Visible = True
  txtDomain(intCurrentConnection).Visible = False
  txtDomain(NewConnection).Visible = True
  lstResourceRecords(intCurrentConnection).Visible = False
  lstResourceRecords(NewConnection).Visible = True
      
  cboConnection.ListIndex = NewConnection
  
  intCurrentConnection = NewConnection
  
End Sub


Private Sub btnQuit_Click()
  Unload Me
  End
End Sub

Private Sub Form_Unload(CANCEL As Integer)
   
  intStatus = ReleaseSocket(lngListenSocket)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup intStatus " & SocketError(intStatus)

End Sub
