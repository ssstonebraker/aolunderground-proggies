Attribute VB_Name = "modVBWinSock"
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

Declare Function lstrcpy& Lib "Kernel32" (ByVal dststring As Any, ByVal srcstring As Any)

Declare Sub hmemcpy Lib "Kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbBytes&)

Global Status As Integer
Global CRLF As String
Global CallSocket As InputSocketDescriptor

Type InputSocketDescriptor
  Family As Integer
  Port As Integer
  Address As Long
  Fill As String * 8
End Type

Type OutputSocketDescriptor
  Family As Integer
  Port As Integer
  Address As Long
  Fill As String * 30
End Type


Type WSAdata_type
  wVersion As Integer
  wHighVersion As Integer
  szDescription As String * 257
  szSystemStatus As String * 129
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpVendorInfo As String * 200
End Type
Global WSAdata As WSAdata_type

Type NameHostEntry
  Name As Long
  Aliases As Long
  AddressType As Integer
  Length As Integer
  AddressPointerPointer As Long
  Buffer As String * 100
End Type

Type sockopt_bool_type         'Used for setting socket options
  val As Integer
End Type
'Global sockopt_bool As sockopt_bool_type

  'These are the Win specific calls which use messages
Declare Function WSAStartup Lib "wsock32.dll" (ByVal a As Integer, B As WSAdata_type) As Integer
Declare Function WSACleanup Lib "wsock32.dll" () As Integer
Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal sock As Integer, ByVal hwnd As Long, ByVal msg As Integer, ByVal event2 As Long) As Integer
Declare Function WSAGetLastError Lib "wsock32.dll" () As Integer

  'Winsock calls in VB format
Declare Function gethostbyname Lib "wsock32.dll" (ByVal HostName As String) As Long
Declare Function gethostbyaddr Lib "wsock32.dll" (HostAddress As Long, ByVal lenaddress As Integer, ByVal pftype As Integer) As Long
Declare Function inet_ntoa Lib "wsock32.dll" (ByVal iaddr As Long) As Long
Declare Function htons Lib "wsock32.dll" (ByVal a As Integer) As Integer

Declare Function socket Lib "wsock32.dll" (ByVal af As Integer, ByVal typesock As Integer, ByVal protocol As Integer) As Integer
Declare Function bind Lib "wsock32.dll" (ByVal s As Integer, addr As InputSocketDescriptor, ByVal namelen As Integer) As Integer
Declare Function connect Lib "wsock32.dll" (ByVal sock As Integer, sockstruct As InputSocketDescriptor, ByVal structlen As Integer) As Integer
Declare Function send Lib "wsock32.dll" (ByVal sock As Integer, ByVal msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function recv Lib "wsock32.dll" (ByVal sock As Integer, ByVal msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function closesocket Lib "wsock32.dll" (ByVal sock As Integer) As Integer
Declare Function setsockopt Lib "wsock32.dll" (ByVal sock As Integer, ByVal level As Integer, ByVal optname As Integer, optval As sockopt_bool_type, ByVal optlen As Integer) As Integer

Declare Function listen Lib "wsock32.dll" (ByVal sn As Integer, ByVal blog As Integer) As Integer
Declare Function accept Lib "wsock32.dll" (ByVal sn As Integer, saddr As OutputSocketDescriptor, namelen As Integer) As Integer


'Some constants declarations
Global Const SOCKET_ERROR = -1
Global Const INVALID_SOCKET = -1

Global Const SOCK_STREAM = 1
Global Const AF_INET = 2
Global Const PF_INET = 2

Global Const IPPROTO_TCP = 6
Global Const SOL_SOCKET = &HFFFF
Global Const SO_DEBUG = &H1
Global Const SO_ACCEPTCONN = &H2
Global Const SO_REUSEADDR = &H4
Global Const SO_KEEPALIVE = &H8
Global Const SO_DONTROUTE = &H10
Global Const SO_BROADCAST = &H20
Global Const SO_USELOOPBACK = &H40
Global Const SO_LINGER = &H80
Global Const SO_OOBINLINE = &H100
Global Const SO_DONTLINGER = &HFF7F

Global Const AF_UNSPEC = 0

Global Const FD_READ = 1
Global Const FD_WRITE = 2
Global Const FD_OOB = 4
Global Const FD_ACCEPT = 8
Global Const FD_CONNECT = &H10
Global Const FD_CLOSE = &H20

Function StartWinSock()
  
  Dim RequestedVersion As Integer
  
  CRLF = Chr(13) + Chr(10)
  frmVBWinSockDialog.Show
  
  Debug.Print "StartWinSock"
  RequestedVersion = &H101    ' Version 1.1
  Status = WSAStartup(RequestedVersion, WSAdata)
  
  If Status Then
    DebugLastError "WinSock startup "
    StartWinSock = False
    Exit Function
  End If
    
  Debug.Print "Startup status " & SocketError(Status)
  Debug.Print "Version " & WSAdata.wVersion
  Debug.Print "High version " & WSAdata.wHighVersion
  Debug.Print "Description " & WSAdata.szDescription
  Debug.Print "System status " & WSAdata.szSystemStatus
  Debug.Print "Max sockets " & WSAdata.iMaxSockets
  Debug.Print "Max datagrams " & WSAdata.iMaxUdpDg
  
  Debug.Print "Pointer to vendor info " & WSAdata.lpVendorInfo
  StartWinSock = True

End Function

Function GetIPAddress(HostAddress As Long, HostName As String) As Boolean
    
  Dim nheHostRecord As NameHostEntry
  Dim lngHostRecordPointer As Long
  Dim lngHostAddressPointer As Long
  Dim lngHostAddressPointerPointer As Long
  
  lngHostAddressPointerPointer = 0
  lngHostAddressPointer = 0
  lngHostRecordPointer = 0

  lngHostRecordPointer = gethostbyname(HostName)
  If lngHostRecordPointer <> 0 Then
      hmemcpy nheHostRecord, ByVal lngHostRecordPointer, Len(nheHostRecord)
  End If
  lngHostAddressPointerPointer = nheHostRecord.AddressPointerPointer
  
  If lngHostAddressPointerPointer <> 0 Then
      hmemcpy lngHostAddressPointer, ByVal lngHostAddressPointerPointer&, 4
  End If
  If lngHostAddressPointerPointer <> 0 Then
      hmemcpy HostAddress, ByVal lngHostAddressPointer, 4
  End If
  
  Debug.Print HostAddress
  dota& = inet_ntoa&(HostAddress)
  dotaddr$ = Space(256)
  temp& = lstrcpy&(dotaddr$, dota&)
  Debug.Print dotaddr$
  
  GetIPAddress = True
End Function

Function CreateSocket(TCPSocket As Long, ByVal TCPPort As Long)

  Debug.Print "Creating socket"
  
    'Get a socket/handle from Winsock
  TCPSocket = socket(PF_INET, SOCK_STREAM, AF_UNSPEC)
  If TCPSocket = INVALID_SOCKET Then
    DebugLastError "Creation of socket"
    CreateSocket = False
    Exit Function
  End If

  Debug.Print "Got socket handle " & TCPSocket
  
    'Now bind the socket
  Debug.Print "Bind socket " & TCPSocket
  
    'We pass this structure during a bind.
  CallSocket.Family = AF_INET
  CallSocket.Port = htons(TCPPort)
  CallSocket.Address = 0
  CallSocket.Fill = " "
  
  Status = bind(TCPSocket, CallSocket, Len(CallSocket))
  If Status = SOCKET_ERROR Then
    DebugLastError "Bind of socket"
    CreateSocket = False
    Exit Function
  End If
  
  Debug.Print "Bind OK"
  CreateSocket = True
  
End Function

Function ConnectSocket(TCPSocket As Long, HostAddress As Long, ByVal TCPPort As Long)

  CallSocket.Family = AF_INET
  CallSocket.Port = htons(TCPPort)     'converts byte order from PC to Network
  CallSocket.Address = HostAddress
  CallSocket.Fill = " "

  Debug.Print "Connect socket " & TCPSocket
  Status = connect(TCPSocket, CallSocket, Len(CallSocket))
  If Status = SOCKET_ERROR Then
    DebugLastError "Connecting socket"
    ConnectSocket = False
    Exit Function
  End If
      
  Debug.Print "Connected to server OK."
  ConnectSocket = True
  
End Function

Function SendSocketBinary(TCPSocket As Long, FullBuffer As String)

  Dim RemainingBuffer As String
  Dim RemainingLength As Long
  
  DialogPrintBinary "[" + Trim$(Str$(TCPSocket)) + "-->]", FullBuffer
  
  RemainingBuffer = FullBuffer
  
  Do
    RemainingLength = Len(RemainingBuffer)
    Status = send(TCPSocket, RemainingBuffer, RemainingLength, 0)
    
    If Status = SOCKET_ERROR Then
      DebugLastError "Send to socket " + Str(TCPSocket)
      SendSocketBinary = False
      Exit Function
    ElseIf Status = RemainingLength Then
      Debug.Print "Send was OK."
      SendSocketBinary = True
      Exit Do
    Else
      RemainingBuffer = Mid(RemainingBuffer, Status + 1)
      Debug.Print "Partial send of " & Status & " bytes"
    End If
  Loop

End Function


Function SendSocket(TCPSocket As Long, FullBuffer As String, Optional OmitCRLF As Boolean) As Boolean
  Dim RemainingBuffer As String
  Dim RemainingLength As Long
  
  DialogPrint "[" + Trim$(Str$(TCPSocket)) + "-->]", FullBuffer
  Debug.Print FullBuffer
  
  If OmitCRLF Then
    RemainingBuffer = FullBuffer
  Else
    RemainingBuffer = FullBuffer + Chr$(13) + Chr$(10)
  End If
  
  Do
    RemainingLength = Len(RemainingBuffer)
    Status = send(TCPSocket, RemainingBuffer, RemainingLength, 0)
    
    If Status = SOCKET_ERROR Then
      DebugLastError "Send to socket " + Str(TCPSocket)
      SendSocket = False
      Exit Function
    ElseIf Status = RemainingLength Then
      Debug.Print "Send was OK."
      SendSocket = True
      Exit Do
    Else
      RemainingBuffer = Mid(RemainingBuffer, Status + 1)
      Debug.Print "Partial send of " & Status & " bytes"
    End If
  Loop
  
End Function

Function ReceiveSocketBinary(TCPSocket As Long, SocketResponse As String) As Long
  Dim ResponseLength As Integer
  
      'Now wait for the response from the server.
  SocketResponse = Space$(4096)
  Status = recv(TCPSocket, SocketResponse, 4096, 0)
  
  If Status = SOCKET_ERROR Then
    DebugLastError "Receive from socket " + Str(TCPSocket)
    ReceiveSocketBinary = False
    Exit Function
  End If
    
  ResponseLength = Status
  SocketResponse = Left(SocketResponse, ResponseLength)
  Debug.Print "Successful receive of " & ResponseLength & " bytes"
  DialogPrintBinary "[-->" + Trim$(Str$(TCPSocket)) + "] ", SocketResponse
  ReceiveSocketBinary = True
  Debug.Print ReceiveSocketBinary
  
End Function

Function ReceiveSocket(TCPSocket As Long, SocketResponse As String, Optional Port As Integer) As Boolean
  Static strResponseRemainder(10) As Variant
  Dim intLoopCount As Integer
  
  If IsMissing(Port) Then Port = 1
  
  If Port >= 0 Then
    intLoopCount = 0
    Do While InStr(1, strResponseRemainder(Port), CRLF) = 0
      SocketResponse = Space$(4096)
      Do
        Status = recv(TCPSocket, SocketResponse, 4096, 0)
        intLoopCount = intLoopCount + 1
      Loop Until Status <> SOCKET_ERROR Or _
                 WSAGetLastError() <> 0 Or _
                 intLoopCount > 30
                 
      If TCPSocket = 0 Then
        MsgBox hello
        End
      End If
      
      If (Status = SOCKET_ERROR) Then
        SocketResponse = ""
        DebugLastError "Receive from socket " + Str(TCPSocket)
        ReceiveSocket = False
        Exit Function
      End If
      Debug.Print "Successful receive of " & Status & " bytes"
      
      strResponseRemainder(Port) = strResponseRemainder(Port) + Trim(SocketResponse)
    
    Loop
    
    intCRLFPosition = InStr(1, strResponseRemainder(Port), CRLF)
    
    SocketResponse = Left(strResponseRemainder(Port), intCRLFPosition - 1)
    strResponseRemainder(Port) = Mid(strResponseRemainder(Port), intCRLFPosition + 2)
  Else
  
    SocketResponse = Space$(32000)
    Status = recv(TCPSocket, SocketResponse, 32000, 0)
    SocketResponse = Trim(SocketResponse)
    
    If Status = SOCKET_ERROR Then
      DebugLastError "Receive from socket " + Str(TCPSocket)
      ReceiveSocket = False
      Exit Function
    End If
    Debug.Print "Successful receive of " & Status & " bytes"
  
  End If
  
  DialogPrint "[-->" + Trim$(Str$(TCPSocket)) + "] " + Chr$(13) + Chr$(10), SocketResponse
  Debug.Print SocketResponse
  
  ReceiveSocket = True

End Function

Function ReleaseSocket(TCPSocket As Long)

  If TCPSocket <= 0 Then
    Debug.Print "Socket not in use"
  End If
  
  Status = closesocket(TCPSocket)
  
  If Status = SOCKET_ERROR Then
    DebugLastError "Close of socket " & Str(TCPSocket)
    ReleaseSocket = False
    Exit Function
  End If
  
  Debug.Print "Closed socket " & TCPSocket & " OK"
  ReleaseSocket = True

End Function

Sub DialogPrintBinary(ByVal PrintPrefix As String, ByVal PrintText As String)
  Dim PrintLine As String
  Dim x As Integer
  Dim DialogText As String
  
  DialogText = frmVBWinSockDialog!txtDialog.Text
  
  DialogText = DialogText + PrintPrefix + Chr(13) + Chr(10)
  
  For x = 1 To Len(PrintText)
    DialogText = DialogText + Right(Space(4) + Str(Asc(Mid$(PrintText, x, 1))), 4)
    If Int(x / 16) = x / 16 Then DialogText = DialogText + Chr(13) + Chr(10)
  Next x
  DialogText = DialogText + Chr(13) + Chr(10) + Chr(13) + Chr(10)
  If Len(DialogText) > 30000 Then DialogText = Right(DialogText, 30000)
    
  frmVBWinSockDialog!txtDialog.Text = DialogText
  frmVBWinSockDialog!txtDialog.Refresh


End Sub

Sub DialogPrint(ByVal PrintPrefix As String, ByVal PrintText As String)
    Dim DialogText As String
    Dim PrintLine As String
    
    DialogText = frmVBWinSockDialog!txtDialog.Text
    
    PrintLine = PrintPrefix + " >" + Trim$(PrintText) + "<"
    
    If Right$(PrintLine, 2) <> Chr$(13) + Chr$(10) Then
        PrintLine = PrintLine + Chr$(13) + Chr$(10)
    End If
    
    DialogText = DialogText + PrintLine
    If Len(DialogText) > 30000 Then DialogText = Right(DialogText, 30000)
    
    frmVBWinSockDialog!txtDialog.Text = DialogText
    frmVBWinSockDialog!txtDialog.Refresh

End Sub

Function SocketError(ErrorCode) As String
  
  Select Case ErrorCode
    Case 0:     SocketError = "OK"
    Case 10004: SocketError = "WSAEINTR"
    Case 10009: SocketError = "WSAEBADF"
    Case 10013: SocketError = "WSAEACCES"
    Case 10014: SocketError = "WSAEFAULT"
    Case 10022: SocketError = "WSAEINVAL"
    Case 10024: SocketError = "WSAEMFILE"
    Case 10035: SocketError = "WSAEWOULDBLOCK"
    Case 10036: SocketError = "WSAEINPROGRESS"
    Case 10037: SocketError = "WSAEALREADY"
    Case 10038: SocketError = "WSAENOTSOCK"
    Case 10039: SocketError = "WSAEDESTADDRREQ"
    Case 10040: SocketError = "WSAEMSGSIZE"
    Case 10041: SocketError = "WSAEPROTOTYPE"
    Case 10042: SocketError = "WSAENOPROTOOPT"
    Case 10043: SocketError = "WSAEPROTONOSUPPORT"
    Case 10044: SocketError = "WSAESOCKTNOSUPPORT"
    Case 10045: SocketError = "WSAEOPNOTSUPP"
    Case 10046: SocketError = "WSAEPFNOSUPPORT"
    Case 10047: SocketError = "WSAEAFNOSUPPORT"
    Case 10048: SocketError = "WSAEADDRINUSE"
    Case 10049: SocketError = "WSAEADDRNOTAVAIL"
    Case 10050: SocketError = "WSAENETDOWN"
    Case 10051: SocketError = "WSAENETUNREACH"
    Case 10052: SocketError = "WSAENETRESET"
    Case 10053: SocketError = "WSAECONNABORTED"
    Case 10054: SocketError = "WSAECONNRESET"
    Case 10055: SocketError = "WSAENOBUFS"
    Case 10056: SocketError = "WSAEISCONN"
    Case 10057: SocketError = "WSAENOTCONN"
    Case 10058: SocketError = "WSAESHUTDOWN"
    Case 10059: SocketError = "WSAETOOMANYREFS"
    Case 10060: SocketError = "WSAETIMEDOUT"
    Case 10061: SocketError = "WSAECONNREFUSED"
    Case 10062: SocketError = "WSAELOOP"
    Case 10063: SocketError = "WSAENAMETOOLONG"
    Case 10064: SocketError = "WSAEHOSTDOWN"
    Case 10065: SocketError = "WSAEHOSTUNREACH"
    Case 10066: SocketError = "WSAENOTEMPTY"
    Case 10067: SocketError = "WSAEPROCLIM"
    Case 10068: SocketError = "WSAEUSERS"
    Case 10069: SocketError = "WSAEDQUOT"
    Case 10070: SocketError = "WSAESTALE"
    Case 10071: SocketError = "WSAEREMOTE"
    Case 10091: SocketError = "WSASYSNOTREADY"
    Case 10092: SocketError = "WSAVERNOTSUPPORTED"
    Case 10093: SocketError = "WSANOTINITIALIZED"
    Case 11001: SocketError = "WSAEWSA_HOST_NOT_FOUND"
    Case 11002: SocketError = "WSATRY_AGAIN"
    Case 11003: SocketError = "WSANO_RECOVERY"
    Case 11004: SocketError = "WSANO_DATA"
    Case Else
       SocketError = "UNKNOWN ERROR CODE" + Str$(ErrorCode)
  End Select
End Function

Function LastSocketError() As String
  Status = WSAGetLastError()
  LastSocketError = SocketError(Status)
End Function

Sub DebugLastError(Action As String)
  Debug.Print "Error during " + Action + ": " + LastSocketError
End Sub
