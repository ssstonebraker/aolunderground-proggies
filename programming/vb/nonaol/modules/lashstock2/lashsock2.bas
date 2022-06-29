Attribute VB_Name = "lashstock2"
'lashstock2.bas by lash
'(c)1999 lash productions99
'tb_lash@juno.com

Option Explicit

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type


Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
   
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Function GetIPAddress() As String

'example
'x = GetIPAddress
'Label1.Caption = "ip: " + x + ""

   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function
Public Function GetIPHostName() As String

'example
'l = GetIPHostName
'Label1.Caption = "hostname: " + l + ""

    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function
Public Function HiByte(ByVal wParam As Integer)

    HiByte = wParam \ &H100 And &HFF&
 
End Function
Public Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function
Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
   End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    SocketsInitialize = True
End Function




