Attribute VB_Name = "modWinsock"
Attribute VB_Description = "The mother-load..."
'****************************************************
'This file passed trought:
'K.Driblinov prg page... tons of C & Vb sources, links to
'other prg sites!!
'http://www.geocities.com/SiliconValley/Lakes/7057/index.htm
'E-Mail: kdriblinov@hotmail.com
'****************************************************

'This is the Winsock API definition file for Visual Basic 4.0
'The file is a port from the C header file found in the
'Winsock Version 1.1 API specification.

'Winsock Version 2.2.x API changes are in progress :)
'They're not done yet 11/29/96

' This file 'winsock.bas' is my version 0.98 1/18/97
' Please send ALL updates, additions/revisions and
'  flagrent typo reports to david.gravereaux@snet.net

'************************** IMPORTANT ***************************
'NOTE: This is a "work in progress" and, as such, could include
'      a gross amount of errors. YOU are on your own!! At worst,
'      this is an EXCELLENT starting place for building
'      -extremely- FAST Internet Apps with Visual Basic 4.0 by
'      using Winsock API calls directly. Thus, avoiding the use
'      of [slow, resource HOG] OLE/ActiveX controls. The WHOLE
'      Window's API is available to the VB programmer.
'      So make use of it! This module contains the COMPLETE
'      Winsock API for v1.1 AND v2.2 [almost done].
'************************** IMPORTANT ***************************

'Reference books I highly recommend:
'  (1) Visual Basic Programmer's Guide to the Win32 API  by Daniel Appleman
'       ISBN 1-56276-287-7
'  (2) Network Programming with Windows Sockets  by Pat Bonner
'       ISBN 0-13-230152-0

'CONTENTS:
'  + Constants and Structures for WinSock 1.1 with additions for
'    WinSock 2.2 where needed (in order as in 'winsock[2].h')
'     + FD constants
'     + ioctl constants
'     + database structures
'     + RFC 790 constants and structures
'     + internet address constants
'     + socket types, families, and option constants
'     + error constants
'     + database error constants
'     + loads and loads of NEW WinSock2 stuff:
'        + new error codes and type definitions
'        + WSABUF and QOS struct
'        + manifest constants
'        + WSAPROTOCOL_INFO structure
'        + CSADDR constants and structures
'        + Client Query API Typedefs
'        + Service Address Registration and Deregistration Data Types
'        + data types for WSAAccept() and overlapped I/O completion routine
'  + Declares for:
'     + Win32, WinSock1.1
'     + Win32, WinSock2.2
'     + Win16, Winsock1.1
'  + Constants, structures, and declares specific to the Microsoft NT Stack


'--Changes made for Win32:
'  1. the dll is now call called 'wsock32.dll'.
'  2. socket descriptor 's' changed to long from integer in API calls.
'  3. 'hWnd' changed to long from integer to reflect the Win32 system.
'  4. socket() call set to Long, this correctly reflects the datatype returned.
'  5. 'Alias "#12"' was removed from the listen call.
'  6. h_addrtype and h_length of Hostent Structure changed to integer from
'     String * 2. 'winsock.h' clearly states these as integer types.
'  7. WSAGETSELECTERROR() and WSAGETSELECTEVENT() functions changed Lparam
'     to long instead of integer. Like it should be.
'  8. Accept() now set to long 'cause it returns a socket descriptor.


'##############################################################
' Compiler directives.
#Const WinSock1 = True
#Const WinSock2 = True
#Const UNICODE = False  'don't set this to true
#Const MIDL_PASS = False
#Const CSADDR_DEFINED = True
#Const AddMicroSoft = True
'##############################################################

Public Const FD_SETSIZE% = 64

Type FD_SET
  fd_count As Integer
  fd_array(FD_SETSIZE) As Integer
End Type 'fd_set

Public Const SZFD_SET = 4

#If Win32 Then
  Declare Function FD_ISSET% Lib "wsock32.dll" Alias "#151" (ByVal s&, ByRef passed_set As FD_SET)
  Declare Function x_WSAFDIsSet% Lib "wsock32.dll" Alias "#151" (ByVal s&, ByRef passed_set As FD_SET)
#ElseIf Win16 Then
  Declare Function FD_ISSET% Lib "winsock.dll" Alias "#151" (ByVal s%, ByRef passed_set As FD_SET)
  Declare Function x_WSAFDIsSet% Lib "winsock.dll" Alias "#151" (ByVal s%, ByRef passed_set As FD_SET)
#End If

Type timeval
  tv_sec As Long
  tv_usec As Long
End Type
Public Const SZTIMEVAL = 8

'/*
' * Commands for ioctlsocket(),  taken from the BSD file fcntl.h.
' *
' *
' * Ioctl's have the command encoded in the lower word,
' * and the size of any in or out parameters in the upper
' * word.  The high 2 bits of the upper word are used
' * to encode the in/out status of the parameter; for now
' * we restrict parameters to at most 128 bytes.
' */
Public Const IOCPARM_MASK = &H7F              ' parameters must be < 128 bytes
Public Const IOC_VOID = &H20000000            ' no parameters
Public Const IOC_OUT = &H40000000             ' copy out parameters
Public Const IOC_IN = &H80000000              ' copy in parameters
Public Const IOC_INOUT = (IOC_IN Or IOC_OUT)  ' 0x20000000 distinguishes new
                                              '   & old ioctl's

Public Const u_long& = 4     ' the size in bytes of a long integer

' The original macros are functions here as IO(), IOR(), and IOW()

Public Const FIONREAD& = &H4004667F         '/* get # bytes to read */
Public Const FIONBIO& = &H8004667E          '/* set/clear non-blocking i/o */
Public Const FIOASYNC& = &H8004667D         '/* set/clear async i/o */

'/* Socket I/O Controls */
Public Const SIOCSHIWAT& = &H80047300       '/* set high watermark */
Public Const SIOCGHIWAT& = &H40047301       '/* get high watermark */
Public Const SIOCSLOWAT& = &H80047302       '/* set low watermark */
Public Const SIOCGLOWAT& = &H40047303       '/* get low watermark */
Public Const SIOCATMARK& = &H40047307       '/* at oob mark? */


'/*
' * Structures returned by network data base library, taken from the
' * BSD file netdb.h.  All addresses are supplied in host order, and
' * returned in network order (suitable for use in system calls).
' */
Type hostent
  h_name As Long            '/* (pointer to string) official name of host */
  h_aliases As Long         '/* (pointer to string) alias list */(might be null-seperated with 2null terminated)
  h_addrtype As Integer     '/* host address type */
  h_length As Integer       '/* length of address list returned */
  h_addr_list As Long       '/* (pointer to a) list of addresses */
  'h_addr as MissingForever     h_addr is a redefine [in C] to the first
                              ' entry in h_addr_list. It can't be used here.
End Type
Public Const SZHOSTENT% = 16


' Use the following structure for calls that require a larger
' buffer space for the hostent structure [ie. WSAAsyncGetHostbyXXX()]
Type BIGhostent 'size = MAXGETHOSTSTRUCT
  h_name As Long
  h_aliases As Long
  h_addrtype As Integer
  h_length As Integer
  h_addr_list As Long
  'h_addr as MissingForever
  blank(1008) As Byte       'Buffer space allocated
End Type



'/*
' * It is assumed here that a network number
' * fits in 32 bits.
' */
Type netent
  n_name As Long            '/* (pointer to string) official name of net */
  n_aliases As Long         '/* (pointer to string) alias list */(might be null-seperated with 2null terminated)
  n_addrtype As Integer     '/* net address type */
  n_net As Long             '/* network # */
End Type

Type servent
  s_name As Long            '/* (pointer to string) official service name */
  s_aliases As Long         '/* (pointer to string) alias list */(might be null-seperated with 2null terminated)
  s_port As Integer         '/* port # */
  s_proto As Long           '/* (pointer to) protocol to use */
End Type

Type protoent
  p_name As Long            '/* (pointer to string) official protocol name */
  p_aliases As Long         '/* (pointer to string) alias list */(might be null-seperated with 2null terminated)
  p_proto As Integer        '/* protocol # */
End Type

'/*
' * Constants and structures defined by the internet system,
' * Per RFC 790, September 1981, taken from the BSD file netinet/in.h.
' */

'/*
' * Protocols
' */
Public Const IPPROTO_IP% = 0           ' dummy for IP
Public Const IPPROTO_ICMP% = 1         ' internet control message protocol
Public Const IPPROTO_IGMP = 2          ' internet group management protocol */
Public Const IPPROTO_GGP% = 3          ' gateway^2 (deprecated)
Public Const IPPROTO_TCP% = 6          ' tcp
Public Const IPPROTO_PUP% = 12         ' pup
Public Const IPPROTO_UDP% = 17         ' user datagram protocol
Public Const IPPROTO_IDP% = 22         ' xns idp
Public Const IPPROTO_ND% = 77          ' UNOFFICIAL net disk proto

Public Const IPPROTO_RAW% = 255        ' raw IP packet
Public Const IPPROTO_MAX% = 256

'/*
' * Port/socket numbers: network standard functions
' */
Public Const IPPORT_ECHO% = 7
Public Const IPPORT_DISCARD% = 9
Public Const IPPORT_SYSTAT% = 11
Public Const IPPORT_DAYTIME% = 13
Public Const IPPORT_NETSTAT% = 15
Public Const IPPORT_FTP% = 21
Public Const IPPORT_TELNET% = 23
Public Const IPPORT_SMTP% = 25
Public Const IPPORT_TIMESERVER% = 37
Public Const IPPORT_NAMESERVER% = 42
Public Const IPPORT_WHOIS% = 43
Public Const IPPORT_MTP% = 57

'/*
' * Port/socket numbers: host specific functions
' */
Public Const IPPORT_TFTP% = 69
Public Const IPPORT_RJE% = 77
Public Const IPPORT_FINGER% = 79
Public Const IPPORT_TTYLINK% = 87
Public Const IPPORT_SUPDUP% = 95

'/*
' * UNIX TCP sockets
' */
Public Const IPPORT_EXECSERVER% = 512
Public Const IPPORT_LOGINSERVER% = 513
Public Const IPPORT_CMDSERVER% = 514
Public Const IPPORT_EFSSERVER% = 520

'/*
' * UNIX UDP sockets
' */
Public Const IPPORT_BIFFUDP% = 512
Public Const IPPORT_WHOSERVER% = 513
Public Const IPPORT_ROUTESERVER% = 520   '/* 520+1 also used */

'/*
' * Ports < IPPORT_RESERVED are reserved for
' * privileged processes (e.g. root).
' */
Public Const IPPORT_RESERVED% = 1024

'/*
' * Link numbers
' */
Public Const IMPLINK_IP% = 155
Public Const IMPLINK_LOWEXPER% = 156
Public Const IMPLINK_HIGHEXPER% = 158

'/*
' * Internet address (32-bit)
' */
Type in_addr
 s_addr As Long           '/* long IP in network-order */
End Type

'/*
' * Definitions of bits in internet address integers.
' * On subnets, the decomposition of addresses to host and net parts
' * is done according to the subnet mask.
' */
Public Const IN_CLASSA_NET& = &HFF000000
Public Const IN_CLASSA_NSHIFT& = 24
Public Const IN_CLASSA_HOST& = &HFFFFFF
Public Const IN_CLASSA_MAX& = 128

Public Const IN_CLASSB_NET& = &HFFFF0000
Public Const IN_CLASSB_NSHIFT& = 16
Public Const IN_CLASSB_HOST& = &HFFFF
Public Const IN_CLASSB_MAX& = 65536

Public Const IN_CLASSC_NET& = &HFF00
Public Const IN_CLASSC_NSHIFT& = 8
Public Const IN_CLASSC_HOST& = &HFF

Public Const INADDR_ANY& = &H0
Public Const INADDR_LOOPBACK& = &H7F000001
Public Const INADDR_BROADCAST& = &HFFFF
Public Const INADDR_NONE& = &HFFFF

'/*
' * Socket address, internet style.
' */
Type sockaddr_in            'size is 16 bytes
  sin_family As Integer
  sin_port As Integer       ' port in network-order
  sin_addr As in_addr       ' long IP in network-order
  sin_zero(8) As Byte       ' this is padding to make the length 16 bytes
End Type

Public Const WSADESCRIPTION_LEN% = 256
Public Const WSASYS_STATUS_LEN% = 128

'Setup the structure for the information returned from
'the WSAStartup() function.
Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription As String * 257
   szSystemStatus As String * 129
   iMaxSockets As Integer   'defined as an unsigned short {0 - 65535} in 'winsock.h'
                            ' (we should convert this to unsigned and place
                            ' in a long {must stay Integer here for length reasons})
                            ' {*not* used for WinSock2.0 and above}
   iMaxUdpDg As Integer     'ditto above (iMaxSockets and iMaxUdpDg might return a
                            ' negative number if we don't convert!)
                            ' {*not* used for WinSock2.0 and above}
   lpVendorInfo As Long     '(pointer to a string) {*not* used for WinSock2.0 and above}
End Type

'/*
' * Options for use with [get/set]sockopt at the IP level.
' */
Public Const IP_OPTIONS% = 1   '/* set/get IP per-packet options */


'/*
' * Definitions related to sockets: types, address families, options,
' * taken from the BSD file sys/socket.h.
' */
'#If WinSock2 Then
'  Public Const INVALID_SOCKET& = &HFFFF
'#Else
  Public Const INVALID_SOCKET% = &HFFFF
'#End If

Public Const SOCKET_ERROR% = -1

'/*
' * The  following  may  be used in place of the address family, socket type, or
' * protocol  in  a  call  to (WSA)Socket to indicate that the corresponding value
' * should  be taken from the supplied WSAPROTOCOL_INFO structure instead of the
' * parameter itself.
' */
Public Const FROM_PROTOCOL_INFO% = -1

'Define socket types
Public Const SOCK_STREAM% = 1     'Stream socket
Public Const SOCK_DGRAM% = 2      'Datagram socket
Public Const SOCK_RAW% = 3        'Raw data socket  [not widly supported by all WinSock Vendors]
Public Const SOCK_RDM% = 4        'Reliable Delivery socket
Public Const SOCK_SEQPACKET% = 5  'Sequenced Packet socket

'/*
' * Option flags per-socket.
' */
Public Const SO_DEBUG% = &H1               ' turn on debugging info recording [vendor-specific]
Public Const SO_ACCEPTCONN% = &H2          ' socket has had listen()
Public Const SO_REUSEADDR% = &H4           ' allow local address reuse
Public Const SO_KEEPALIVE% = &H8           ' keep connections alive
Public Const SO_DONTROUTE% = &H10          ' just use interface addresses
Public Const SO_BROADCAST% = &H20          ' permit sending of broadcast msgs
Public Const SO_USELOOPBACK% = &H40        ' bypass hardware when possible
Public Const SO_LINGER% = &H80             ' linger on close if data present
Public Const SO_OOBINLINE% = &H100         ' leave received OOB data in line
Public Const SO_DONTLINGER% = &HFF7F       ' same as SO_LINGER

'/*
' * Additional options.
' */
Public Const SO_SNDBUF% = &H1001           ' send buffer size
Public Const SO_RCVBUF% = &H1002           ' receive buffer size
Public Const SO_SNDLOWAT& = &H1003         ' send low-water mark
Public Const SO_RCVLOWAT& = &H1004         ' get low-water mark
Public Const SO_SNDTIMEO% = &H1005         ' send timeout
Public Const SO_RCVTIMEO% = &H1006         ' receive timeout
Public Const SO_ERROR% = &H1007            ' get error status and clear
Public Const SO_TYPE% = &H1008             ' get socket type

'/*
' * WinSock 2 extension -- new options
' */
#If WinSock2 Then
Public Const SO_GROUP_ID = &H2001           ' ID of a socket group */
Public Const SO_GROUP_PRIORITY = &H2002     ' the relative priority within a group*/
Public Const SO_MAX_MSG_SIZE = &H2003       ' maximum message size */
#If UNICODE Then
  Public Const SO_PROTOCOL_INFO = &H2005
#Else
  Public Const SO_PROTOCOL_INFO = &H2004
#End If
#End If

Public Const PVD_CONFIG = &H3001             ' configuration info for service provider */

'/*
' * TCP options flag.
' */
Public Const TCP_NODELAY% = &H1


'/*
' * Address families.
' */
Public Const AF_UNSPEC% = 0                ' unspecified
'/*
' * Although  AF_UNSPEC  is  defined for backwards compatibility, using
' * AF_UNSPEC for the "af" parameter when creating a socket is STRONGLY
' * DISCOURAGED.    The  interpretation  of  the  "protocol"  parameter
' * depends  on the actual address family chosen.  As environments grow
' * to  include  more  and  more  address families that use overlapping
' * protocol  values  there  is  more  and  more  chance of choosing an
' * undesired address family when AF_UNSPEC is used.
' */
Public Const AF_UNIX% = 1                ' local to host (pipes, portals)
Public Const AF_INET% = 2                ' internetwork: UDP, TCP, etc.
Public Const AF_IMPLINK% = 3             ' arpanet imp addresses
Public Const AF_PUP% = 4                 ' pup protocols: e.g. BSP
Public Const AF_CHAOS% = 5               ' mit CHAOS protocols
Public Const AF_NS% = 6                  ' XEROX NS protocols
Public Const AF_IPX% = AF_NS             ' IPX protocols: IPX, SPX, etc.
Public Const AF_ISO% = 7                 ' ISO protocols
Public Const AF_OSI% = AF_ISO            ' OSI is ISO
Public Const AF_ECMA% = 8                ' european computer manufacturers
Public Const AF_DATAKIT% = 9             ' datakit protocols
Public Const AF_CCITT% = 10              ' CCITT protocols, X.25 etc
Public Const AF_SNA% = 11                ' IBM SNA
Public Const AF_DECnet% = 12             ' DECnet
Public Const AF_DLI% = 13                ' Direct data link interface
Public Const AF_LAT% = 14                ' LAT
Public Const AF_HYLINK% = 15             ' NSC Hyperchannel
Public Const AF_APPLETALK% = 16          ' AppleTalk
Public Const AF_NETBIOS% = 17            ' NetBios-style addresses
#If WinSock2 Then
  Public Const AF_VOICEVIEW% = 18        ' VoiceView
  Public Const AF_FIREFOX% = 19          ' Protocols from Firefox
  Public Const AF_UNKNOWN1% = 20         ' Somebody is using this!
  Public Const AF_BAN% = 21              ' Banyan
  Public Const AF_ATM% = 22              ' Native ATM Services
  Public Const AF_INET6% = 23            ' Internetwork Version 6
  Public Const AF_MAX% = 24
#Else
  Public Const AF_MAX% = 18
#End If
  

'/*
' * Structure used by kernel to store most
' * addresses.
' */
Type sockaddr
  sa_family As Integer          '/* address family */
  sa_data(14) As Byte           '/* up to 14 bytes of direct address */
End Type
Public Const SADDRLEN% = 16

'/*
' * Structure used by kernel to pass protocol
' * information in raw sockets.
' */
Type sockproto
  sp_family As Integer          '/* address family */
  sp_protocol As Integer        '/* protocol */
End Type


'/*
' * Protocol families, same as address families for now.
' */
Public Const PF_UNSPEC% = AF_UNSPEC
Public Const PF_UNIX% = AF_UNIX
Public Const PF_INET% = AF_INET
Public Const PF_IMPLINK% = AF_IMPLINK
Public Const PF_PUP% = AF_PUP
Public Const PF_CHAOS% = AF_CHAOS
Public Const PF_NS% = AF_NS
Public Const PF_IPX% = AF_IPX
Public Const PF_ISO% = AF_ISO
Public Const PF_OSI% = AF_OSI
Public Const PF_ECMA% = AF_ECMA
Public Const PF_DATAKIT% = AF_DATAKIT
Public Const PF_CCITT% = AF_CCITT
Public Const PF_SNA% = AF_SNA
Public Const PF_DECnet% = AF_DECnet
Public Const PF_DLI% = AF_DLI
Public Const PF_LAT% = AF_LAT
Public Const PF_HYLINK% = AF_HYLINK
Public Const PF_APPLETALK% = AF_APPLETALK
#If WinSock2 Then
  Public Const PF_VOICEVIEW% = AF_VOICEVIEW
  Public Const PF_FIREFOX% = AF_FIREFOX
  Public Const PF_UNKNOWN1% = AF_UNKNOWN1
  Public Const PF_BAN% = AF_BAN
  Public Const PF_ATM% = AF_ATM
  Public Const PF_INET6% = AF_INET6
#End If
Public Const PF_MAX% = AF_MAX


'/*
' * Structure used for manipulating linger option.
' */
Type linger
  l_onoff As Integer          '/* option on/off */
  l_linger As Integer         '/* linger time */
End Type

'/*
' * Level number for (get/set)sockopt() to apply to socket itself.
' */
Public Const SOL_SOCKET% = &HFFFF     '/* options for socket level */

'/*
' * Maximum queue length specifiable by listen.
' */
#If WinSock2 Then
  Public Const SOMAXCONN& = &H7FFFFFFF  ' maximum # of queued connections prior
#Else                                   '   to accept()
  Public Const SOMAXCONN% = 5
#End If

Public Const MSG_OOB% = &H1             '/* process out-of-band data */
Public Const MSG_PEEK% = &H2            '/* peek at incoming message */
Public Const MSG_DONTROUTE% = &H4       '/* send without using routing tables */

Public Const MSG_PARTIAL = &H8000       '/* partial send or recv for message xport */

'/*
' * WinSock 2 extension -- new flags for WSASend(), WSASendTo(), WSARecv() and
' *                          WSARecvFrom()
' */
#If WinSock2 Then
Public Const MSG_INTERRUPT = &H10       '/* send/recv in the interrupt context */
#End If

Public Const MSG_MAXIOVLEN% = 16

'/*
' * Define constant based on rfc883, used by gethostbyxxxx() calls.
' */
Public Const MAXGETHOSTSTRUCT% = 1024

'/*
' * bit values and indices for FD_XXX network events
' */
Public Const FD_READ% = &H1
Public Const FD_WRITE% = &H2
Public Const FD_OOB% = &H4
Public Const FD_ACCEPT% = &H8
Public Const FD_CONNECT% = &H10
Public Const FD_CLOSE% = &H20

'/*
' * WinSock 2 extensions -- bit values and indices for FD_XXX network events
' */
#If WinSock2 Then
Public Const FD_QOS% = &H40
Public Const FD_GROUP_QOS% = &H80
Public Const FD_ALL_EVENTS% = &HFF
Public Const FD_MAX_EVENTS% = 8
#End If

'/*
' * All Windows Sockets error constants are biased by WSABASEERR from
' * the "normal"
' */
Public Const WSABASEERR% = 10000
'/*
' * Windows Sockets definitions of regular Microsoft C error constants
' */
Public Const WSAEINTR% = (WSABASEERR + 4)
Public Const WSAEBADF% = (WSABASEERR + 9)
Public Const WSAEACCES% = (WSABASEERR + 13)
Public Const WSAEFAULT% = (WSABASEERR + 14)
Public Const WSAEINVAL% = (WSABASEERR + 22)
Public Const WSAEMFILE% = (WSABASEERR + 24)

'/*
' * Windows Sockets definitions of regular Berkeley error constants
' */
Public Const WSAEWOULDBLOCK% = (WSABASEERR + 35)
Public Const WSAEINPROGRESS% = (WSABASEERR + 36)
Public Const WSAEALREADY% = (WSABASEERR + 37)
Public Const WSAENOTSOCK% = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ% = (WSABASEERR + 39)
Public Const WSAEMSGSIZE% = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE% = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT% = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT% = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT% = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP% = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT% = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT% = (WSABASEERR + 47)
Public Const WSAEADDRINUSE% = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL% = (WSABASEERR + 49)
Public Const WSAENETDOWN% = (WSABASEERR + 50)
Public Const WSAENETUNREACH% = (WSABASEERR + 51)
Public Const WSAENETRESET% = (WSABASEERR + 52)
Public Const WSAECONNABORTED% = (WSABASEERR + 53)
Public Const WSAECONNRESET% = (WSABASEERR + 54)
Public Const WSAENOBUFS% = (WSABASEERR + 55)
Public Const WSAEISCONN% = (WSABASEERR + 56)
Public Const WSAENOTCONN% = (WSABASEERR + 57)
Public Const WSAESHUTDOWN% = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS% = (WSABASEERR + 59)
Public Const WSAETIMEDOUT% = (WSABASEERR + 60)
Public Const WSAECONNREFUSED% = (WSABASEERR + 61)
Public Const WSAELOOP% = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG% = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN% = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH% = (WSABASEERR + 65)
Public Const WSAENOTEMPTY% = (WSABASEERR + 66)
Public Const WSAEPROCLIM% = (WSABASEERR + 67)
Public Const WSAEUSERS% = (WSABASEERR + 68)
Public Const WSAEDQUOT% = (WSABASEERR + 69)
Public Const WSAESTALE% = (WSABASEERR + 70)
Public Const WSAEREMOTE% = (WSABASEERR + 71)

'/*
' * Extended Windows Sockets error constant definitions
' */
Public Const WSASYSNOTREADY% = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED% = (WSABASEERR + 92)
Public Const WSANOTINITIALISED% = (WSABASEERR + 93)

'/*
' * WinSock 2 extensions -- Extended Windows Sockets error constant definitions
' */
#If WinSock2 Then
  Public Const WSAEDISCON = (WSABASEERR + 101)
  Public Const WSAENOMORE = (WSABASEERR + 102)
  Public Const WSAECANCELLED = (WSABASEERR + 103)
  Public Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
  Public Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
  Public Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
  Public Const WSASYSCALLFAILURE = (WSABASEERR + 107)
  Public Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
  Public Const WSATYPE_NOT_FOUND = (WSABASEERR + 109)
  Public Const WSA_E_NO_MORE = (WSABASEERR + 110)
  Public Const WSA_E_CANCELLED = (WSABASEERR + 111)
  Public Const WSAEREFUSED = (WSABASEERR + 112)
#End If


'/*
' * Error return codes from gethostbyname() and gethostbyaddr()
' * (when using the resolver). Note that these errors are
' * retrieved via WSAGetLastError() and must therefore follow
' * the rules for avoiding clashes with error numbers from
' * specific implementations or language run-time systems.
' * For this reason the codes are based at WSABASEERR+1001.
' * Note also that [WSA]NO_ADDRESS is defined only for
' * compatibility purposes.
' */

'/* Authoritative Answer: Host not found */
Public Const WSAHOST_NOT_FOUND% = (WSABASEERR + 1001)
Public Const HOST_NOT_FOUND% = WSAHOST_NOT_FOUND

'/* Non-Authoritative: Host not found, or SERVERFAIL */
Public Const WSATRY_AGAIN% = (WSABASEERR + 1002)
Public Const TRY_AGAIN% = WSATRY_AGAIN

'/* Non recoverable errors, FORMERR, REFUSED, NOTIMP */
Public Const WSANO_RECOVERY% = (WSABASEERR + 1003)
Public Const NO_RECOVERY% = WSANO_RECOVERY

'/* Valid name, no data record of requested type */
Public Const WSANO_DATA% = (WSABASEERR + 1004)
Public Const NO_DATA% = WSANO_DATA

'/* no address, look for MX record */
Public Const WSANO_ADDRESS% = WSANO_DATA
Public Const NO_ADDRESS% = WSANO_ADDRESS

'/*
' * Windows Sockets errors redefined as regular Berkeley error constants.
' * These are commented out in Windows NT to avoid conflicts with errno.h.
' * Use the WSA constants instead.
' */
'       (I think these are here for backwards compatability reasons -ed)
'
#If 0 Then
  Public Const EWOULDBLOCK% = WSAEWOULDBLOCK
  Public Const EINPROGRESS% = WSAEINPROGRESS
  Public Const EALREADY% = WSAEALREADY
  Public Const ENOTSOCK% = WSAENOTSOCK
  Public Const EDESTADDRREQ% = WSAEDESTADDRREQ
  Public Const EMSGSIZE% = WSAEMSGSIZE
  Public Const EPROTOTYPE% = WSAEPROTOTYPE
  Public Const ENOPROTOOPT% = WSAENOPROTOOPT
  Public Const EPROTONOSUPPORT% = WSAEPROTONOSUPPORT
  Public Const ESOCKTNOSUPPORT% = WSAESOCKTNOSUPPORT
  Public Const EOPNOTSUPP% = WSAEOPNOTSUPP
  Public Const EPFNOSUPPORT% = WSAEPFNOSUPPORT
  Public Const EAFNOSUPPORT% = WSAEAFNOSUPPORT
  Public Const EADDRINUSE% = WSAEADDRINUSE
  Public Const EADDRNOTAVAIL% = WSAEADDRNOTAVAIL
  Public Const ENETDOWN% = WSAENETDOWN
  Public Const ENETUNREACH% = WSAENETUNREACH
  Public Const ENETRESET% = WSAENETRESET
  Public Const ECONNABORTED% = WSAECONNABORTED
  Public Const ECONNRESET% = WSAECONNRESET
  Public Const ENOBUFS% = WSAENOBUFS
  Public Const EISCONN% = WSAEISCONN
  Public Const ENOTCONN% = WSAENOTCONN
  Public Const ESHUTDOWN% = WSAESHUTDOWN
  Public Const ETOOMANYREFS% = WSAETOOMANYREFS
  Public Const ETIMEDOUT% = WSAETIMEDOUT
  Public Const ECONNREFUSED% = WSAECONNREFUSED
  Public Const ELOOP% = WSAELOOP
  Public Const ENAMETOOLONG% = WSAENAMETOOLONG
  Public Const EHOSTDOWN% = WSAEHOSTDOWN
  Public Const EHOSTUNREACH% = WSAEHOSTUNREACH
  Public Const ENOTEMPTY% = WSAENOTEMPTY
  Public Const EPROCLIM% = WSAEPROCLIM
  Public Const EUSERS% = WSAEUSERS
  Public Const EDQUOT% = WSAEDQUOT
  Public Const ESTALE% = WSAESTALE
  Public Const EREMOTE% = WSAEREMOTE
#End If '0

'/*
' * WinSock 2 extension -- new error codes and type definition
' */
#If Win32 And WinSock2 Then
  
  Type WSAEVENT
    event As Long
  End Type
  '#define WSAOVERLAPPED           OVERLAPPED
  'typedef struct _OVERLAPPED *    LPWSAOVERLAPPED;

  '#define WSA_IO_PENDING          (ERROR_IO_PENDING)
  '#define WSA_IO_INCOMPLETE       (ERROR_IO_INCOMPLETE)
  '#define WSA_INVALID_HANDLE      (ERROR_INVALID_HANDLE)
  '#define WSA_INVALID_PARAMETER   (ERROR_INVALID_PARAMETER)
  '#define WSA_NOT_ENOUGH_MEMORY   (ERROR_NOT_ENOUGH_MEMORY)
  '#define WSA_OPERATION_ABORTED   (ERROR_OPERATION_ABORTED)

  'public Const WSA_INVALID_EVENT& = Null
  '#define WSA_MAXIMUM_WAIT_EVENTS (MAXIMUM_WAIT_OBJECTS)
  Public Const WSA_WAIT_FAILED& = -1
  '#define WSA_WAIT_EVENT_0        (WAIT_OBJECT_0)
  '#define WSA_WAIT_IO_COMPLETION  (WAIT_IO_COMPLETION)
  '#define WSA_WAIT_TIMEOUT        (WAIT_TIMEOUT)
  '#define WSA_INFINITE            (INFINITE)

#ElseIf Win16 And WinSock2 Then

  '#define WSAAPI                  FAR PASCAL
  'typedef DWORD                   WSAEVENT, FAR * LPWSAEVENT;

  '} WSAOVERLAPPED, FAR * LPWSAOVERLAPPED;

  Public Const WSA_IO_PENDING = WSAEWOULDBLOCK
  Public Const WSA_IO_INCOMPLETE = WSAEWOULDBLOCK
  Public Const WSA_INVALID_HANDLE = WSAENOTSOCK
  Public Const WSA_INVALID_PARAMETER = WSAEINVAL
  Public Const WSA_NOT_ENOUGH_MEMORY = WSAENOBUFS
  Public Const WSA_OPERATION_ABORTED = WSAEINTR

  Public Const WSA_INVALID_EVENT = Null
  '#define WSA_MAXIMUM_WAIT_EVENTS (MAXIMUM_WAIT_OBJECTS)
  Public Const WSA_WAIT_FAILED& = -1
  Public Const WSA_WAIT_EVENT_0& = 0
  Public Const WSA_WAIT_TIMEOUT& = &H102&
  Public Const WSA_INFINITE& = -1
#End If
  

'/*
' *
' * WinSock 2 extensions
' *
' */
#If WinSock2 Then
  
  Type WSAOVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
  End Type

  '/*
  ' * WinSock 2 extension -- WSABUF and QOS struct
  ' */
  Type WSABUF
    len As Long          '/* the length of the buffer */
    buf As Long          '/* the pointer to the buffer */
  End Type

  Type GUARANTEE  'these might not be byte values!
    BestEffortService As Byte
    ControlledLoadService As Byte
    PredictiveService As Byte
    GuaranteedDelayService As Byte
    GuaranteedService As Byte
  End Type  ' defined as 'enum' type from 'winsock2.h'   What's that??

  Type flowspec
    TokenRate As Integer              '/* In Bytes/sec */
    TokenBucketSize As Integer        '/* In Bytes */
    PeakBandwidth As Integer          '/* In Bytes/sec */
    Latency As Integer                '/* In microseconds */
    DelayVariation As Integer         '/* In microseconds */
    LevelOfGuarantee As GUARANTEE     '/* Guaranteed, Predictive */
                                      '/*   or Best Effort       */
    CostOfCall As Integer             '/* Reserved for future use, */
                                      '/*   must be set to 0 now   */
    NetworkAvailability As Integer    '/* read-only:         */
                                      '/*   1 if accessible, */
                                      '/*   0 if not         */
  End Type

  Type QOS
      SendingFlowspec As flowspec        '/* the flow spec for data sending */
      ReceivingFlowspec As flowspec      '/* the flow spec for data receiving */
      ProviderSpecific As WSABUF         '/* additional provider specific stuff */
  End Type

  '/*
  ' * WinSock 2 extension -- manifest constants for return values of the condition function
  ' */
  Public Const CF_ACCEPT = &H0
  Public Const CF_REJECT = &H1
  Public Const CF_DEFER = &H2

  '/*
  ' * WinSock 2 extension -- manifest constants for shutdown()
  ' */
  Public Const SD_RECEIVE = &H0
  Public Const SD_SEND = &H1
  Public Const SD_BOTH = &H2

  '/*
  ' * WinSock 2 extension -- data type and manifest constants for socket groups
  ' */
  Type GROUP
    grp As Integer     'an unsigned integer
  End Type
  
  Public Const SG_UNCONSTRAINED_GROUP = &H1
  Public Const SG_CONSTRAINED_GROUP = &H2

  '/*
  ' * WinSock 2 extension -- data type for WSAEnumNetworkEvents()
  ' */
  Type WSANETWORKEVENTS
    lNetworkEvents As Long
    iErrorCode(FD_MAX_EVENTS) As Integer
  End Type

  '/*
  ' * WinSock 2 extension -- WSAPROTOCOL_INFO structure and associated
  ' * manifest constants
  ' */
  Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(8) As Byte
  End Type

  Public Const MAX_PROTOCOL_CHAIN = 7

  Public Const BASE_PROTOCOL = 1
  Public Const LAYERED_PROTOCOL = 0


  Type WSAPROTOCOLCHAIN
    ChainLen As Integer                           '/* the length of the chain,     */
                                                  '/* length = 0 means layered protocol, */
                                                  '/* length = 1 means base protocol, */
                                                  '/* length > 1 means protocol chain */
    ChainEntries(MAX_PROTOCOL_CHAIN) As Long      '/* a list of dwCatalogEntryIds */
  End Type


  Public Const WSAPROTOCOL_LEN = 255

  Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As GUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Integer
    iAddressFamily As Integer
    iMaxSockAddr As Integer
    iMinSockAddr As Integer
    iSocketType As Integer
    iProtocol As Integer
    iProtocolMaxOffset As Integer
    iNetworkByteOrder As Integer
    iSecurityScheme As Integer
    dwMessageSize As Long
    dwProviderReserved As Long
    #If UNICODE Then
      szProtocol(WSAPROTOCOL_LEN + 1) As Integer  'for the 16-bit unicode value
    #Else
      szProtocol(WSAPROTOCOL_LEN + 1) As Byte  'for the 8-bit ascii value
    #End If
  End Type

  '/* Flag bit definitions for dwProviderFlags */
  Public Const PFL_MULTIPLE_PROTO_ENTRIES& = &H1&
  Public Const PFL_RECOMMENDED_PROTO_ENTRY& = &H2&
  Public Const PFL_HIDDEN& = &H4&
  Public Const PFL_MATCHES_PROTOCOL_ZERO& = &H8&

  '/* Flag bit definitions for dwServiceFlags1 */
  Public Const XP1_CONNECTIONLESS& = &H1&
  Public Const XP1_GUARANTEED_DELIVERY& = &H2&
  Public Const XP1_GUARANTEED_ORDER& = &H4&
  Public Const XP1_MESSAGE_ORIENTED& = &H8&
  Public Const XP1_PSEUDO_STREAM& = &H10&
  Public Const XP1_GRACEFUL_CLOSE& = &H20&
  Public Const XP1_EXPEDITED_DATA& = &H40&
  Public Const XP1_CONNECT_DATA& = &H80&
  Public Const XP1_DISCONNECT_DATA& = &H100&
  Public Const XP1_SUPPORT_BROADCAST& = &H200&
  Public Const XP1_SUPPORT_MULTIPOINT& = &H400&
  Public Const XP1_MULTIPOINT_CONTROL_PLANE& = &H800&
  Public Const XP1_MULTIPOINT_DATA_PLANE& = &H1000&
  Public Const XP1_QOS_SUPPORTED& = &H2000&
  Public Const XP1_INTERRUPT& = &H4000&
  Public Const XP1_UNI_SEND& = &H8000&
  Public Const XP1_UNI_RECV& = &H10000
  Public Const XP1_IFS_HANDLES& = &H20000
  Public Const XP1_PARTIAL_MESSAGE& = &H40000

  Public Const BIGENDIAN% = &H0
  Public Const LITTLEENDIAN% = &H1

  Public Const SECURITY_PROTOCOL_NONE% = &H0

  '/*
  ' * WinSock 2 extension -- manifest constants for WSAJoinLeaf()
  ' */
  Public Const JL_SENDER_ONLY = &H1
  Public Const JL_RECEIVER_ONLY = &H2
  Public Const JL_BOTH = &H4

  '/*
  ' * WinSock 2 extension -- manifest constants for WSASocket()
  ' */
  Public Const WSA_FLAG_OVERLAPPED = &H1
  Public Const WSA_FLAG_MULTIPOINT_C_ROOT = &H2
  Public Const WSA_FLAG_MULTIPOINT_C_LEAF = &H4
  Public Const WSA_FLAG_MULTIPOINT_D_ROOT = &H8
  Public Const WSA_FLAG_MULTIPOINT_D_LEAF = &H10

  '/*
  ' * WinSock 2 extension -- manifest constants for WSAIoctl()
  ' */
  Public Const IOC_UNIX& = &H0&
  Public Const IOC_WS2& = &H8000000
  Public Const IOC_PROTOCOL& = &H10000000
  Public Const IOC_VENDOR& = &H18000000

  Public Const SIO_ASSOCIATE_HANDLE = (IOC_IN Or IOC_WS2 Or 1)
  Public Const SIO_ENABLE_CIRCULAR_QUEUEING = (IOC_VOID Or IOC_WS2 Or 2)
  Public Const SIO_FIND_ROUTE = (IOC_OUT Or IOC_WS2 Or 3)
  Public Const SIO_FLUSH = (IOC_VOID Or IOC_WS2 Or 4)
  Public Const SIO_GET_BROADCAST_ADDRESS = (IOC_VOID Or IOC_WS2 Or 5)
  Public Const SIO_GET_EXTENSION_FUNCTION_POINTER = (IOC_INOUT Or IOC_WS2 Or 6)
  Public Const SIO_GET_QOS = (IOC_INOUT Or IOC_WS2 Or 7)
  Public Const SIO_GET_GROUP_QOS = (IOC_INOUT Or IOC_WS2 Or 8)
  Public Const SIO_MULTIPOINT_LOOPBACK = (IOC_IN Or IOC_WS2 Or 9)
  Public Const SIO_MULTICAST_SCOPE = (IOC_IN Or IOC_WS2 Or 10)
  Public Const SIO_SET_QOS = (IOC_IN Or IOC_WS2 Or 11)
  Public Const SIO_SET_GROUP_QOS = (IOC_IN Or IOC_WS2 Or 12)
  Public Const SIO_TRANSLATE_HANDLE = (IOC_INOUT Or IOC_WS2 Or 13)

  '/*
  ' * WinSock 2 extension -- manifest constants for SIO_TRANSLATE_HANDLE ioctl
  ' */
  Public Const TH_NETDEV& = &H1
  Public Const TH_TAPI& = &H2

  '/*
  ' * Manifest constants and type definitions related to name resolution and
  ' * registration (RNR) API
  ' */
  Type BLOB
    cbSize As Long
    #If MIDL_PASS Then
      pBlobData() As Byte   'I guess.... use "redim pBlobData(cbSize)" when needed
    #Else
      pBlobData As Byte
    #End If
  End Type
  
  '/*
  ' * Service Install Flags
  ' */
  Public Const SERVICE_MULTIPLE& = &H1
  
  '/*
  ' *& Name Spaces
  ' */
  Public Const NS_ALL = (0)

  Public Const NS_SAP = (1)
  Public Const NS_NDS = (2)
  Public Const NS_PEER_BROWSE = (3)

  Public Const NS_TCPIP_LOCAL = (10)
  Public Const NS_TCPIP_HOSTS = (11)
  Public Const NS_DNS = (12)
  Public Const NS_NETBT = (13)
  Public Const NS_WINS = (14)

  Public Const NS_NBP = (20)

  Public Const NS_MS = (30)
  Public Const NS_STDA = (31)
  Public Const NS_NTDS = (32)

  Public Const NS_X500 = (40)
  Public Const NS_NIS = (41)
  Public Const NS_NISPLUS = (42)

  Public Const NS_WRQ = (50)

  '/*
  ' * Resolution flags for WSAGetAddressByName().
  ' * Note these are also used by the 1.1 API GetAddressByName, so
  ' * leave them around.
  ' */
  Public Const RES_UNUSED_1& = (&H1)
  Public Const RES_FLUSH_CACHE& = (&H2)
  Public Const RES_SERVICE& = (&H4)

  Public Const SERVICE_TYPE_VALUE_SAPID$ = "SapId"
  Public Const SERVICE_TYPE_VALUE_TCPPORT$ = "TcpPort"
  Public Const SERVICE_TYPE_VALUE_UDPPORT$ = "UdpPort"
  Public Const SERVICE_TYPE_VALUE_OBJECTID$ = "ObjectId"
  
  
  #If CSADDR_DEFINED Then
    
    '/*
    ' * SockAddr Information
    ' */
    Type SOCKET_ADDRESS
      lpSockaddr As Long            '(pointer) to SOCKADDR structure (?)
      iSockaddrLength As Integer
    End Type
    '} SOCKET_ADDRESS, *PSOCKET_ADDRESS, FAR * LPSOCKET_ADDRESS ;
  
    '/*
    ' * CSAddr Information
    ' */
    Type CSADDR_INFO
      LocalAddr As SOCKET_ADDRESS
      RemoteAddr As SOCKET_ADDRESS
      iSocketType As Integer
      iProtocol As Integer
    End Type
    '} CSADDR_INFO, *PCSADDR_INFO, FAR * LPCSADDR_INFO ;
  
  #End If
  
  '/*
  ' *  Address Family/Protocol Tuples
  ' */
  Type AFPROTOCOLS
    iAddressFamily As Integer
    iProtocol As Integer
  End Type
  '} AFPROTOCOLS, *PAFPROTOCOLS, *LPAFPROTOCOLS;

  '/*
  ' * Client Query API Typedefs
  ' */

  '/*
  ' * The comparators
  ' */
  '
  '  this type is not correct! Can't stuff values into
  '  a user-defined type and declare it as a constant.
  Type WSAEcomparator
    COMP_EQUAL As Byte    'this value is to be zero
    COMP_NOTLESS As Byte  'this byte(?) contains what COMP_NOTLESS represents
  End Type
  '} WSAECOMPARATOR, *PWSAECOMPARATOR, *LPWSAECOMPARATOR;

  Type WSAVersion
    dwVersion As Long
    ecHow As WSAEcomparator
  End Type
  '}WSAVERSION, *PWSAVERSION, *LPWSAVERSION;

  Type WSAQuerySet
    dwSize As Long
    lpszServiceInstanceName As String
    lpServiceClassId As GUID
    lpVersion As WSAVersion
    lpszComment As String
    dwNameSpace As Long
    lpNSProviderId As GUID
    lpszContext As String
    dwNumberOfProtocols As Long
    lpafpProtocols As AFPROTOCOLS
    lpszQueryString As String
    dwNumberOfCsAddrs As Long
    lpcsaBuffer As CSADDR_INFO
    dwOutputFlags As Long
    lpBlob As BLOB
  End Type
  '} WSAQUERYSETW, *PWSAQUERYSETW, *LPWSAQUERYSETW;

  Public Const LUP_DEEP% = &H1
  Public Const LUP_CONTAINERS% = &H2
  Public Const LUP_NOCONTAINERS% = &H4
  Public Const LUP_NEAREST% = &H8
  Public Const LUP_RETURN_NAME% = &H10
  Public Const LUP_RETURN_TYPE% = &H20
  Public Const LUP_RETURN_VERSION% = &H40
  Public Const LUP_RETURN_COMMENT% = &H80
  Public Const LUP_RETURN_ADDR% = &H100
  Public Const LUP_RETURN_BLOB% = &H200
  Public Const LUP_RETURN_ALIASES% = &H400
  Public Const LUP_RETURN_QUERY_STRING% = &H800
  Public Const LUP_RETURN_ALL% = &HFF0
  Public Const LUP_RES_SERVICE% = &H8000

  Public Const LUP_FLUSHCACHE% = &H1000
  Public Const LUP_FLUSHPREVIOUS% = &H2000

  '//
  '// Return flags
  '//
  Public Const RESULT_IS_ALIAS% = &H1

  '/*
  ' * Service Address Registration and Deregistration Data Types.
  ' */

  'typedef enum _WSAESETSERVICEOP
  '{
  '    RNRSERVICE_REGISTER=0,
  '    RNRSERVICE_DEREGISTER,
  '    RNRSERVICE_DELETE
  '} WSAESETSERVICEOP, *PWSAESETSERVICEOP, *LPWSAESETSERVICEOP;

  '/*
  ' * Service Installation/Removal Data Types.
  ' */
  Type WSANSClassInfo
    lpszName As String
    dwNameSpace As Long
    dwValueType As Long
    dwValueSize As Long
    'lpValue as LPVOID??
  End Type
  '}WSANSCLASSINFOA, *PWSANSCLASSINFOA, *LPWSANSCLASSINFOA;

  Type WSAServiceClassInfo
    lpServiceClassId As GUID
    lpszServiceClassName As String
    dwCount As Long
    lpClassInfos As WSANSClassInfo
  End Type
  '}WSASERVICECLASSINFOW, *PWSASERVICECLASSINFOW, *LPWSASERVICECLASSINFOW;

  Type WSANAMESPACE_INFO
    NSProviderId As GUID
    dwNameSpace As Long
    fActive As Boolean
    dwVersion As Long
    lpszIdentifier As String
  End Type
  '} WSANAMESPACE_INFOW, *PWSANAMESPACE_INFOW, *LPWSANAMESPACE_INFOW;

  '/*
  ' * WinSock 2 extensions -- data types for the condition function in
  ' * WSAAccept() and overlapped I/O completion routine.
  ' */

  'typedef
  'int
  '(CALLBACK * LPCONDITIONPROC)(
  '  LPWSABUF lpCallerId,           long
  '  LPWSABUF lpCallerData,         long
  '  LPQOS lpSQOS,                  long
  '  LPQOS lpGQOS,                  long
  '  LPWSABUF lpCalleeId,           long
  '  LPWSABUF lpCalleeData,         long
  '  GROUP FAR * g,
  '  DWORD dwCallbackData
  '  );

Type WSAOVERLAPPED_COMPLETION_ROUTINE
  dwError As Long
  cbTransferred As Long
  lpOverlapped As WSAOVERLAPPED
  dwFlags As Long
End Type

#End If 'WinSock2


'/*
' *
' * Socket function prototypes
' *
' */


'/*
' * Winsock 1.1 API function prototypes for Win32
' */
'############################################################################
#If Win32 And WinSock1 Then
  Declare Function accept& Lib "wsock32.dll" Alias "#1" (ByVal s&, ByRef Addr As sockaddr, ByRef namelen%)
   ' acceptIn() uses the sockaddr_in structure instead of the plain sockaddr for Internet style addresses
   ' Although one could take interest in a per-byte basis with the plain sockaddr, it seems easier to
   ' just use the call below.
  Declare Function acceptIn& Lib "wsock32.dll" Alias "#1" (ByVal s&, ByRef Addr As sockaddr_in, ByRef namelen%)
   ' use acceptNull() when you don't need an address structure returned
   ' code sample:  di = acceptNull(SocketDesc, vbNullString, vbNullString)
  Declare Function acceptNull& Lib "wsock32.dll" Alias "#1" (ByVal s&, ByVal sNull$, ByVal sNull$)
  Declare Function bind% Lib "wsock32.dll" Alias "#2" (ByVal s&, ByRef Addr As sockaddr, ByVal namelen%)
  Declare Function bindIn% Lib "wsock32.dll" Alias "#2" (ByVal s&, ByRef Addr As sockaddr_in, ByVal namelen%)
  Declare Function closesocket% Lib "wsock32.dll" Alias "#3" (ByVal s&)
  Declare Function Connect% Lib "wsock32.dll" Alias "#4" (ByVal s&, ByRef Addr As sockaddr, ByVal namelen%)
  Declare Function connectIn% Lib "wsock32.dll" Alias "#4" (ByVal s&, ByRef Addr As sockaddr_in, ByVal namelen%)
  Declare Function ioctlsocket% Lib "wsock32.dll" (ByVal s&, ByVal cmd&, ByRef argp&)
  Declare Function getpeername% Lib "wsock32.dll" Alias "#5" (ByVal s&, ByRef peername As sockaddr, ByRef namelen%)
  Declare Function getpeernameIn% Lib "wsock32.dll" Alias "#5" (ByVal s&, ByRef peername As sockaddr_in, ByRef namelen%)
  Declare Function getsockname% Lib "wsock32.dll" Alias "#6" (ByVal s&, ByRef sockname As sockaddr, ByRef namelen%)
  Declare Function getsocknameIn% Lib "wsock32.dll" Alias "#6" (ByVal s&, ByRef sockname As sockaddr_in, ByRef namelen%)
  Declare Function getsockopt% Lib "wsock32.dll" Alias "#7" (ByVal s&, ByVal level%, ByVal optname%, ByRef optval&, ByRef optlen%)
  Declare Function htonl& Lib "wsock32.dll" Alias "#8" (ByVal hostlong&)
  Declare Function htons% Lib "wsock32.dll" Alias "#9" (ByVal hostshort%)
  Declare Function inet_addr& Lib "wsock32.dll" Alias "#10" (ByVal cp$)
  Declare Function inet_ntoa& Lib "wsock32.dll" Alias "#11" (ByVal inet&)
  Declare Function listen% Lib "wsock32.dll" (ByVal s As Long, ByVal backlog%)
  Declare Function ntohl& Lib "wsock32.dll" Alias "#14" (ByVal netlong&)
  Declare Function ntohs% Lib "wsock32.dll" Alias "#15" (ByVal netshort%)
  Declare Function recv% Lib "wsock32.dll" Alias "#16" (ByVal s&, ByVal buf&, ByVal buflen%, ByVal flags%)
  Declare Function recvfrom% Lib "wsock32.dll" Alias "#17" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, fromaddr As sockaddr, fromlen%)
  Declare Function recvfromIn% Lib "wsock32.dll" Alias "#17" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, fromaddr As sockaddr_in, fromlen%)
  ' Visual Basic note...since select is a keyword in Visual Basic the function
  ' has been renamed
  Declare Function WSASelect% Lib "wsock32.dll" Alias "#18" (ByVal nfds%, ByRef readfds As FD_SET, ByRef writefds As FD_SET, ByRef exceptfds As FD_SET, ByVal TimeOut As timeval)
  Declare Function send% Lib "wsock32.dll" Alias "#19" (ByVal s&, ByRef buf$, ByVal buflen%, ByVal flags%)
  Declare Function sendto% Lib "wsock32.dll" Alias "#20" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef toaddr As sockaddr, ByRef tolen%)
  Declare Function sendtoIn% Lib "wsock32.dll" Alias "#20" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef toaddr As sockaddr_in, ByRef tolen%)
  Declare Function setsockopt% Lib "wsock32.dll" Alias "#21" (ByVal s&, ByVal level%, ByVal optname%, ByRef optval&, ByVal optlen%)
  Declare Function shutdown% Lib "wsock32.dll" Alias "#22" (ByVal s&, ByVal how%)
  Declare Function Socket& Lib "wsock32.dll" Alias "#23" (ByVal af%, ByVal socktype%, ByVal protocol%)
  Declare Function gethostbyaddr& Lib "wsock32.dll" Alias "#51" (ByRef Addr&, ByVal addrlen%, ByVal addrtype%)
  Declare Function GetHostByName& Lib "wsock32.dll" Alias "#52" (ByVal HostName$)
  Declare Function GetHostName% Lib "wsock32.dll" Alias "#57" (ByVal HostName$, ByVal namelen%)
  Declare Function getservbyport& Lib "wsock32.dll" Alias "#56" (ByVal Port%, ByVal protoname&)
  Declare Function getservbyname& Lib "wsock32.dll" Alias "#55" (ByVal servname&, ByVal protoname&)
  Declare Function getprotobynumber& Lib "wsock32.dll" Alias "#54" (ByVal protonumber%)
  Declare Function getprotobyname& Lib "wsock32.dll" Alias "#53" (ByVal protoname$)
  Declare Function WSAStartup% Lib "wsock32.dll" Alias "#115" (ByVal wVersionRequired%, lpWSAData As WSAData)
  Declare Function WSACleanUp% Lib "wsock32.dll" Alias "#116" ()
  Declare Function WSASetLastError% Lib "wsock32.dll" Alias "#112" (ByVal iError%)
  Declare Function WSAGetLastError% Lib "wsock32.dll" Alias "#111" ()
  Declare Function WSAIsBlocking% Lib "wsock32.dll" Alias "#114" ()
  Declare Function WSAUnhookBlockingHook% Lib "wsock32.dll" Alias "#110" ()
  Declare Function WSASetBlockingHook& Lib "wsock32.dll" Alias "#109" (lpFunc&)
  Declare Function WSACancelBlockingCall% Lib "wsock32.dll" Alias "#113" ()
  Declare Function WSAAsyncGetServByName% Lib "wsock32.dll" Alias "#107" (ByVal hWnd&, ByVal wMsg%, ByVal HostName$, ByVal proto$, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetServByPort% Lib "wsock32.dll" Alias "#106" (ByVal hWnd&, ByVal wMsg%, ByVal Port%, ByVal proto$, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetProtoByName% Lib "wsock32.dll" Alias "#105" (ByVal hWnd&, ByVal wMsg%, ByVal protoname$, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetProtoByNumber% Lib "wsock32.dll" Alias "#104" (ByVal hWnd&, ByVal wMsg%, ByVal number%, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetHostByName% Lib "wsock32.dll" Alias "#103" (ByVal hWnd&, ByVal wMsg%, ByVal HostName$, ByRef buf As BIGhostent, ByVal buflen%)
  Declare Function WSAAsyncGetHostByAddr% Lib "wsock32.dll" Alias "#102" (ByVal hWnd&, ByVal wMsg%, ByVal Addr$, ByVal addrlen%, ByVal addrtype%, ByRef buf As BIGhostent, ByVal buflen%)
  Declare Function WSACancelAsyncRequest% Lib "wsock32.dll" Alias "#108" (ByVal hAsyncTaskHandle%)
  Declare Function WSAAsyncSelect% Lib "wsock32.dll" Alias "#101" (ByVal s&, ByVal hWnd&, ByVal wMsg%, ByVal lEvent&)
#End If

'/*
' * WinSock 2.2 API function prototypes for Win32
' */
'############################################################################
#If Win32 And WinSock2 Then
  Declare Function accept2& Lib "ws2_32.dll" Alias "#1" (ByVal s&, ByRef Addr As sockaddr, namelen%)
   ' acceptIn() uses the sockaddr_in structure instead of the plain sockaddr for Internet style addresses
   ' although one could take interest in a per-byte basis with the plain sockaddr, it seems easier to
   ' just use the call below.
  Declare Function acceptIn2& Lib "ws2_32.dll" Alias "#1" (ByVal s&, ByRef Addr As sockaddr_in, namelen%)
   ' use acceptNull() when you don't need an address structure returned
   ' code sample:  di = acceptNull(SocketDesc, vbNullString, vbNullString)
  Declare Function acceptNull2& Lib "ws2_32.dll" Alias "#1" (ByVal s&, ByVal sNull$, ByVal sNull$)
  Declare Function bind2% Lib "ws2_32.dll" Alias "#2" (ByVal s&, ByRef Addr As sockaddr, ByVal namelen%)
  Declare Function bindIn2% Lib "ws2_32.dll" Alias "#2" (ByVal s&, ByRef Addr As sockaddr_in, ByVal namelen%)
  Declare Function closesocket2% Lib "ws2_32.dll" Alias "#3" (ByVal s&)
  Declare Function connect2% Lib "ws2_32.dll" Alias "#4" (ByVal s&, ByRef Addr As sockaddr, ByVal namelen%)
  Declare Function connectIn2% Lib "ws2_32.dll" Alias "#4" (ByVal s&, ByRef Addr As sockaddr_in, ByVal namelen%)
  Declare Function ioctlsocket2% Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s&, ByVal cmd&, ByRef argp&)
  Declare Function getpeername2% Lib "ws2_32.dll" Alias "#5" (ByVal s&, ByRef peername As sockaddr, namelen%)
  Declare Function getpeernameIn2% Lib "ws2_32.dll" Alias "#5" (ByVal s&, ByRef peername As sockaddr_in, namelen%)
  Declare Function getsockname2% Lib "ws2_32.dll" Alias "#6" (ByVal s&, ByRef sockname As sockaddr, namelen%)
  Declare Function getsocknameIn2% Lib "ws2_32.dll" Alias "#6" (ByVal s&, ByRef sockname As sockaddr_in, namelen%)
  Declare Function getsockopt2% Lib "ws2_32.dll" Alias "#7" (ByVal s&, ByVal level%, ByVal optname%, ByVal optval$, optlen%)
  Declare Function htonl2& Lib "ws2_32.dll" Alias "#8" (ByVal hostlong&)
  Declare Function htons2% Lib "ws2_32.dll" Alias "#9" (ByVal hostshort%)
  Declare Function inet_addr2& Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp$) '"#10"
  Declare Function inet_ntoa2& Lib "ws2_32.dll" Alias "inet_ntoa" (ByVal inet&) '"#11"
  Declare Function listen2% Lib "ws2_32.dll" Alias "listen" (ByVal s&, ByVal backlog&)
  Declare Function ntohl2& Lib "ws2_32.dll" Alias "#14" (ByVal netlong&)
  Declare Function ntohs2% Lib "ws2_32.dll" Alias "#15" (ByVal netshort%)
  Declare Function recv2% Lib "ws2_32.dll" Alias "recv" (ByVal s&, ByRef buf As Byte, ByVal buflen%, ByVal flags%)
  Declare Function recvfrom2% Lib "ws2_32.dll" Alias "#17" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef fromaddr As sockaddr, fromlen%)
  Declare Function recvfromIn2% Lib "ws2_32.dll" Alias "#17" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef fromaddr As sockaddr_in, fromlen%)
  ' Visual Basic note...since select is a keyword in Visual Basic the function
  ' has been renamed
  Declare Function WSASelect2% Lib "ws2_32.dll" Alias "#18" (ByVal nfds%, ByRef readfds As FD_SET, ByRef writefds As FD_SET, ByRef exceptfds As FD_SET, ByRef TimeOut As timeval)
  Declare Function send2% Lib "ws2_32.dll" Alias "#19" (ByVal s&, ByVal buf$, ByVal buflen%, ByVal flags%)
  Declare Function sendto2% Lib "ws2_32.dll" Alias "#20" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef toaddr As sockaddr, ByVal tolen%)
  Declare Function sendtoIn2% Lib "ws2_32.dll" Alias "#20" (ByVal s&, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef toaddr As sockaddr_in, ByVal tolen%)
  Declare Function setsockopt2% Lib "ws2_32.dll" Alias "#21" (ByVal s&, ByVal level%, ByVal optname%, optval&, ByVal optlen%)
  Declare Function shutdown2% Lib "ws2_32.dll" Alias "#22" (ByVal s&, ByVal how%)
  Declare Function socket2& Lib "ws2_32.dll" Alias "#23" (ByVal af%, ByVal socktype%, ByVal protocol%)
  Declare Function gethostbyaddr2& Lib "ws2_32.dll" Alias "#51" (ByRef Addr&, ByVal addrlen%, ByVal addrtype%)
  Declare Function GetHostByName2& Lib "ws2_32.dll" Alias "#52" (ByVal HostName$)
  Declare Function GetHostName2% Lib "ws2_32.dll" Alias "#57" (ByVal HostName$, ByVal namelen%)
  Declare Function getservbyport2& Lib "ws2_32.dll" Alias "#56" (ByVal Port%, ByVal protoname&)
  Declare Function getservbyname2& Lib "ws2_32.dll" Alias "#55" (ByVal servname&, ByVal protoname&)
  Declare Function getprotobynumber2& Lib "ws2_32.dll" Alias "#54" (ByVal protonumber%)
  Declare Function getprotobyname2& Lib "ws2_32.dll" Alias "#53" (ByVal protoname$)
  Declare Function WSACleanUp2% Lib "ws2_32.dll" Alias "#116" ()
  Declare Sub WSASetLastError2 Lib "ws2_32.dll" Alias "#112" (ByVal iError%)
  Declare Function WSAGetLastError2% Lib "ws2_32.dll" Alias "#111" ()
  Declare Function WSAAsyncGetServByName2% Lib "ws2_32.dll" Alias "#107" (ByVal hWnd&, ByVal wMsg%, ByVal HostName$, ByVal proto$, ByRef buf As servent, ByVal buflen%)
  Declare Function WSAAsyncGetServByPort2% Lib "ws2_32.dll" Alias "#106" (ByVal hWnd&, ByVal wMsg%, ByVal Port%, ByVal proto$, ByRef buf As servent, ByVal buflen%)
  Declare Function WSAAsyncGetProtoByName2% Lib "ws2_32.dll" Alias "#105" (ByVal hWnd&, ByVal wMsg%, ByVal protoname$, ByRef buf As protoent, ByVal buflen%)
  Declare Function WSAAsyncGetProtoByNumber2% Lib "ws2_32.dll" Alias "#104" (ByVal hWnd&, ByVal wMsg%, ByVal number%, ByRef buf As hostent, ByVal buflen%)
  Declare Function WSAAsyncGetHostByName2% Lib "ws2_32.dll" Alias "#103" (ByVal hWnd&, ByVal wMsg%, ByVal HostName$, ByRef buf As BIGhostent, ByVal buflen%)
  Declare Function WSAAsyncGetHostByAddr2% Lib "ws2_32.dll" Alias "#102" (ByVal hWnd&, ByVal wMsg%, ByVal Addr$, ByVal addrlen%, ByVal addrtype%, ByRef buf As BIGhostent, ByVal buflen%)
  Declare Function WSACancelAsyncRequest2% Lib "ws2_32.dll" Alias "#108" (ByVal hAsyncTaskHandle&)
  Declare Function WSAAsyncSelect2% Lib "ws2_32.dll" Alias "#101" (ByVal s&, ByVal hWnd&, ByVal wMsg%, ByVal lEvent&)
  ' conditionproc constant needs to be finished for this
  'Declare Function WSAAccept& Lib "ws2_32.dll" (ByVal s&, ByRef addr As sockaddr, ByRef namelen%, ByRef lpfnCondition As CONDITIONPROC, ByRef dwCallbackData&)
  Declare Function WSACloseEvent Lib "ws2_32.dll" (ByVal hEvent As WSAEVENT) As Boolean
  Declare Function WSAConnect% Lib "ws2_32.dll" (ByVal s&, ByRef name As sockaddr, ByRef namelen%, ByRef lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByRef lpSQOS As QOS, ByRef lpGQOS As QOS)
  Declare Function WSAConnectIn% Lib "ws2_32.dll" (ByVal s&, ByRef name As sockaddr_in, ByRef namelen%, ByRef lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByRef lpSQOS As QOS, ByRef lpGQOS As QOS)
  Declare Function WSACreateEvent& Lib "ws2_32.dll" ()
  #If UNICODE Then
    Declare Function WSADuplicateSocket Lib "ws2_32.dll" Alias "WSADuplicateSocketW" (ByVal s&, ByVal dwProcessId&, ByRef lpProtocolInfo As WSAPROTOCOL_INFO)
  #Else
    Declare Function WSADuplicateSocket Lib "ws2_32.dll" Alias "WSADuplicateSocketA" (ByVal s&, ByVal dwProcessId&, ByRef lpProtocolInfo As WSAPROTOCOL_INFO)
  #End If
  Declare Function WSAEnumNetworkEvents% Lib "ws2_32.dll" (ByVal s&, ByVal hEventObject&, ByRef lpNetworkEvents As WSANETWORKEVENTS)
  #If UNICODE Then
    Declare Function WSAEnumProtocols% Lib "ws2_32.dll" Alias "WSAEnumProtocolsW" (ByVal lpiProtocols%, ByRef lpProtocolBuffer As WSAPROTOCOL_INFO, ByRef lpdwBufferLength&)
  #Else
    Declare Function WSAEnumProtocols% Lib "ws2_32.dll" Alias "WSAEnumProtocolsA" (ByVal lpiProtocols%, ByRef lpProtocolBuffer As WSAPROTOCOL_INFO, ByRef lpdwBufferLength&)
  #End If
  Declare Function WSAEventSelect% Lib "ws2_32.dll" (ByVal s&, ByVal hEventObject&, ByVal lNetworkEvents&)
  Declare Function WSAGetOverlappedResult Lib "ws2_32.dll" (ByVal s&, ByVal lpOverlapped As WSAOVERLAPPED, ByRef lpcbTransfer&, ByVal fWait As Boolean, ByRef lpdwFlags&) As Boolean
  Declare Function WSAGetQOSByName Lib "ws2_32.dll" (ByVal s&, ByVal lpQOSName As WSABUF, ByRef lpQOS As QOS) As Boolean
  Declare Function WSAHtonl% Lib "ws2_32.dll" (ByVal s&, ByVal hostlong&, ByRef lpnetlong&)
  Declare Function WSAHtons% Lib "ws2_32.dll" (ByVal s&, ByVal hostshort%, ByRef lpnetshort%)
  Declare Function WSAIoctl% Lib "ws2_32.dll" (ByVal s&, ByVal dwIoControlCode&, ByRef lpvInBuffer&, ByVal cbInBuffer, ByRef lpvOutBuffer, ByVal cbOutBuffer&, ByRef lpcbBytesReturned&, ByVal lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE)
  Declare Function WSAJoinLeaf& Lib "ws2_32.dll" (ByVal s&, ByVal name As sockaddr, ByVal namelen%, ByVal lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByVal lpSQOS As QOS, ByVal lpGQOS As QOS, ByVal dwFlags&)
  Declare Function WSAJoinLeafIn& Lib "ws2_32.dll" (ByVal s&, ByVal name As sockaddr_in, ByVal namelen%, ByVal lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByVal lpSQOS As QOS, ByVal lpGQOS As QOS, ByVal dwFlags&)
  Declare Function WSANtohl% Lib "ws2_32.dll" (ByVal s&, ByVal netlong&, ByRef lphostlong&)
  Declare Function WSANtohs% Lib "ws2_32.dll" (ByVal s&, ByVal netshort%, ByRef lphostshort%)
  Declare Function WSARecv% Lib "ws2_32.dll" (ByVal s&, ByRef lpBuffers As WSABUF, ByVal dwBufferCount&, ByRef lpNumberOfBytesRecvd&, ByRef lpFlags&, ByVal lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE)
  Declare Function WSARecvDisconnect% Lib "ws2_32.dll" (ByVal s&, ByRef lpInboundDisconnectData As WSABUF)
  Declare Function WSARecvFrom% Lib "ws2_32.dll" (ByVal s&, ByRef lpBuffers As WSABUF, ByVal dwBufferCount&, ByRef lpNumberOfBytesRecvd&, ByRef lpFlags&, ByRef lpFrom As sockaddr, ByRef lpFromlen%, ByVal lpOverlapped%, ByVal lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE)
  Declare Function WSAResetEvent Lib "ws2_32.dll" (hEvent As WSAEVENT) As Boolean
  Declare Function WSASend% Lib "ws2_32.dll" (ByVal s&, ByRef lpBuffers As WSABUF, ByVal dwBufferCount&, ByRef lpNumberOfBytesRecvd&, ByVal dwFlags&, ByVal lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE)
  Declare Function WSASendDisconnect% Lib "ws2_32.dll" (ByVal s&, ByVal lpOutboundDisconnectData As WSABUF)
  Declare Function WSASendTo% Lib "ws2_32.dll" (ByVal s&, ByRef lpBuffers As WSABUF, ByVal dwBufferCount&, ByRef lpNumberOfBytesSent&, ByVal dwFlags&, ByVal lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE)
  Declare Function WSASetEvent Lib "ws2_32.dll" (ByVal hEvent As WSAEVENT) As Boolean
  #If UNICODE Then
    Declare Function WSASocket& Lib "ws2_32.dll" Alias "WSASocketW" (ByVal af%, ByVal socktype%, ByVal protocol%, ByVal lpProtocolInfo As WSAPROTOCOL_INFO, ByVal g As GROUP, ByVal dwFlags&)
  #Else
    Declare Function WSASocket& Lib "ws2_32.dll" Alias "WSASocketA" (ByVal af%, ByVal socktype%, ByVal protocol%, ByVal lpProtocolInfo As WSAPROTOCOL_INFO, ByVal g As GROUP, ByVal dwFlags&)
  #End If
  Declare Function WSAStartup2% Lib "ws2_32.dll" Alias "#115" (ByVal wVersionRequired%, ByRef lpWSAData As WSAData)
  Declare Function WSAWaitForMultipleEvents& Lib "ws2_32.dll" (ByVal cEvents&, ByRef lphEvents As WSAEVENT, ByVal fWaitAll As Boolean, ByVal dwTimeout&, ByVal fAlertable As Boolean)
  #If UNICODE Then
    Declare Function WSAAddressToString% Lib "ws2_32.dll" Alias "WSAAddressToStringW" (ByRef lpsaAddress As sockaddr, ByVal dwAddressLength&, ByVal lpProtocolInfo As WSAPROTOCOL_INFO, ByVal lpszAddressString$, ByVal lpdwAddressStringLength&)
    Declare Function WSAStringToAddress% Lib "ws2_32.dll" Alias "WSAStringToAddressW" (ByVal AddressString$, ByVal AddressFamily%, ByVal AddressFamily%, ByRef lpAddress As sockaddr, ByRef lpAddressLength%)
    Declare Function WSAEnumNameSpaceProviders% Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersW" (ByRef lpdwBufferLength&, ByRef lpnspBuffer As WSANAMESPACE_INFO)
    Declare Function WSAGetServiceClassInfo Lib "ws2_32.dll" Alias "WSAGetServiceClassInfoW" (ByVal lpProviderId As GUID, ByVal lpServiceClassId As GUID, ByRef lpdwBufSize&, ByRef lpServiceClassInfo As WSAServiceClassInfo)
    Declare Function WSAGetServiceClassNameByClassId% Lib "ws2_32.dll" Alias "WSAGetServiceClassNameByClassIdW" (ByVal lpServiceClassId As GUID, ByRef lpszServiceClassName$, ByRef lpdwBufferLength&)
    Declare Function WSAInstallServiceClass% Lib "ws2_32.dll" Alias "WSAInstallServiceClassW" (ByVal lpServiceClassInfo As WSAServiceClassInfo)
    Declare Function WSALookupServiceBegin% Lib "ws2_32.dll" Alias "WSALookupServiceBeginW" (ByVal lpqsRestrictions As WSAQuerySet, ByVal dwControlFlags&, ByRef lphLookup As Long)
    Declare Function WSALookupServiceNext% Lib "ws2_32.dll" Alias "WSALookupServiceNextW" (ByVal hLookup As Long, ByVal dwControlFlags&, ByRef lpdwBufferLength&, ByRef lpqsResults As WSAQuerySet)
  #Else
    Declare Function WSAAddressToString% Lib "ws2_32.dll" Alias "WSAAddressToStringA" (ByRef lpsaAddress As sockaddr, ByVal dwAddressLength&, ByVal lpProtocolInfo As WSAPROTOCOL_INFO, ByVal lpszAddressString$, ByVal lpdwAddressStringLength&)
    Declare Function WSAStringToAddress% Lib "ws2_32.dll" Alias "WSAStringToAddressA" (ByVal AddressString$, ByVal AddressFamily%, ByVal AddressFamily%, ByRef lpAddress As sockaddr, ByRef lpAddressLength%)
    Declare Function WSAEnumNameSpaceProviders% Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersA" (ByRef lpdwBufferLength&, ByRef lpnspBuffer As WSANAMESPACE_INFO)
    Declare Function WSAGetServiceClassInfo Lib "ws2_32.dll" Alias "WSAGetServiceClassInfoA" (ByVal lpProviderId As GUID, ByVal lpServiceClassId As GUID, ByRef lpdwBufSize&, ByRef lpServiceClassInfo As WSAServiceClassInfo)
    Declare Function WSAGetServiceClassNameByClassId% Lib "ws2_32.dll" Alias "WSAGetServiceClassNameByClassIdA" (ByVal lpServiceClassId As GUID, ByRef lpszServiceClassName$, ByRef lpdwBufferLength&)
    Declare Function WSAInstallServiceClass% Lib "ws2_32.dll" Alias "WSAInstallServiceClassA" (ByVal lpServiceClassInfo As WSAServiceClassInfo)
    Declare Function WSALookupServiceBegin% Lib "ws2_32.dll" Alias "WSALookupServiceBeginA" (ByVal lpqsRestrictions As WSAQuerySet, ByVal dwControlFlags&, ByRef lphLookup As Long)
    Declare Function WSALookupServiceNext% Lib "ws2_32.dll" Alias "WSALookupServiceNextA" (ByVal hLookup As Long, ByVal dwControlFlags&, ByRef lpdwBufferLength&, ByRef lpqsResults As WSAQuerySet)
  #End If
  Declare Function WSALookupServiceEnd% Lib "ws2_32.dll" (ByVal hLookup As Long)
  Declare Function WSARemoveServiceClass% Lib "ws2_32.dll" (ByVal lpServiceClassId As GUID)
  ' needs WSASETSERVICEOP constant for below:
  'Declare Function WSASetService% Lib "ws2_32.dll" Alias "WSASetServiceA" (ByVal lpqsRegInfo As WSAQuerySet, ByVal essOperation As WSASETSERVICEOP, ByVal dwControlFlags&)
#End If

'/*
' * Winsock 1.1 API function prototypes for Win16
' */
#If Win16 Then
'############################################################################
  Declare Function accept% Lib "winsock.dll" Alias "#1" (ByVal s%, ByRef Addr As sockaddr, ByRef namelen%)
   ' acceptIn() uses the sockaddr_in structure instead of the plain sockaddr for Internet style addresses
   ' although one could take interest in a per-byte basis with the plain sockaddr, it seems easier to
   ' just use the call below.
  Declare Function acceptIn% Lib "winsock.dll" Alias "#1" (ByVal s%, ByRef Addr As sockaddr_in, ByRef namelen%)
   ' use acceptNull() when you don't need an address structure returned
   ' code sample:  di = acceptNull(SocketDesc, vbNullString, vbNullString)
  Declare Function acceptNull% Lib "winsock.dll" Alias "#1" (ByVal s%, ByVal sNull$, ByVal sNull$)
  Declare Function bind% Lib "winsock.dll" Alias "#2" (ByVal s%, ByRef Addr As sockaddr, ByVal namelen%)
  Declare Function bindIn% Lib "winsock.dll" Alias "#2" (ByVal s%, ByRef Addr As sockaddr_in, ByVal namelen%)
  Declare Function closesocket% Lib "winsock.dll" Alias "#3" (ByVal s%)
  Declare Function Connect% Lib "winsock.dll" Alias "#4" (ByVal s%, ByRef Addr As sockaddr, ByVal namelen%)
  Declare Function connectIn% Lib "winsock.dll" Alias "#4" (ByVal s%, ByRef Addr As sockaddr_in, ByVal namelen%)
  Declare Function ioctlsocket% Lib "winsock.dll" (ByVal s%, ByVal cmd&, ByRef argp&)
  Declare Function getpeername% Lib "winsock.dll" Alias "#5" (ByVal s%, ByRef peername As sockaddr, ByRef namelen%)
  Declare Function getpeernameIn% Lib "winsock.dll" Alias "#5" (ByVal s%, ByRef peername As sockaddr_in, ByRef namelen%)
  Declare Function getsockname% Lib "winsock.dll" Alias "#6" (ByVal s%, ByRef sockname As sockaddr, ByRef namelen%)
  Declare Function getsocknameIn% Lib "winsock.dll" Alias "#6" (ByVal s%, ByRef sockname As sockaddr_in, ByRef namelen%)
  Declare Function getsockopt% Lib "winsock.dll" Alias "#7" (ByVal s%, ByVal level%, ByVal optname%, ByVal optval$, ByRef optlen%)
  Declare Function htonl& Lib "winsock.dll" Alias "#8" (ByVal hostlong&)
  Declare Function htons% Lib "winsock.dll" Alias "#9" (ByVal hostshort%)
  Declare Function inet_addr& Lib "winsock.dll" Alias "#10" (ByVal cp$)
  Declare Function inet_ntoa& Lib "winsock.dll" Alias "#11" (ByVal inet&)
  Declare Function listen% Lib "winsock.dll" (ByVal s As Integer, ByVal backlog%)
  Declare Function ntohl& Lib "winsock.dll" Alias "#14" (ByVal netlong&)
  Declare Function ntohs% Lib "winsock.dll" Alias "#15" (ByVal netshort%)
  Declare Function recv% Lib "winsock.dll" Alias "#16" (ByVal s%, ByRef buf&, ByVal buflen%, ByVal flags%)
  Declare Function recvfrom% Lib "winsock.dll" Alias "#17" (ByVal s%, ByRef buf&, ByVal buflen%, ByVal flags%, ByRef fromaddr As sockaddr_in, ByRef fromlen%)
  ' Visual Basic note...since select is a keyword in Visual Basic the function
  ' has been renamed
  Declare Function WSASelect% Lib "winsock.dll" Alias "#18" (ByVal nfds%, ByRef readfds As FD_SET, ByRef writefds As FD_SET, ByRef exceptfds As FD_SET, ByVal TimeOut As timeval)
  Declare Function send% Lib "winsock.dll" Alias "#19" (ByVal s%, ByRef buf&, ByVal buflen%, ByVal flags%)
  Declare Function sendto% Lib "winsock.dll" Alias "#20" (ByVal s%, ByRef buf&, ByVal buflen%, ByVal flags%, toaddr As sockaddr_in, ByVal tolen%)
  Declare Function setsockopt% Lib "winsock.dll" Alias "#21" (ByVal s%, ByVal level%, ByVal optname%, ByRef optval&, ByVal optlen%)
  Declare Function shutdown% Lib "winsock.dll" Alias "#22" (ByVal s%, ByVal how%)
  Declare Function Socket% Lib "winsock.dll" Alias "#23" (ByVal af%, ByVal socktype%, ByVal protocol%)
  Declare Function gethostbyaddr& Lib "winsock.dll" Alias "#51" (ByRef Addr&, ByVal addrlen%, ByVal addrtype%)
  Declare Function GetHostByName& Lib "winsock.dll" Alias "#52" (ByVal HostName$)
  Declare Function GetHostName% Lib "winsock.dll" Alias "#57" (ByVal HostName$, ByVal namelen%)
  Declare Function getservbyport& Lib "winsock.dll" Alias "#56" (ByVal Port%, ByVal protoname&)
  Declare Function getservbyname& Lib "winsock.dll" Alias "#55" (ByVal servname&, ByVal protoname&)
  Declare Function getprotobynumber& Lib "winsock.dll" Alias "#54" (ByVal protonumber%)
  Declare Function getprotobyname& Lib "winsock.dll" Alias "#53" (ByVal protoname$)
  Declare Function WSAStartup% Lib "winsock.dll" Alias "#115" (ByVal wVersionRequired%, ByRef lpWSAData As WSAData)
  Declare Function WSACleanUp% Lib "winsock.dll" Alias "#116" ()
  Declare Function WSASetLastError% Lib "winsock.dll" Alias "#112" (ByVal iError%)
  Declare Function WSAGetLastError% Lib "winsock.dll" Alias "#111" ()
  Declare Function WSAIsBlocking% Lib "winsock.dll" Alias "#114" ()
  Declare Function WSAUnhookBlockingHook% Lib "winsock.dll" Alias "#110" ()
  Declare Function WSASetBlockingHook& Lib "winsock.dll" Alias "#109" (ByRef lpFunc&)
  Declare Function WSACancelBlockingCall% Lib "winsock.dll" Alias "#113" ()
  Declare Function WSAAsyncGetServByName% Lib "winsock.dll" Alias "#107" (ByVal hWnd%, ByVal wMsg%, ByVal HostName$, ByVal proto$, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetServByPort% Lib "winsock.dll" Alias "#106" (ByVal hWnd%, ByVal wMsg%, ByVal Port%, ByVal proto$, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetProtoByName% Lib "winsock.dll" Alias "#105" (ByVal hWnd%, ByVal wMsg%, ByVal protoname$, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetProtoByNumber% Lib "winsock.dll" Alias "#104" (ByVal hWnd%, ByVal wMsg%, ByVal number%, ByRef buf&, ByVal buflen%)
  Declare Function WSAAsyncGetHostByName% Lib "winsock.dll" Alias "#103" (ByVal hWnd%, ByVal wMsg%, ByVal HostName$, ByRef buf As BIGhostent, ByVal buflen%)
  Declare Function WSAAsyncGetHostByAddr% Lib "winsock.dll" Alias "#102" (ByVal hWnd%, ByVal wMsg%, ByVal Addr$, ByVal addrlen%, ByVal addrtype%, ByRef buf As BIGhostent, ByVal buflen%)
  Declare Function WSACancelAsyncRequest% Lib "winsock.dll" Alias "#108" (ByVal hAsyncTaskHandle%)
  Declare Function WSAAsyncSelect% Lib "winsock.dll" Alias "#101" (ByVal s%, ByVal hWnd%, ByVal wMsg%, ByVal lEvent&)
#End If

'Winsock Subroutines as described in the Winsock standard
'  look for the file 'hostname.exe' in your system directory
Declare Function GetHostAddressByName& Lib "HostName" Alias "GETHOSTADDRESSFROMNAME" (ByVal yourhostname$)



'########################################################################
'########################################################################
'     Options and Functions specific to the Microsoft NT4.0 ws2_32.dll
'       -Taken from MSWSOCK.H from the winsock 2.2.x Beta1.6 SDK
'       -This section isn't finished, nor is it tested!
'       -Should be a great starting point ;)
'########################################################################
'########################################################################

#If WinSock2 And AddMicroSoft Then
'/*
' * Options for connect and disconnect data and options.  Used only by
' * non-TCP/IP transports such as DECNet, OSI TP4, etc.
' */
Public Const SO_CONNDATA = &H7000
Public Const SO_CONNOPT = &H7001
Public Const SO_DISCDATA = &H7002
Public Const SO_DISCOPT = &H7003
Public Const SO_CONNDATALEN = &H7004
Public Const SO_CONNOPTLEN = &H7005
Public Const SO_DISCDATALEN = &H7006
Public Const SO_DISCOPTLEN = &H7007

'/*
' * Option for opening sockets for synchronous access.
' */
Public Const SO_OPENTYPE = &H7008

Public Const SO_SYNCHRONOUS_ALERT = &H10
Public Const SO_SYNCHRONOUS_NONALERT = &H20

'/*
' * Other NT-specific options.
' */
Public Const SO_MAXDG = &H7009
Public Const SO_MAXPATHDG = &H700A
Public Const SO_UPDATE_ACCEPT_CONTEXT = &H700B
Public Const SO_CONNECT_TIME = &H700C

'/*
' * TCP options.
' */
Public Const TCP_BSDURGENT = &H7000

'/*
' * Microsoft extended APIs.
' */
Declare Function WSARecvEx% Lib "ws2_32.dll" (ByVal s&, ByRef buf&, ByVal buflen%, flags%)

Type TRANSMIT_FILE_BUFFERS
  'Head as LPVOID??
  HeadLength As Long
  'Tail as LPVOID??
  TailLength As Long
End Type

Public Const TF_DISCONNECT = &H1
Public Const TF_REUSE_SOCKET = &H2
Public Const TF_WRITE_BEHIND = &H4
Declare Function TransmitFile Lib "ws2_32.dll" (ByVal hSocket&, hFile&, nNumberOfBytesToWrite&, nNumberOfBytesPerSend&, lpOverlapped%, lpTransmitBuffers As TRANSMIT_FILE_BUFFERS, dwReserved&) As Boolean
Declare Function AcceptEx Lib "ws2_32.dll" (sListenSocket&, sAcceptSocket&, lpOutputBuffer, dwReceiveDataLength&, dwLocalAddressLength&, dwRemoteAddressLength&, lpdwBytesReceived, lpOverlapped) As Boolean
Declare Function GetAcceptExSockaddrs Lib "ws2_32.dll" (lpOutputBuffer, dwReceiveDataLength&, dwLocalAddressLength&, dwRemoteAddressLength&, LocalSockaddr As sockaddr, LocalSockaddrLength%, RemoteSockaddr As sockaddr, RemoteSockaddrLength%)

'/*
' * "QueryInterface" versions of the above APIs.
' */
Type LPFN_TRANSMITFILE
  hSocket As Long
  'hFile As HANDLE?? file handle? 5 to 1 odds its a long integer
  nNumberOfBytesToWrite As Long
  nNumberOfBytesPerSend As Long
  lpOverlapped As Long  'pointer to OVERLAPPED structure
  lpTransmitBuffers As TRANSMIT_FILE_BUFFERS
  dwReserved As Long
End Type

Public WSAID_TRANSMITFILE As GUID
'#define WSAID_TRANSMITFILE \
'        {0xb5367df0,0xcbac,0x11cf,{0x95,0xca,0x00,0x80,0x5f,0x48,0xa1,0x92}}

Type LPFN_ACCEPTEX
    sListenSocket As Long
    sAcceptSocket As Long
    'lpOutputBuffer as PVOID??
    dwReceiveDataLength As Long
    dwLocalAddressLength As Long
    dwRemoteAddressLength As Long
    lpdwBytesReceived As Long  'pointer to long value
    lpOverlapped As Long  'pointer to OVERLAPPED structure
End Type

Public WSAID_ACCEPTEX As GUID
'#define WSAID_ACCEPTEX \
'        {0xb5367df1,0xcbac,0x11cf,{0x95,0xca,0x00,0x80,0x5f,0x48,0xa1,0x92}}

Type LPFN_GETACCEPTEXSOCKADDRS
    'lpOutputBuffer as PVOID??
    dwReceiveDataLength As Long
    dwLocalAddressLength As Long
    dwRemoteAddressLength As Long
    LocalSockaddr As sockaddr
    LocalSockaddrLength As Integer
    RemoteSockaddr As sockaddr
    RemoteSockaddrLength As Integer
End Type

Public WSAID_GETACCEPTEXSOCKADDRS As GUID
'#define WSAID_GETACCEPTEXSOCKADDRS \
'        {0xb5367df2,0xcbac,0x11cf,{0x95,0xca,0x00,0x80,0x5f,0x48,0xa1,0x92}}
#End If

Sub FD_CLR(ByVal fd As Integer, fdset As FD_SET)
    '--------------------
    ' remove fd from set
    '--------------------
    Dim i As Integer
    For i = 0 To (fdset.fd_count) - 1             ' loop thru entries
        If fdset.fd_array(i) = fd Then            ' if match found
            Dim j As Integer
            fdset.fd_count = (fdset.fd_count) - 1 ' reduce count
            While j < fdset.fd_count              ' move others up in array
                fdset.fd_array(j) = fdset.fd_array(j + 1)
            Wend
        End If
    Next i

End Sub

Sub FD_SET(ByVal fd As Integer, fdset As FD_SET)
    '---------------
    ' add fd to set
    '---------------
    If fdset.fd_count < FD_SETSIZE Then
        fdset.fd_array(fdset.fd_count) = fd      'put entry in last
        fdset.fd_count = (fdset.fd_count) + 1    'increment count
    End If
End Sub

Sub FD_ZERO(fdset As FD_SET)
    fdset.fd_count = 0               'set count to zero
End Sub

Function WSA_VERSION(MajorVer As Integer, MinorVer As Integer) As Integer
    WSA_VERSION = (MinorVer * 256) + MajorVer
End Function

Function WSAGETASYNCBUFLEN(lParam As Long) As Integer
    WSAGETASYNCBUFLEN = Int(lParam Mod 65536)
End Function

Function WSAGETASYNCERROR(lParam As Long) As Integer
    WSAGETASYNCERROR = Int(lParam \ 65536) 'Shift Right 8 Places
End Function

Function WSAGETSELECTERROR(lParam As Long) As Integer
    WSAGETSELECTERROR = Int(lParam \ 65536)
End Function

Function WSAGETSELECTEVENT(lParam As Long) As Integer
    WSAGETSELECTEVENT = Int(lParam Mod 65536)
End Function

Function IO&(x$, y%)
 
 '#define _IO(x,y)        (IOC_VOID|((x)<<8)|(y))
 IO = (IOC_VOID Or (Asc(x) * 256) Or (y))

End Function

Function IOR&(x$, y%, t&)
  
  '#define _IOR(x,y,t)     (IOC_OUT|(((long)sizeof(t)&IOCPARM_MASK)<<16)|(x<<8)|y)
  IOR = (IOC_OUT Or ((t And IOCPARM_MASK) * 65536) Or (Asc(x) * 256) Or (y))
  
End Function

Function IOW&(x$, y%, t&)

  '#define _IOW(x,y,t)     (IOC_IN|(((long)sizeof(t)&IOCPARM_MASK)<<16)|(x<<8)|y)
  IOW = (IOC_IN Or ((t And IOCPARM_MASK) * 65536) Or (Asc(x) * 256) Or y)

End Function

#If Win32 And AddMicroSoft Then
' these should be constants, but can't use a user-defined type
' for a constant in this version of VB.
Sub WSAStuffPublics()

WSAID_TRANSMITFILE.data1 = &HB5367DF0
WSAID_TRANSMITFILE.data2 = &HCBAC
WSAID_TRANSMITFILE.data3 = &H11CF
WSAID_TRANSMITFILE.data4(1) = &H95
WSAID_TRANSMITFILE.data4(2) = &HCA
WSAID_TRANSMITFILE.data4(3) = &H0
WSAID_TRANSMITFILE.data4(4) = &H80
WSAID_TRANSMITFILE.data4(5) = &H5F
WSAID_TRANSMITFILE.data4(6) = &H48
WSAID_TRANSMITFILE.data4(7) = &HA1
WSAID_TRANSMITFILE.data4(8) = &H92

WSAID_ACCEPTEX.data1 = &HB5367DF1
WSAID_ACCEPTEX.data2 = &HCBAC
WSAID_ACCEPTEX.data3 = &H11CF
WSAID_ACCEPTEX.data4(1) = &H95
WSAID_ACCEPTEX.data4(2) = &HCA
WSAID_ACCEPTEX.data4(3) = &H0
WSAID_ACCEPTEX.data4(4) = &H80
WSAID_ACCEPTEX.data4(5) = &H5F
WSAID_ACCEPTEX.data4(6) = &H48
WSAID_ACCEPTEX.data4(7) = &HA1
WSAID_ACCEPTEX.data4(8) = &H92

WSAID_GETACCEPTEXSOCKADDRS.data1 = &HB5367DF2
WSAID_GETACCEPTEXSOCKADDRS.data2 = &HCBAC
WSAID_GETACCEPTEXSOCKADDRS.data3 = &H11CF
WSAID_GETACCEPTEXSOCKADDRS.data4(1) = &H95
WSAID_GETACCEPTEXSOCKADDRS.data4(2) = &HCA
WSAID_GETACCEPTEXSOCKADDRS.data4(3) = &H0
WSAID_GETACCEPTEXSOCKADDRS.data4(4) = &H80
WSAID_GETACCEPTEXSOCKADDRS.data4(5) = &H5F
WSAID_GETACCEPTEXSOCKADDRS.data4(6) = &H48
WSAID_GETACCEPTEXSOCKADDRS.data4(7) = &HA1
WSAID_GETACCEPTEXSOCKADDRS.data4(8) = &H92

End Sub
#End If
