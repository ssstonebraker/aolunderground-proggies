Attribute VB_Name = "modWinInet"
Option Explicit

Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Public Const MAX_PATH = 260
Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000


Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Public Const ERROR_NO_MORE_FILES = 18

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' InternetErrorDlg
Public Declare Function InternetErrorDlg Lib "wininet.dll" _
(ByVal hWnd As Long, ByVal hInternet As Long, ByVal dwError As Long, ByVal dwFlags As Long, ByVal lppvData As Long) As Long

' InternetErrorDlg constants
Public Const FLAGS_ERROR_UI_FILTER_FOR_ERRORS = &H1
Public Const FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS = &H2
Public Const FLAGS_ERROR_UI_FLAGS_GENERATE_DATA = &H4
Public Const FLAGS_ERROR_UI_FLAGS_NO_UI = &H8
Public Const FLAGS_ERROR_UI_SERIALIZE_DIALOGS = &H10

Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Type INTERNET_BUFFERS
    dwStructSize As Long        ' used for API versioning. Set to sizeof(INTERNET_BUFFERS)
    Next As Long                ' INTERNET_BUFFERS chain of buffers
    lpcszHeader As Long       ' pointer to headers (may be NULL)
    dwHeadersLength As Long     ' length of headers if not NULL
    dwHeadersTotal As Long      ' size of headers if not enough buffer
    lpvBuffer As Long           ' pointer to data buffer (may be NULL)
    dwBufferLength As Long      ' length of data buffer if not NULL
    dwBufferTotal As Long       ' total size of chunk, or content-length if not chunked
    dwOffsetLow As Long         ' used for read-ranges (only used in HttpSendRequest2)
    dwOffsetHigh As Long
End Type

Public Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" _
(ByVal hHttpRequest As Long, lpBuffersIn As INTERNET_BUFFERS, ByVal lpBuffersOut As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Public Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" _
(ByVal hHttpRequest As Long, ByVal lpBuffersOut As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long


Public Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

' Internet Errors
Public Const INTERNET_ERROR_BASE = 12000

Public Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Public Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Public Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Public Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Public Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Public Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Public Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Public Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Public Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Public Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Public Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Public Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Public Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Public Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Public Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Public Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Public Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Public Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Public Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)

Public Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Public Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Public Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Public Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Public Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Public Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Public Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)

' FTP API errors

Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Public Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Public Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)

' gopher API errors

Public Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
Public Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
Public Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
Public Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
Public Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)

' HTTP API errors

Public Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
Public Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
Public Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
Public Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
Public Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
Public Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
Public Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)

' additional Internet API error codes

Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Public Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Public Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Public Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)

Public Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Public Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Public Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Public Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)

' InternetAutodial specific errors

Public Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)

Public Const INTERNET_ERROR_LAST = ERROR_INTERNET_FAILED_DUETOSECURITYCHECK

'
' flags common to open functions (not InternetOpen()):
'

Public Const INTERNET_FLAG_RELOAD = &H80000000             ' retrieve the original item

'
' flags for InternetOpenUrl():
'

Public Const INTERNET_FLAG_RAW_DATA = &H40000000           ' FTP/gopher find: receive the item as raw (structured) data
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible

'
' flags for InternetOpen():
'

Public Const INTERNET_FLAG_ASYNC = &H10000000              ' this request is asynchronous (where supported)

'
' protocol-specific flags:
'

Public Const INTERNET_FLAG_PASSIVE = &H8000000             ' used for FTP connections

'
' additional cache flags
'

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Public Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000     ' make this item persistent in cache
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000          ' use offline semantics
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE

'
' additional flags
'

Public Const INTERNET_FLAG_SECURE = &H800000               ' use PCT/SSL if applicable (HTTP)
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000      ' use keep-alive semantics
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000     ' don't handle redirections automatically
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000        ' do background read prefetch
Public Const INTERNET_FLAG_NO_COOKIES = &H80000            ' no automatic cookie handling
Public Const INTERNET_FLAG_NO_AUTH = &H40000               ' no automatic authentication handling
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000     ' return cache file if net request fails

'
' Security Ignore Flags, Allow HttpOpenRequest to overide
'  Secure Channel (SSL/PCT) failures of the following types.
'

Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000       ' ex: https:// to http://
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000      ' ex: http:// to https://
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000      ' expired X509 Cert.
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000        ' bad common name in X509 Cert.

'
' more caching flags
'

Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800           ' asking wininet to update an item if it is newer
Public Const INTERNET_FLAG_HYPERLINK = &H400               ' asking wininet to do hyperlinking semantic which works right for scripts
Public Const INTERNET_FLAG_NO_UI = &H200                   ' no cookie popup
Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100          ' asking wininet to add "pragma: no-cache"
Public Const INTERNET_FLAG_CACHE_ASYNC = &H80              ' ok to perform lazy cache-write
Public Const INTERNET_FLAG_FORMS_SUBMIT = &H40             ' this is a forms submit
Public Const INTERNET_FLAG_NEED_FILE = &H10                ' need a file for this request
Public Const INTERNET_FLAG_MUST_CACHE_REQUEST = INTERNET_FLAG_NEED_FILE

'
' flags for FTP
'

Public Const INTERNET_FLAG_TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII       ' = &H00000001
Public Const INTERNET_FLAG_TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY     ' = &H00000002

'
' flags field masks
'

Public Const SECURITY_INTERNET_MASK = INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
                                 INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or _
                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Or _
                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP

Public Const INTERNET_FLAGS_MASK = INTERNET_FLAG_RELOAD Or _
                                 INTERNET_FLAG_RAW_DATA Or _
                                 INTERNET_FLAG_EXISTING_CONNECT Or _
                                 INTERNET_FLAG_ASYNC Or _
                                 INTERNET_FLAG_PASSIVE Or _
                                 INTERNET_FLAG_NO_CACHE_WRITE Or _
                                 INTERNET_FLAG_MAKE_PERSISTENT Or _
                                 INTERNET_FLAG_FROM_CACHE Or _
                                 INTERNET_FLAG_SECURE Or _
                                 INTERNET_FLAG_KEEP_CONNECTION Or _
                                 INTERNET_FLAG_NO_AUTO_REDIRECT Or _
                                 INTERNET_FLAG_READ_PREFETCH Or _
                                 INTERNET_FLAG_NO_COOKIES Or _
                                 INTERNET_FLAG_NO_AUTH Or _
                                 INTERNET_FLAG_CACHE_IF_NET_FAIL Or _
                                 SECURITY_INTERNET_MASK Or _
                                 INTERNET_FLAG_RESYNCHRONIZE Or _
                                 INTERNET_FLAG_HYPERLINK Or _
                                 INTERNET_FLAG_NO_UI Or _
                                 INTERNET_FLAG_PRAGMA_NOCACHE Or _
                                 INTERNET_FLAG_CACHE_ASYNC Or _
                                 INTERNET_FLAG_FORMS_SUBMIT Or _
                                 INTERNET_FLAG_NEED_FILE Or _
                                 INTERNET_FLAG_TRANSFER_BINARY Or _
                                 INTERNET_FLAG_TRANSFER_ASCII
                                

Public Const INTERNET_ERROR_MASK_INSERT_CDROM = &H1

Public Const INTERNET_OPTIONS_MASK = (Not INTERNET_FLAGS_MASK)

'
' common per-API flags (new APIs)
'

Public Const WININET_API_FLAG_ASYNC = &H1                  ' force async operation
Public Const WININET_API_FLAG_SYNC = &H4                   ' force sync operation
Public Const WININET_API_FLAG_USE_CONTEXT = &H8            ' use value supplied in dwContext (even if 0)

'
' INTERNET_NO_CALLBACK - if this value is presented as the dwContext parameter
' then no call-backs will be made for that API
'

Public Const INTERNET_NO_CALLBACK = 0

Public Declare Function InternetDial Lib "wininet.dll" _
(ByVal hWnd As Long, ByVal sConnectoid As String, _
 ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long

Public Declare Function InternetHangUp Lib "wininet.dll" _
(ByVal dwConnection As Long, ByVal dwReserved As Long) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hInternetSession As Long, ByVal sUrl As String, _
ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long


