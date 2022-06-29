Attribute VB_Name = "Module1"
Option Explicit
Public Const MAX_PATH = 260
Public Const INTERNET_FLAG_RELOAD = &H80000000
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
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As Currency
        ftLastAccessTime As Currency
        ftLastWriteTime As Currency
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

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long

Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_SERVICE_FTP = 1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
   (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Boolean
   
Public Declare Function InternetWriteFile Lib "wininet.dll" _
(ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWite As Long, _
dwNumberOfBytesWritten As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" _
(ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal FLAGS As Long, ByVal Context As Long) As Long

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
      
      


Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Long

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long


Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean


Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))
Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Declare Function InternetGetLastResponseInfo Lib "wininet.dll" _
      Alias "InternetGetLastResponseInfoA" _
       (ByRef lpdwError As Long, _
       ByVal lpszErrorBuffer As String, _
       ByRef lpdwErrorBufferLength As Long) As Boolean
       
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
(ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, _
 ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
 Arguments As Long) As Long

Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long


Function Win32ToVbTime(ft As Currency) As Date

    Dim ftl As Currency

    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then

        ' Local time is nanoseconds since 01-01-1601

        ' In Currency that comes out as milliseconds

        ' Divide by milliseconds per day to get days since 1601

        ' Subtract days from 1601 to 1899 to get VB Date equivalent

        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)

    Else

        MsgBox Err.LastDllError

    End If

End Function


