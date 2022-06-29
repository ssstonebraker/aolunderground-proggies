FTP Simple Sample


This sample demonstrates how to use WinInet FTP APIs in Visual Basic
application.

The sample shows following concepts: 

1. How to enumerate directory on the FTP server and return file
information such as creation date and size.

2. How to upload large files to the FTP server without blocking entire
application and with reporting transfer progress.

There are two ways of uploading a file:   a) by using FtpPutFile ()
API. This API however blocks until entire file has been uploaded. Upon
clicking "Put" button the sample will use this method.

b) by using FtpOpenFile and InternetWriteFile. Once file is open it
can be upload in chunks. This enables application to report upload
status and avoid blocking by calling DoEvents() between calling
InternetWriteFile. Upon clicking "Put Large File" button the sample
will use this method.

3. How to get text information for WinInet errors and how to retrieve
extended error information.

Note: for the simplicity sake sample does not implement downloading of
the large files. This functionality is similar to the method b) above,
however InternetReadFile API instead of InternetWriteFile should be
used.

Notes.

1. Sample uses preconfigured access to the Internet. WinInet FTP APIs
do not work if Internet access is accomplished via CERN type proxy.
Please see this KB for more information:

HOWTO: FTP with CERN-Based Proxy Using WinInet API            [ie_dev]
ID: Q166961    CREATED: 15-APR-1997   MODIFIED: 23-SEP-1998

2. WinInet documentation can be found here:
http://www.microsoft.com/workshop/c-frame.htm#/workshop/networking/default.asp,
click on "Win32 Internet Functions". This KB article list all WinInet
error codes:

WinInet Error Codes (12001 through 12156)[proxysvr] ID: Q193625
CREATED: 02-OCT-1998   MODIFIED: 05-OCT-1998





