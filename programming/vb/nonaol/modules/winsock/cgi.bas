Attribute VB_Name = "modCGI"
'----------------------------------------------------------------------
'       ***********
'       * CGI.BAS *
'       ***********
'
' VERSION: 1.7  (March 18, 1995)
'
' AUTHOR:  Robert B. Denny <rdenny@netcom.com>
'
' Common routines needed to establish a VB environment for
' CGI "scripts" that run behind the Windows Web Server.
'
' INTRODUCTION
'
' The Common Gateway Interface (CGI) version 1.1 specifies a minimal
' set of data that is made available to the back-end application by
' an HTTP (Web) server. It also specifies the details for passing this
' information to the back-end. The latter part of the CGI spec is
' specific to Unix-like environments. The NCSA httpd for Windows does
' supply the data items (and more) specified by CGI/1.1, however it
' uses a different method for passing the data to the back-end.
'
' DEVELOPMENT
'
' Windows httpd requires any Windows back-end program to be an
' executable image. This means that you must convert your VB
' application into an executable (.EXE) before it can be tested
' with the server.
'
' ENVIRONMENT
'
' The Windows httpd server executes script requests by doing a
' WinExec with a command line in the following form:
'
'   prog-name cgi-profile input-file output-file url-args
'
' Assuming you are familiar with the CGI specification, the above
' should be "intuitively obvious" except for the cgi-profile, which
' is described in the next section.
'
' THE CGI PROFILE FILE
'
' The Unix CGI passes data to the back end by defining environment
' variables which can be used by shell scripts. The Windows httpd
' server passes data to its back end via the profile file. The
' format of the profile is that of a Windows ".INI" file. The keyword
' names have been changed cosmetically.
'
' There are 7 sections in a CGI profile file, [CGI], [Accept],
' [System], [Extra Headers], and [Form Literal], [Form External],
' and [Form huge]. They are described below:
'
' [CGI]                <== The standard CGI variables
' CGI Version=         The version of CGI spoken by the server
' Request Protocol=    The server's info protocol (e.g. HTTP/1.0)
' Request Method=      The method specified in the request (e.g., "GET")
' Executable Path=     Physical pathname of the back-end (this program)
' Logical Path=        Extra path info in logical space
' Physical Path=       Extra path info in local physical space
' Query String=        String following the "?" in the request URL
' Content Type=        MIME content type of info supplied with request
' Content Length=      Length, bytes, of info supplied with request
' Server Software=     Version/revision of the info (HTTP) server
' Server Name=         Server's network hostname (or alias from config)
' Server Port=         Server's network port number
' Server Admin=        E-Mail address of server's admin. (config)
' Referer=             URL of referring document (HTTP/1.0 draft 12/94)
' From=                E-Mail of client user  (HTTP/1.0 draft 12/94)
' Remote Host=         Remote client's network hostname
' Remote Address=      Remote client's network address
' Authenticated Username=Username if present in request
' Authenticated Password=Password if present in request
' Authentication Method=Method used for authentication (e.g., "Basic")
' Authentication Realm=Name of realm for users/groups
' RFC-931 Identity=    (deprecated, removed from code)
'
' [Accept]             <== What the client says it can take
' The MIME types found in the request header as
'    Accept: xxx/yyy; zzzz...
' are entered in this section as
'    xxx/yyy=zzzz...
' If only the MIME type appears, the form is
'    xxx/yyy=Yes
'
' [System]             <== Windows interface specifics
' GMT Offset=          Offset of local timezone from GMT, seconds (LONG!)
' Output File=         Pathname of file to receive results
' Content File=        Pathname of file containing request content (raw)
' Debug Mode=          If server's back-end debug flag is set (Yes/No)
'
' [Extra Headers]
' Any "extra" headers found in the request that activated this
' program. They are listed in "key=value" form. Usually, you'll see
' at least the name of the browser here.
'
' [Form Literal]
' If the request was a POST from a Mosaic form (with content type of
' "application/x-www-form-urlencoded"), the server will decode the
' form data. Raw form input is of the form "key=value&key=value&...",
' with the value parts "URL-encoded". The server splits the key=value
' pairs at the '&', then spilts the key and value at the '=',
' URL-decodes the value string and puts the result into key=value
' (decoded) form in the [Form Literal] section of the INI.
'
' [Form External]
' If the decoded value string is more than 254 characters long,
' or if the decoded value string contains any control characters,
' the server puts the decoded value into an external tempfile and
' lists the field in this section as:
'    key=<pathname> <length>
' where <pathname> is the path and name of the tempfile containing
' the decoded value string, and <length> is the length in bytes
' of the decoded value string.
'
' NOTE: BE SURE TO OPEN THIS FILE IN BINARY MODE UNLESS YOU ARE
'       CERTAIN THAT THE FORM DATA IS TEXT!
'
' [Form Huge]
' If the raw value string is more than 65,536 bytes long, the server
' does no decoding. In this case, the server lists the field in this
' section as:
'    key=<offset> <length>
' where <offset> is the offset from the beginning of the Content File
' at which the raw value string for this key is located, and <length>
' is the length in bytes of the raw value string. You can use the
' <offset> to perform a "Seek" to the start of the raw value string,
' and use the length to know when you have read the entire raw string
' into your decoder. Note that VB has a limit of 64K for strings, so
'
' Examples:
'
'    [Form Literal]
'    smallfield=123 Main St. #122
'
'    [Form External]
'    field300chars=C:\TEMP\HS19AF6C.000 300
'    fieldwithlinebreaks=C:\TEMP\HS19AF6C.001 43
'
'    [Form Huge]
'    field230K=C:\TEMP\HS19AF6C.002 276920
'
' =====
' USAGE
' =====
' Include CGI.BAS in your VB project. Set the project options for
' "Sub Main" startup. The Main() procedure is in this module, and it
' handles all of the setup of the VB CGI environment, as described
' above. Once all of this is done, the Main() calls YOUR main procedure
' which must be called CGI_Main(). The output file is open, use Send()
' to write to it. The input file is NOT open, and "huge" form fields
' have not been decoded.
'
' (New in V1.3) If your program is started without command-line args,
' the code assumes you want to run it interactively. This is useful
' for providing a setup screen, etc. Instead of calling CGI_Main(),
' it calls Inter_Main(). Your module must also implement this
' function. If you don't need an interactive mode, just create
' Inter_Main() and put a 1-line call to MsgBox alerting the
' user that the program is not meant to be run interactively.
' The samples furnished with the server do this.
'
' If a Visual Basic runtime error occurs, it will be trapped and result
' in an HTTP error response being sent to the client. Check out the
' Error Handler() sub. When your program finishes, be sure to RETURN
' TO MAIN(). Don't just do an "End".
'
' Have a look at the stuff below to see what's what.
'
'----------------------------------------------------------------------
' Author:   Robert B. Denny <rdenny@netcom.com>
'           June 7, 1994
'
' Revision History:
'   26-May-94 rbd   Initial experimental release
'   07-Jun-94 rbd   Revised keyword names and form decoding per
'                   httpd 1.2b8, fixed section name of Output File.
'   13-Dec-94 rbd   Move FreeFile calls to just before opening files.
'   04-Feb-94 rbd   Fix Authenticated User -> Username, added new
'                   variables Referer, From and Authentication Realm.
'                   Removed RFC931 stuff (deprecated)
'   11-Feb-95 rbd   Added Inter_Main() support and stub.
'   01-Mar-95 rbd   Add support for password pass-through, clean
'                   up HTML in error handler. Add SendNoOp().
'                   Add Server: header to error handler msg.
'   17-Mar-95 rbd   Fix error handler to remove deprecated
'                   MIME-Version header. Add GMT offset, new CGI var.
'                   Add WebDate() function for producing HTTP/1.0
'                   compliant date/time. Add Date: header to error
'                   messages.
'   18-Mar-95 rbd   Add CGI_ERR_START for catching CGI.BAS defined
'                   errors in user code. Decode our "user defined:
'                   errors in handler instead of saying "User
'                   defined error".
'----------------------------------------------------------------------
Option Explicit
'
' ==================
' Manifest Constants
' ==================
'
Const MAX_CMDARGS = 8       ' Max # of command line args
Const ENUM_BUF_SIZE = 4096  ' Key enumeration buffer, see GetProfile()
' These are the limits in the server
Const MAX_XHDR = 100        ' Max # of "extra" request headers
Const MAX_ACCTYPE = 100     ' Max # of Accept: types in request
Const MAX_FORM_TUPLES = 100 ' Max # form key=value pairs
Const MAX_HUGE_TUPLES = 16  ' Max # "huge" form fields
'
'
' =====
' Types
' =====
'
Type Tuple                  ' Used for Accept: and "extra" headers
    key As String           ' and for holding POST form key=value pairs
    value As String
End Type

Type HugeTuple              ' Used for "huge" form fields
    key As String           ' Keyword (decoded)
    offset As Long          ' Byte offset into Content File of value
    length As Long          ' Length of value, bytes
End Type
'
'
' ================
' Global Constants
' ================
'
' -----------
' Error Codes
' -----------
'
Global Const ERR_ARGCOUNT = 32767
Global Const ERR_BAD_REQUEST = 32766        ' HTTP 400
Global Const ERR_UNAUTHORIZED = 32765       ' HTTP 401
Global Const ERR_PAYMENT_REQUIRED = 32764   ' HTTP 402
Global Const ERR_FORBIDDEN = 32763          ' HTTP 403
Global Const ERR_NOT_FOUND = 32762          ' HTTP 404
Global Const ERR_INTERNAL_ERROR = 32761     ' HTTP 500
Global Const ERR_NOT_IMPLEMENTED = 32760    ' HTTP 501
Global Const ERR_TOO_BUSY = 32758           ' HTTP 503 (experimental)
Global Const ERR_NO_FIELD = 32757           ' GetxxxField "no field"
Global Const CGI_ERR_START = 32757          ' Start of our errors

' ====================
' CGI Global Variables
' ====================
'
' ----------------------
' Standard CGI variables
' ----------------------
'
Global CGI_ServerSoftware As String
Global CGI_ServerName As String
Global CGI_ServerPort As Integer
Global CGI_RequestProtocol As String
Global CGI_ServerAdmin As String
Global CGI_Version As String
Global CGI_RequestMethod As String
Global CGI_LogicalPath As String
Global CGI_PhysicalPath As String
Global CGI_ExecutablePath As String
Global CGI_QueryString As String
Global CGI_Referer As String
Global CGI_From As String
Global CGI_RemoteHost As String
Global CGI_RemoteAddr As String
Global CGI_AuthUser As String
Global CGI_AuthPass As String
Global CGI_AuthType As String
Global CGI_AuthRealm As String
Global CGI_ContentType As String
Global CGI_ContentLength As Long
'
' ------------------
' HTTP Header Arrays
' ------------------
'
Global CGI_AcceptTypes(MAX_ACCTYPE) As Tuple    ' Accept: types
Global CGI_NumAcceptTypes As Integer            ' # of live entries in array
Global CGI_ExtraHeaders(MAX_XHDR) As Tuple      ' "Extra" headers
Global CGI_NumExtraHeaders As Integer           ' # of live entries in array
'
' --------------
' POST Form Data
' --------------
'
Global CGI_FormTuples(MAX_FORM_TUPLES) As Tuple ' POST form key=value pairs
Global CGI_NumFormTuples As Integer             ' # of live entries in array
Global CGI_HugeTuples(MAX_HUGE_TUPLES) As HugeTuple ' Form "huge tuples
Global CGI_NumHugeTuples As Integer             ' # of live entries in array

'
' ----------------
' System Variables
' ----------------
'
Global CGI_GMTOffset As Variant         ' GMT offset (time serial)
Global CGI_ContentFile As String        ' Content/Input file pathname
Global CGI_OutputFile As String         ' Output file pathname
Global CGI_DebugMode As Integer         ' Script Tracing flag from server
'
'
' ========================
' Windows API Declarations
' ========================
'
' NOTE: Declaration of GetPrivateProfileString is specially done to
' permit enumeration of keys by passing NULL key value. See GetProfile().
'
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpSection As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
'
'
' ===============
' Local Variables
' ===============
'
Dim CGI_ProfileFile As String           ' Profile file pathname
Dim CGI_OutputFN As Integer             ' Output file number
Dim ErrorString As String




'---------------------------------------------------------------------------
'
'   ErrorHandler() - Global error handler
'
' If a VB runtime error occurs dusing execution of the program, this
' procedure generates an HTTP/1.0 HTML-formatted error message into
' the output file, then exits the program.
'
' This should be armed immediately on entry to the program's main()
' procedure. Any errors that occur in the program are caught, and
' an HTTP/1.0 error messsage is generated into the output file. The
' presence of the HTTP/1.0 on the first line of the output file causes
' NCSA httpd for WIndows to send the output file to the client with no
' interpretation or other header parsing.
'---------------------------------------------------------------------------
Sub ErrorHandler(code As Integer)

    On Error Resume Next     ' Give it a good try!

    Seek #CGI_OutputFN, 1    ' Rewind output file just in case
    Send ("HTTP/1.0 500 Internal Error")
    Send ("Server: " + CGI_ServerSoftware)
    Send ("Date: " + WebDate(Now))
    Send ("Content-type: text/html")
    Send ("")
    Send ("<HTML><HEAD>")
    Send ("<TITLE>Error in " + CGI_ExecutablePath + "</TITLE>")
    Send ("</HEAD><BODY>")
    Send ("<H1>Error in " + CGI_ExecutablePath + "</H1>")
    Send ("An internal Visual Basic error has occurred in " + CGI_ExecutablePath + ".")
    Send ("<PRE>" + ErrorString + "</PRE>")
    Send ("<I>Please</I> note what you were doing when this problem occurred,")
    Send ("so we can identify and correct it. Write down the Web page you were using,")
    Send ("any data you may have entered into a form or search box, and")
    Send ("anything else that may help us duplicate the problem. Then contact the")
    Send ("administrator of this service: ")
    Send ("<A HREF=""mailto:" & CGI_ServerAdmin & """>")
    Send ("<ADDRESS>&lt;" + CGI_ServerAdmin + "&gt;</ADDRESS>")
    Send ("</A></BODY></HTML>")

    Close #CGI_OutputFN

    '======
     End            ' Terminate the program
    '======
End Sub

'---------------------------------------------------------------------------
'
'   GetAcceptTypes() - Create the array of accept type structs
'
' Enumerate the keys in the [Accept] section of the profile file,
' then get the value for each of the keys.
'---------------------------------------------------------------------------
Private Sub GetAcceptTypes()
    Dim sList As String
    Dim i As Integer, j As Integer, l As Integer, n As Integer

    sList = GetProfile("Accept", "") ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    n = 0                                   ' Index in array
    Do While ((i < l) And (n < MAX_ACCTYPE)) ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_AcceptTypes(n).key = Mid$(sList, i, j - i) ' Get Key, then value
        CGI_AcceptTypes(n).value = GetProfile("Accept", CGI_AcceptTypes(n).key)
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    CGI_NumAcceptTypes = n                  ' Fill in global count

End Sub

'---------------------------------------------------------------------------
'
'   GetArgs() - Parse the command line
'
' Chop up the command line, fill in the argument vector, return the
' argument count (similar to the Unix/C argc/argv handling)
'---------------------------------------------------------------------------
Private Function GetArgs(argv() As String) As Integer
    Dim buf As String
    Dim i As Integer, j As Integer, l As Integer, n As Integer

    buf = Trim$(Command$)                   ' Get command line

    l = Len(buf)                            ' Length of command line
    If l = 0 Then                           ' If empty
        GetArgs = 0                         ' Return argc = 0
        Exit Function
    End If

    i = 1                                   ' Start at 1st character
    n = 0                                   ' Index in argvec
    Do While ((i < l) And (n < MAX_CMDARGS)) ' Safety stop here
        j = InStr(i, buf, " ")              ' J -> next space
        If j = 0 Then Exit Do               ' Exit loop on last arg
        argv(n) = Trim$(Mid$(buf, i, j - i)) ' Get this token, trim it
        i = j + 1                           ' Skip that blank
        Do While Mid$(buf, i, 1) = " "      ' Skip any additional whitespace
            i = i + 1
        Loop
        n = n + 1                           ' Bump array index
    Loop

    argv(n) = Trim$(Mid$(buf, i, (l - i + 1))) ' Get last arg
    GetArgs = n + 1                         ' Return arg count

End Function

'---------------------------------------------------------------------------
'
'   GetExtraHeaders() - Create the array of extra header structs
'
' Enumerate the keys in the [Extra Headers] section of the profile file,
' then get the value for each of the keys.
'---------------------------------------------------------------------------
Private Sub GetExtraHeaders()
    Dim sList As String
    Dim i As Integer, j As Integer, l As Integer, n As Integer

    sList = GetProfile("Extra Headers", "") ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    n = 0                                   ' Index in array
    Do While ((i < l) And (n < MAX_XHDR))   ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_ExtraHeaders(n).key = Mid$(sList, i, j - i) ' Get Key, then value
        CGI_ExtraHeaders(n).value = GetProfile("Extra Headers", CGI_ExtraHeaders(n).key)
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    CGI_NumExtraHeaders = n                 ' Fill in global count

End Sub

'---------------------------------------------------------------------------
'
'   GetFormTuples() - Create the array of POST form input key=value pairs
'
'---------------------------------------------------------------------------
Private Sub GetFormTuples()
    Dim sList As String
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, n As Integer
    Dim s As Long
    Dim buf As String
    Dim extName As String
    Dim extFile As Integer
    Dim extlen As Long

    n = 0                                   ' Index in array

    '
    ' Do the easy one first: [Form Literal]
    '
    sList = GetProfile("Form Literal", "")  ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    Do While ((i < l) And (n < MAX_FORM_TUPLES)) ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_FormTuples(n).key = Mid$(sList, i, j - i) ' Get Key, then value
        CGI_FormTuples(n).value = GetProfile("Form Literal", CGI_FormTuples(n).key)
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    '
    ' Now do the external ones: [Form External]
    '
    sList = GetProfile("Form External", "") ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    extFile = FreeFile
    Do While ((i < l) And (n < MAX_FORM_TUPLES)) ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_FormTuples(n).key = Mid$(sList, i, j - i) ' Get Key, then pathname
        buf = GetProfile("Form External", CGI_FormTuples(n).key)
        k = InStr(buf, " ")                 ' Split file & length
        extName = Mid$(buf, 1, k - 1)           ' Pathname
        k = k + 1
        extlen = CLng(Mid$(buf, k, Len(buf) - k + 1)) ' Length
        '
        ' Use feature of GET to read content in one call
        '
        Open extName For Binary Access Read As #extFile
        CGI_FormTuples(n).value = String$(extlen, " ") ' Breathe in...
        Get #extFile, , CGI_FormTuples(n).value 'GULP!
        Close #extFile
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop

    CGI_NumFormTuples = n                   ' Number of fields decoded
    n = 0                                   ' Reset counter
    '
    ' Finally, the [Form Huge] section. Will this ever get executed?
    '
    sList = GetProfile("Form Huge", "")     ' Get key list
    l = Len(sList)                          ' Length incl. trailing null
    i = 1                                   ' Start at 1st character
    Do While ((i < l) And (n < MAX_FORM_TUPLES)) ' Safety stop here
        j = InStr(i, sList, Chr$(0))        ' J -> next null
        CGI_HugeTuples(n).key = Mid$(sList, i, j - i) ' Get Key
        buf = GetProfile("Form Huge", CGI_HugeTuples(n).key) ' "offset length"
        k = InStr(buf, " ")                 ' Delimiter
        CGI_HugeTuples(n).offset = CLng(Mid$(buf, 1, (k - 1)))
        CGI_HugeTuples(n).length = CLng(Mid$(buf, k, (Len(buf) - k + 1)))
        i = j + 1                           ' Bump pointer
        n = n + 1                           ' Bump array index
    Loop
    
    CGI_NumHugeTuples = n                   ' Fill in global count

End Sub

'---------------------------------------------------------------------------
'
'   GetProfile() - Get a value or enumerate keys in CGI_Profile file
'
' Get a value given the section and key, or enumerate keys given the
' section name and "" for the key. If enumerating, the list of keys for
' the given section is returned as a null-separated string, with a
' double null at the end.
'
' VB handles this with flair! I couldn't believe my eyes when I tried this.
'---------------------------------------------------------------------------
Private Function GetProfile(sSection As String, sKey As String) As String
    Dim retLen As Integer
    Dim buf As String * ENUM_BUF_SIZE

    If sKey <> "" Then
        retLen = GetPrivateProfileString(sSection, sKey, "", buf, ENUM_BUF_SIZE, CGI_ProfileFile)
    Else
        retLen = GetPrivateProfileString(sSection, 0&, "", buf, ENUM_BUF_SIZE, CGI_ProfileFile)
    End If
    If retLen = 0 Then
        GetProfile = ""
    Else
        GetProfile = Left$(buf, retLen)
    End If

End Function

'----------------------------------------------------------------------
'
' Get the value of a "small" form field given the key
'
' Signals an error if field does not exist
'
'----------------------------------------------------------------------
Function GetSmallField(key As String) As String
    Dim i As Integer

    For i = 0 To (CGI_NumFormTuples - 1)
        If CGI_FormTuples(i).key = key Then
            GetSmallField = Trim$(CGI_FormTuples(i).value)
            Exit Function           ' ** DONE **
        End If
    Next i
    '
    ' Field does not exist
    '
    Error ERR_NO_FIELD
End Function

'---------------------------------------------------------------------------
'
'   InitializeCGI() - Fill in all of the CGI variables, etc.
'
' Read the profile file name from the command line, then fill in
' the CGI globals, the Accept type list and the Extra headers list.
' Then open the input and output files.
'
' Returns True if OK, False if some sort of error. See ReturnError()
' for info on how errors are handled.
'
' NOTE: Assumes that the CGI error handler has been armed with On Error
'---------------------------------------------------------------------------
Sub InitializeCGI()
    Dim sect As String
    Dim argc As Integer
    Static argv(MAX_CMDARGS) As String
    Dim buf As String

    CGI_DebugMode = True    ' Initialization errors are very bad

    '
    ' Parse the command line. We need the profile file name (duh!)
    ' and the output file name NOW, so we can return any errors we
    ' trap. The error handler writes to the output file.
    '
    argc = GetArgs(argv())
    CGI_ProfileFile = argv(0)

    sect = "CGI"
    CGI_ServerSoftware = GetProfile(sect, "Server Software")
    CGI_ServerName = GetProfile(sect, "Server Name")
    CGI_RequestProtocol = GetProfile(sect, "Request Protocol")
    CGI_ServerAdmin = GetProfile(sect, "Server Admin")
    CGI_Version = GetProfile(sect, "CGI Version")
    CGI_RequestMethod = GetProfile(sect, "Request Method")
    CGI_LogicalPath = GetProfile(sect, "Logical Path")
    CGI_PhysicalPath = GetProfile(sect, "Physical Path")
    CGI_ExecutablePath = GetProfile(sect, "Executable Path")
    CGI_QueryString = GetProfile(sect, "Query String")
    CGI_RemoteHost = GetProfile(sect, "Remote Host")
    CGI_RemoteAddr = GetProfile(sect, "Remote Address")
    CGI_Referer = GetProfile(sect, "Referer")
    CGI_From = GetProfile(sect, "From")
    CGI_AuthUser = GetProfile(sect, "Authenticated Username")
    CGI_AuthPass = GetProfile(sect, "Authenticated Password")
    CGI_AuthRealm = GetProfile(sect, "Authentication Realm")
    CGI_AuthType = GetProfile(sect, "Authentication Method")
    CGI_ContentType = GetProfile(sect, "Content Type")
    buf = GetProfile(sect, "Content Length")
    If buf = "" Then
        CGI_ContentLength = 0
    Else
        CGI_ContentLength = CLng(buf)
    End If
    buf = GetProfile(sect, "Server Port")
    If buf = "" Then
        CGI_ServerPort = -1
    Else
        CGI_ServerPort = CInt(buf)
    End If

    sect = "System"
    CGI_ContentFile = GetProfile(sect, "Content File")
    CGI_OutputFile = GetProfile(sect, "Output File")     'argv(2)
    CGI_OutputFN = FreeFile
    Open CGI_OutputFile For Output Access Write As #CGI_OutputFN
    buf = GetProfile(sect, "GMT Offset")
    CGI_GMTOffset = CVDate(CDbl(buf) / 86400#) ' Timeserial offset
    buf = GetProfile(sect, "Debug Mode")    ' Y or N
    If (Left$(buf, 1) = "Y") Then           ' Must start with Y
        CGI_DebugMode = True
    Else
        CGI_DebugMode = False
    End If

    GetAcceptTypes          ' Enumerate Accept: types into tuples
    GetExtraHeaders         ' Enumerate extra headers into tuples
    GetFormTuples           ' Decode any POST form input into tuples

End Sub

'----------------------------------------------------------------------
'
'   main() - CGI script back-end main procedure
'
' This is the main() for the VB back end. Note carefully how the error
' handling is set up, and how program cleanup is done. If no command
' line args are present, call Inter_Main() and exit.
'----------------------------------------------------------------------
Sub Main()
    On Error GoTo ErrorHandler

    If Trim$(Command$) = "" Then    ' Interactive start
        Inter_Main                  ' Call interactive main
        Exit Sub                    ' Exit the program
    End If

    InitializeCGI       ' Create the CGI environment

    '===========
    CGI_Main            ' Execute the actual "script"
    '===========

Cleanup:
    Close #CGI_OutputFN
    Exit Sub                        ' End the program
'------------
ErrorHandler:
    Select Case Err                 ' Decode our "user defined" errors
        Case ERR_NO_FIELD:
            ErrorString = "Unknown form field"
        Case Else:
            ErrorString = Error$    ' Must be VB error
    End Select

    ErrorString = ErrorString & " (error #" & Err & ")"
    On Error GoTo 0                 ' Prevent recursion
    ErrorHandler (Err)              ' Generate HTTP error result
    Resume Cleanup
'------------
End Sub

'----------------------------------------------------------------------
'
'  Send() - Shortcut for writing to output file
'
'----------------------------------------------------------------------
Sub Send(s As String)
    Print #CGI_OutputFN, s
End Sub

'---------------------------------------------------------------------------
'
'   SendNoOp() - Tell browser to do nothing.
'
' Most browsers will do nothing. Netscape 1.0N leaves hourglass
' cursor until the mouse is waved around. Enhanced Mosaic 2.0
' oputs up an alert saying "URL leads nowhere". Your results may
' vary...
'
'---------------------------------------------------------------------------
Sub SendNoOp()

    Send ("HTTP/1.0 204 No Response")
    Send ("Server: " + CGI_ServerSoftware)
    Send ("")

End Sub

'---------------------------------------------------------------------------
'
'   WebDate - Return an HTTP/1.0 compliant date/time string
'
' Inputs:   t = Local time as VB Variant (e.g., returned by Now())
' Returns:  Properly formatted HTTP/1.0 date/time in GMT
'---------------------------------------------------------------------------
Function WebDate(dt As Variant) As String
    Dim t As Variant
    
    t = CVDate(dt - CGI_GMTOffset)      ' Convert time to GMT
    WebDate = Format$(t, "ddd dd mmm yyyy hh:mm:ss") & " GMT"

End Function

