VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl download 
   BackColor       =   &H80000008&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "download.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1590
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private CancelSearch As Boolean
Public DownloadSuccess As Boolean

Public Function FormatFileSize(ByVal dblFileSize As Double) As String

' FormatFileSize:   Formats dblFileSize in bytes into
'                   X GB or X MB or X KB or X bytes depending
'                   on size (a la Win9x Properties tab)

Select Case dblFileSize
    Case 0 To 999   ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1000 To 1023999    ' KB
        FormatFileSize = Format(dblFileSize / 1024, "##0.0") & " KB"
    Case 1024000 To (1024 * 10 ^ 6) - 1 ' MB
        FormatFileSize = Format(dblFileSize / (1024 ^ 2), "##0.0#") & " MB"
    Case Is > (1024 * 10 ^ 6)
        FormatFileSize = Format(dblFileSize / (1024 ^ 3), "##0.0#") & " GB"
End Select

End Function

Public Function CancelDownload()
    CancelSearch = True
End Function

Public Function FormatTime(ByVal sglTime As Single) As String
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599 ' Minutes Seconds
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else       ' Hours Minutes
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))

End Function


Public Function DownloadFile(strURL As String, strDestination As String, pbar As ProgressBar, downloadedpart As Label, timeleft As Label, Optional UserName As String = "", Optional Password As String = "") As Boolean

' Funtion DownloadFile
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

Dim bData() As Byte         ' Data var
Dim intFile As Integer      ' FreeFile var
Dim a As Variant            ' Temp var
Dim intKReceived As Integer ' KB received so far
Dim intKFileLength As Long  ' KB total length of file
Dim lastTime As Single      ' time last chunk received
Dim sglRate As Single       ' var to hold transfer rate
Dim sglTime As Single       ' var to hold time remaining
Dim strFile As String       ' temp filename var
Dim strHeader As String     ' HTTP header store
Dim strHost As String       ' HTTP Host

On Local Error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, False, True)

' Show the status form
'Animation1.Open App.Path & "\downld2.avi"
downloadedpart.Caption = "Getting file information..."
DoEvents

' Download file
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    .Execute , "GET"
End With

downloadedpart.Caption = "Saving:" & vbCr & vbCr & strFile & " from " _
              & IIf(Len(strHost) < 33, strHost, "..." & Left(strHost, 30))

lastTime = Timer

' While initiating connection, yield CPU to Windows
While Inet1.StillExecuting
    DoEvents
    ' If user pressed Cancel button on StatusForm
    ' then fail, cancel, and exit this download
    If CancelSearch Then
        GoTo ExitDownload
    End If
Wend

' Get first header ("HTTP/X.X XXX ...")
strHeader = Inet1.GetHeader

' Trap common HTTP Errors
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK!

    Case "401"  ' Not authorized
      
        MsgBox "Authorization failed!", _
               vbCritical, _
               "Unauthorized"
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
     
        MsgBox "The file, " & _
               vbDoubleQuote & Inet1.URL & vbDoubleQuote & _
               " was not found!", _
               vbCritical, _
               "File Not Found"
        GoTo ExitDownload
        
    Case vbCrLf
        
        MsgBox "Cannot establish connection." & vbCr & vbCr & _
               "Check your Internet connection and try again.", _
               vbExclamation, _
               "Cannot Establish Connection"
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        MsgBox "The server returned the following response:" & vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Download Error"
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
strHeader = Inet1.GetHeader("Content-Length")
intKFileLength = CInt(Val(strHeader) / 1024)
If intKFileLength = 0 Then
    ' Failed; File length would never be 0!
    GoTo ExitDownload
End If

' Check for available disk space first!
'
' Note: I have left out the DiskFreeSpace function to
' keep this sample simple and self-contained in a single
' form.
'
' If you want to include it, put the following section
' into a Module...
'
' ------------------- section for public module -------------------
'****************************************************************
'Windows API/Global Declarations for :FreeDiskSpace
'****************************************************************
'Declare Function GetDiskFreeSpace Lib "kernel32" _
                 Alias "GetDiskFreeSpaceA" _
                 (ByVal lpRootPathName As String, _
                 lpSectorsPerCluster As Long, _
                 lpBytesPerSector As Long, _
                 lpNumberOfFreeClusters As Long, _
                 lpTotalNumberOfClusters As Long) As Long

'Public Function DiskFreeSpace(strDrive As String) As Double

' DiskFreeSpace:    returns the amount of free space on a drive
'                   in Windows9x/2000/NT4+
'Dim SectorsPerCluster As Long
'Dim BytesPerSector As Long
'Dim NumberOfFreeClusters As Long
'Dim TotalNumberOfClusters As Long
'Dim FreeBytes As Long
'Dim spaceInt As Integer
'strDrive = QualifyPath(strDrive)
' Call the API function
'GetDiskFreeSpace strDrive, _
                 SectorsPerCluster, _
                 BytesPerSector, _
                 NumberOFreeClusters, _
                 TotalNumberOfClusters

' Calculate the number of free bytes
'DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector
'End Function
' ------------------- end section for public module -------------------

' ...and un-comment the following 8 lines

'If DiskFreeSpace(Left(strDestination, InStr(strDestination, "\"))) / 1024 < intKFileLength Then
'    ' Not enough free space to download file
'    MsgBox "There is not enough free space on disk for this file!" _
'    & vbCr & vbCr & "Please free up some disk space and try again.", _
'    vbCritical, _
'    "Insufficient Disk Space"
'    GoTo ExitDownload
'End If

' Prepare display
pbar.Value = 0
pbar.Max = intKFileLength
DoEvents

intKReceived = 0

On Local Error GoTo FileErrorHandler

' If no errors occurred, then spank the file to disk
If Inet1.ResponseCode = 0 Then
    intFile = FreeFile()        ' Set intFile to an unused file.
    ' Open a file to write to.
    Open strDestination For Binary Access Write As #intFile
    ' Get the first chunk.
    bData = Inet1.GetChunk(1024, icByteArray)
    a = bData                   ' Must assign array to ANOTHER var cus
                                ' VB has a cow with LenB(bData)!
    
    Do While LenB(a) > 0        ' while there's still data...
        Put #intFile, , bData   ' Put it into our destination file
        ' Get next chunk.
        bData = Inet1.GetChunk(1024, icByteArray)
        a = bData
        If CancelSearch Then
            Close #intFile
            Kill strDestination
            GoTo ExitDownload
        End If
        intKReceived = intKReceived + 1
        If intKReceived < intKFileLength Then   ' to avoid -1's
            sglRate = intKReceived / (Timer - lastTime)
            sglTime = (intKFileLength - intKReceived) / sglRate
            timeleft.Caption = "Estimated Time Left: " & _
                           FormatTime(sglTime) & _
                           " (" & _
                           FormatFileSize(intKReceived * 1024#) & _
                           " of " & _
                           FormatFileSize(intKFileLength * 1024#) & _
                           " copied)" & vbCr & vbCr & _
                           "Transfer Rate: " & _
                           Format(sglRate, "###,##0.0") & " KB/Sec"
            pbar.Value = intKReceived
            Caption = Format((intKReceived / intKFileLength), "##0%") & _
                      " of " & strFile & " Completed"
        End If
    Loop
    Put #intFile, , bData
    Close #intFile
End If

StatusLabel2 = Empty
DoEvents

ExitDownload:
If intKReceived >= intKFileLength Then
    StatusLabel = "Download completed!"
    DownloadSuccess = True
    pbar.Value = pbar.Max
Else
    ' Delete partially downloaded file, if it exists
    If Not Dir(strDestination) = Empty Then Kill strDestination
    If Not CancelSearch Then
        StatusLabel = "Download failed!"
        MsgBox "Download failed!", _
        vbCritical, _
        "Error Downloading File"
    End If
End If

' Make sure that the Internet connection is closed
Inet1.Cancel
DoEvents
' and exit this function
'Unload Me
DoEvents
Exit Function

InternetErrorHandler:
    CancelSearch = True
    Inet1.Cancel
    MsgBox "Error: " & Err.Description & " occurred.", _
           vbCritical, _
           "Error Downloading File"
    DoEvents
    Resume Next
    
FileErrorHandler:
    MsgBox "Cannot write file to disk!", _
           vbCritical, _
           "Error Downloading File"
    Resume Next
    
End Function


Private Sub UserControl_Terminate()
  CancelSearch = True
End Sub
