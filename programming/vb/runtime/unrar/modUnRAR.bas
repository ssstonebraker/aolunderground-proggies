Attribute VB_Name = "modUnRAR"
Option Explicit

' Title:    UnRAR VB6 BAS Module
' Author:   Leigh Bowers (Compulsion Software)
' Email:    compulsion@esheep.freeserve.co.uk
' Version:  1.01
' Date:     16th April 1999
' Revised:  15th June 1999
' Requires: UnRAR.DLL (supplied)
' License:  Freely Distributable

' Notes:-

' Please give me a small mention if you use all or part
' of this code in one of your VB programs ;-)

' Created utilising information contained within the
' UnRARDLL.txt & TechNote.txt files (as supplied with
' the full WinRAR package).

' Open Mode Constants

Private Const RAR_OM_LIST As Byte = 0
Private Const RAR_OM_EXTRACT As Byte = 1

' Error Constants

Private Const ERAR_NO_MEMORY As Byte = 11
Private Const ERAR_BAD_DATA As Byte = 12
Private Const ERAR_BAD_ARCHIVE As Byte = 13
Private Const ERAR_EOPEN As Byte = 15
Private Const ERAR_UNKNOWN_FORMAT As Byte = 14
Private Const ERAR_SMALL_BUF As Byte = 20
Private Const ERAR_ECLOSE As Byte = 17
Private Const ERAR_END_ARCHIVE As Byte = 10
Private Const ERAR_ECREATE As Byte = 16
Private Const ERAR_EREAD As Byte = 18
Private Const ERAR_EWRITE As Byte = 19

' Operation Constants

Private Const RAR_SKIP As Byte = 0
Private Const RAR_TEST As Byte = 1
Private Const RAR_EXTRACT As Byte = 2

' Volume Constants

Private Const RAR_VOL_ASK As Byte = 0
Private Const RAR_VOL_NOTIFY As Byte = 1

' User Defined Types

Private Type RARHeaderData
    ArcName As String * 260
    FileName As String * 260
    Flags As Long
    PackSize As Long
    UnpSize As Long
    HostOS As Long
    FileCRC As Long
    FileTime As Long
    UnpVer As Long
    Method As Long
    FileAttr As Long
    CmtBuf As String    ' Pointer (char *CmtBuf in C)
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type

Private Type RAROpenArchiveData
    ArcName As String   ' Pointer (char *ArcName in C)
    OpenMode As Long
    OpenResult As Long
    CmtBuf As String    ' Pointer (char *CmtBuf in C)
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type

' RAR DLL Declares

Public Declare Function RAROpen Lib "UnRAR.dll" Alias "RAROpenArchive" (ByRef RAROpenData As RAROpenArchiveData) As Long
Public Declare Function RARClose Lib "UnRAR.dll" Alias "RARCloseArchive" (ByVal HandleToArchive As Long) As Long
Public Declare Function RARReadHdr Lib "UnRAR.dll" Alias "RARReadHeader" (ByVal HandleToArcRecord As Long, ByRef ArcHeaderRead As RARHeaderData) As Long
Public Declare Function RARProcFile Lib "UnRAR.dll" Alias "RARProcessFile" (ByVal HandleToArcHeader As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long
Public Declare Sub RARSetChangeVolProc Lib "UnRAR.dll" (ByVal HandleToArchive As Long, ByVal Mode As Long)
Public Declare Sub RARSetPassword Lib "UnRAR.dll" (ByVal HandleToArchive As Long, ByVal Password As String)

Function RARExtract(ByVal sRARArchive As String, ByVal sDestPath As String, Optional ByVal sPassword As String) As Integer

' Description:-
' Exrtact file(s) from RAR archive.

' Parameters:-
' sRARArchive   = RAR Archive filename
' sDestPath     = Destination path for extracted file(s)
' sPassword     = Password [OPTIONAL]

' Returns:-
' Integer       = 0  Failed (no files, incorrect PW etc)
'                 -1 Failed to open RAR archive
'                 >0 Number of files extracted
    
Dim lHandle As Long
Dim lStatus As Long
Dim uRAR As RAROpenArchiveData
Dim uHeader As RARHeaderData
Dim iFileCount As Integer
    
    RARExtract = -1
    
    ' Open the RAR

    uRAR.ArcName = sRARArchive
    uRAR.OpenMode = RAR_OM_EXTRACT
    lHandle = RAROpen(uRAR)

    ' Failed to open RAR ?

    If uRAR.OpenResult <> 0 Then Exit Function
    
    ' Password ?
    
    If sPassword <> "" Then
        RARSetPassword lHandle, sPassword
    End If
    
    ' Extract file(s)...
    
    iFileCount = 0

    ' Is there at lease one archived file to extract ?
    lStatus = RARReadHdr(lHandle, uHeader)

    Do Until lStatus <> 0

        ' Process (extract) the current file within the archive
        If RARProcFile(lHandle, RAR_EXTRACT, "", sDestPath + uHeader.FileName) = 0 Then
            iFileCount = iFileCount + 1
        End If
        ' Is there another archived file in this RAR ?
        lStatus = RARReadHdr(lHandle, uHeader)

    Loop

    ' Close the RAR

    RARClose lHandle

    ' Return

    RARExtract = iFileCount

End Function

