Attribute VB_Name = "modUnACE"
Option Explicit

' Title:    UnACE
' Author:   Leigh Bowers
' Email:    compulsion@esheep.freeserve.co.uk
' WWW:      http://www.esheep.freeserve.co.uk/compulsion
' Version:  1.0 *Standard*
' Date:     15th June 1999
' Requires: UnACE.DLL
' License:  Freely Distributable (non-commercial use)

' Notes:    This is the Standard version.
'           The Pro (FULL) release can do selective
'           extracts (single files), is more configurable,
'           enhanced error control, retrieve archive info,
'           more methods and many extras...

' Constants...

Private Const ACEERR_MEM As Byte = 1
Private Const ACEERR_FILES As Byte = 2
Private Const ACEERR_FOUND As Byte = 3
Private Const ACEERR_FULL As Byte = 4
Private Const ACEERR_OPEN As Byte = 5
Private Const ACEERR_READ As Byte = 6
Private Const ACEERR_WRITE As Byte = 7
Private Const ACEERR_CLINE As Byte = 8
Private Const ACEERR_CRC As Byte = 9
Private Const ACEERR_OTHER As Byte = 10
Private Const ACEERR_EXISTS As Byte = 11
Private Const ACEERR_END As Byte = 128
Private Const ACEERR_HANDLE As Byte = 129
Private Const ACEERR_CONSTANT As Byte = 130
Private Const ACEERR_NOPASSW As Byte = 131
Private Const ACEERR_METHOD As Byte = 132
Private Const ACEERR_USER As Byte = 255

Private Const ACEOPEN_EXTRACT As Byte = 1

Private Const ACECMD_EXTRACT As Byte = 2

' Types...

Private Type ACEOpenArchiveData
    ArcName As String
    OpenMode As Long
    OpenResult As Long
    Flags As Long
    Host As Long
    AV As String * 51
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type

Private Type ACEHeaderData
  ArcName As String * 260
  FileName As String * 260
  Flags As Long
  PackSize As Long
  UnpSize As Long
  FileCRC As Long
  FileTime As Long
  Method As Long
  QUAL As Long
  FileAttr As Long
  CmtBuf As String
  CmtBufSize As Long
  CmtSize As Long
  CmtState As Long
End Type

' API's...

Public Declare Function ACEOpen Lib "UnACE.dll" Alias "ACEOpenArchive" (ByRef ACEOpenData As ACEOpenArchiveData) As Long
Public Declare Function ACEClose Lib "UnACE.dll" Alias "ACECloseArchive" (ByVal HandleToArchive As Long) As Long
Public Declare Function ACEReadHeader Lib "UnACE.dll" (ByVal HandleToArchive As Long, ByRef ACEHeaderRead As ACEHeaderData) As Long
Public Declare Function ACEProcessFile Lib "UnACE.dll" (ByVal HandleToArchive As Long, ByVal Operation As Long, ByVal DestPath As String) As Long
Public Declare Function ACESetPassword Lib "UnACE.dll" (ByVal HandleToArchive As Long, ByVal Password As String) As Long

Public Function ACEExtract(ByVal sACEArchive As String, ByVal sDestPath As String, Optional ByVal sPassword As String) As Long

    ' ACE Extract (Archive, Destination, <Password>)

Dim lHandle As Long
Dim lStatus As Long
Dim uACE As ACEOpenArchiveData
Dim uHeader As ACEHeaderData

Dim lFileCount As Long

    ' Attempt to open the archive...

    uACE.ArcName = sACEArchive
    uACE.OpenMode = ACEOPEN_EXTRACT
    lHandle = ACEOpen(uACE)
    
    ' Success?
    
    If uACE.OpenResult <> 0 Then Exit Function
    
    ' Password protected?
    
    If sPassword <> "" Then
        ACESetPassword lHandle, sPassword
    End If
    
    lFileCount = 0
    
    ' Extract files from archive...
    
    lStatus = ACEReadHeader(lHandle, uHeader)
    Do Until lStatus <> 0
        If ACEProcessFile(lHandle, ACECMD_EXTRACT, sDestPath + IIf(Right(sDestPath, 1) = "\", "", "\") + uHeader.FileName) = 0 Then
            lFileCount = lFileCount + 1
        End If
        lStatus = ACEReadHeader(lHandle, uHeader)
    Loop
    
    ' Close the archive...
    
    ACEClose lHandle
    
    ' Returns...
    '           0 = Failed or No-files to extract
    '          >0 = Number of files extracted
    
    ACEExtract = lFileCount

End Function


