Attribute VB_Name = "modIni"
'Welcome to the ViViD INI Reader and Writer.
'Coded by: ViViDVaSt(ViViDVaST@hotmail.com)
'1/13/01


Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal sSectionName As String, ByVal sKeyName As String, ByVal sString As String, ByVal sFileName As String) As Long

'*******************************************************************
'This sub writes to an INI file with any 3 letter
'extention name.
'*******************************************************************
Sub writetoini(ByVal strSectionName As String, ByVal strHeadingName As String, ByVal strValue As String, ByVal strpath As String)

Dim strTotalPath As String
Dim intChar As Integer
Dim intdistance As Integer
Dim intLength As Integer



intLength = Len(strpath)

For intChar = 1 To intLength
If Mid(strpath, intChar, 1) = "\" Then
intdistance = intChar
On Error Resume Next
strTotalPath = Mid(strpath, 1, intdistance)
MkDir (strTotalPath)
End If
Next intChar
lReturn = WritePrivateProfileString(strSectionName, strHeadingName, strValue, strpath)
'**********************************
End Sub
'*******************************************************************

'*******************************************************************
'This sub reads from an INI file with any 3 letter
'extention name.
'*******************************************************************
Function ReadFromIni(ByVal strSubSection As String, ByVal strHeading As String, ByVal strPathName As String) As String

 Dim RetStr As String
   RetStr = String(255, Chr(0))
   ReadFromIni = Left(RetStr, GetPrivateProfileString(strSubSection, ByVal strHeading, "", RetStr, Len(RetStr), strPathName))
End Function
'*******************************************************************



