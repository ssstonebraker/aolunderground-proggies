Attribute VB_Name = "MainModule"
#If Win16 Then
    Declare Function WritePrivateProfileString Lib "KERNEL" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$) As Integer
    Declare Function GetPrivateProfileString Lib "KERNEL" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$) As Integer
#Else
    Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$) As Long
    Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Long, ByVal FileName$) As Long
#End If

Sub Main()
Dim ReturnString As String
'--- Check to see if we are in the VB.INI File.  If not, Add ourselves to the INI file
    #If Win16 Then
        Section$ = "Add-Ins16"
    #Else
        Section$ = "Add-Ins32"
    #End If
    
    'Check to see if the Align.Connector entry is already in the VB.INI file.  Add if not.
    ReturnString = String$(255, Chr$(0))
    GetPrivateProfileString Section$, "Align.Connector", "NotFound", ReturnString, Len(ReturnString) + 1, "VB.INI"
    If Left(ReturnString, InStr(ReturnString, Chr(0)) - 1) = "NotFound" Then
        WritePrivateProfileString Section$, "Align.Connector", "0", "VB.INI"
    End If
End Sub

