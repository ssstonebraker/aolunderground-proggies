Attribute VB_Name = "Registry_reader"
'Bas Made by MaRZ © 98
'Email: MaRZ001@Juno.COM
'This bas will read form the System Registry.
'There are a couple of examples,Get Windows
'Version, Get Windows Bit, Ger Printer Name,
'and others.
Option Explicit
'---------------------------------------------------------------
'-Registry API Declarations...
'---------------------------------------------------------------
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'---------------------------------------------------------------
'- Registry Api Constants...
'---------------------------------------------------------------
' Reg Key ROOT Types...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG As Long = &H80000005
Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long
Dim R As Long
RetVal$ = ""
Const KEY_ALL_ACCESS As Long = &HF0063
Const ERROR_SUCCESS As Long = 0
Const REG_SZ As Long = 1
R = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
If R <> ERROR_SUCCESS Then GoTo Quit_Now
SZ = 256: V$ = String$(SZ, 0)
R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal V$, SZ)
If R = ERROR_SUCCESS And dwType = REG_SZ Then
RetVal$ = Left$(V$, SZ)
Else
RetVal$ = Left$(V$, SZ)
End If
If hInKey = 0 Then R = RegCloseKey(hSubKey)
Quit_Now:
RegGetString$ = RetVal$
End Function
Function GetCurrPrinter() As String
'Gets current Printer Driver name from the System Registry
'Use:
'Make 1 Button
'Make 1 Textbox
'Add the following code to the Click event for Command1:
'Dim PName As String
'PName = GetCurrPrinter()
'Text1="Current Printer: " & PName
GetCurrPrinter = RegGetString$(HKEY_CURRENT_CONFIG, "System\CurrentControlSet\Control\Print\Printers", "Default")
End Function
Function GetBit() As String
'Gets current Windows Bit(ie 16,24,32) name from the System Registry
'Use:
'Make 1 Button
'Make 1 Textbox
'Add the following code to the Click event for Command1:
'Dim PName As String
'PName = GetBit()
'Text1="Current Bit: " & PName
GetBit = RegGetString$(HKEY_CURRENT_CONFIG, "Display\Settings", "BitsPerPixel")
End Function
Function GetWVer() As String
'Gets current Windows Version along with the Sub Version name from the System Registry
'Use:
'Make 1 Button
'Make 1 Textbox
'Add the following code to the Click event for Command1:
'Dim PName As String
'Dim PName2 As String
'PName = GetWVer()
'PName2 = GetWSubVer()
'Text1="Current Bit: " & PName
'Text1 = Text1 + PName2
GetWVer = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "VersionNumber")
End Function
Function GetWSubVer() As String
GetWSubVer = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "SubVersionNumber")
End Function
Function GetWUsernm() As String
'Gets current Windows User's Name from the System Registry
'Use:
'Make 1 Button
'Make 1 Textbox
'Add the following code to the Click event for Command1:
'Dim PName As String
'PName = GetWUsernm()
'Text1="Current Windows User's Name: " & PName
GetWUsernm = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
End Function
Function GetWVer2() As String
'Gets current Windows Version (ie 95,98) name from the System Registry
'Use:
'Make 1 Button
'Make 1 Textbox
'Add the following code to the Click event for Command1:
'Dim PName As String
'PName = GetWVer2()
'Text1="Current Windows Version: " & PName
GetWVer2 = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductName")
End Function
