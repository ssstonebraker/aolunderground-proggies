VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Control panel ""sound"" example by Þútõ²"
   ClientHeight    =   4695
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "REGISTRY.frx":0000
      ToolTipText     =   "My info"
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long ' for the sound to play
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_REGISTRY_RECOVERED = 1014&
Const ERROR_REGISTRY_CORRUPT = 1015&
Const ERROR_REGISTRY_IO_FAILED = 1016&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_NOSTOP = &H10
Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Private Sub form_load()
Dim retValue As Long
Dim Result As Long
Dim keyID As Long
Dim KeyValue As String
Dim subKey As String
Dim bufSize As Long
Dim regkey As String
regkey = "\AppEvents\Schemes\Apps\KnK\Startup\.current"
retValue = RegCreateKey(HKEY_CURRENT_USER, regkey, keyID)
    If retValue = 0 Then 'if i dont find the registry file
        'Create key if none exists, prevents errors.
        'NOTE: notice subkey = ""
        'this is because it does not have
        'a clue where the user has the file
        subKey = ""
        retValue = RegQueryValueEx(keyID, subKey, 0&, REG_SZ, _
                   0&, bufSize)
        'No value, set it
        If bufSize > 2 Then 'A little trickey, if data
                            ' is greater than 2 (from above) then it:
        
             KeyValue = String(bufSize + 1, " ") ' i wouldent touch this, i kept getting an "out of memory" error when i did. wondered why. This seems to solve it.
             retValue = RegQueryValueEx(keyID, subKey, 0&, REG_SZ, _
                      ByVal KeyValue, bufSize) 'FINALLY gets the data, it's safe to pass on
             Call sndPlaySound(KeyValue, SND_FLAG)
        End If
End If
'in my opinion, this is one of the best ways to
'acess the registry. It will allow you to get
'any value in any area of the registry you need.
'Jus my opinion though
'NOTE: i know i could done a lot more optomizing
'and maby removed some shit, or made it into a function
'i dident tho because this works plenty fast enough
'and was fine by itself
End Sub





Private Sub Form_Unload(Cancel As Integer)
'same as above, except you change the regkey to make it work
Dim retValue As Long
Dim Result As Long
Dim keyID As Long
Dim KeyValue As String
Dim subKey As String
Dim bufSize As Long
Dim regkey As String

    regkey = "\AppEvents\Schemes\Apps\KnK\Shutdown\.current"
    retValue = RegCreateKey(HKEY_CURRENT_USER, regkey, keyID)
    If retValue = 0 Then
        subKey = ""
        retValue = RegQueryValueEx(keyID, subKey, 0&, REG_SZ, _
                   0&, bufSize)
        If bufSize > 2 Then
    
        
             KeyValue = String(bufSize + 1, " ")
             retValue = RegQueryValueEx(keyID, subKey, 0&, REG_SZ, _
                      ByVal KeyValue, bufSize)
                      Call sndPlaySound(KeyValue, SND_NOSTOP)
        End If
End If

End Sub


