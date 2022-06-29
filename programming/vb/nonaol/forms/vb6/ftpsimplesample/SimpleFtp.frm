VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple FTP"
   ClientHeight    =   4635
   ClientLeft      =   3045
   ClientTop       =   1410
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton butClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Semantec"
      Height          =   852
      Left            =   3000
      TabIndex        =   22
      Top             =   2640
      Width           =   2172
      Begin VB.OptionButton Option4 
         Caption         =   "Passive"
         Height          =   252
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1092
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Active (default)"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Dir"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   720
      Width           =   1212
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Binary"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   2880
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ASCII"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   3120
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SetCurrentDir"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton Command6 
      Caption         =   "GetCurrentDir"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   1212
   End
   Begin VB.CommandButton Command5 
      Caption         =   "GetFile"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   2640
      Width           =   1212
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Put Large File "
      Height          =   372
      Left            =   5280
      TabIndex        =   11
      Top             =   3120
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
      Height          =   372
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   372
      Left            =   5280
      TabIndex        =   12
      Top             =   3600
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PutFile"
      Height          =   372
      Left            =   5280
      TabIndex        =   9
      Top             =   2160
      Width           =   1212
   End
   Begin VB.TextBox Text5 
      Height          =   372
      Left            =   960
      TabIndex        =   6
      Top             =   2160
      Width           =   4212
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   4212
   End
   Begin VB.TextBox Text3 
      Height          =   372
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "user@domain.com"
      Top             =   1200
      Width           =   4212
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Text            =   "anonymous"
      Top             =   720
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   960
      TabIndex        =   0
      Text            =   "ftp.microsoft.com"
      Top             =   240
      Width           =   4212
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transfer Type"
      Height          =   852
      Left            =   960
      TabIndex        =   25
      Top             =   2640
      Width           =   1932
   End
   Begin VB.Label Label7 
      Height          =   372
      Left            =   2760
      TabIndex        =   19
      Top             =   3600
      Width           =   2172
   End
   Begin VB.Label Label6 
      Caption         =   "Large Transfer   Byte Count:"
      Height          =   372
      Left            =   960
      TabIndex        =   18
      Top             =   3600
      Width           =   1212
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Remote File or Directory"
      Height          =   612
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   732
   End
   Begin VB.Label Label4 
      Caption         =   "Local File:"
      Height          =   372
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   372
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      Height          =   252
      Left            =   480
      TabIndex        =   14
      Top             =   720
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   252
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hOpen As Long, hConnection As Long, hFile As Long
Dim dwType As Long
Dim dwSeman As Long

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
Dim dwRet As Long
Dim dwTemp As Long
Dim szString As String * 2048
Dim szErrorMessage As String

dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                  GetModuleHandle("wininet.dll"), dwError, 0, _
                  szString, 256, 0)
szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
Debug.Print szErrorMessage
MsgBox szErrorMessage, , "SimpleFtp"
If (dwError = 12003) Then
    ' Extended error information was returned
    dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
    Debug.Print szString
    Form2.Show
    Form2.Text1.Text = szString
End If
End Sub

Private Sub butClose_Click()
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = 0
    MsgBox "Disconnected.", , "SimpleFtp"
End Sub

Private Sub Command1_Click()
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
        If (FtpPutFile(hConnection, Text4.Text, Text5.Text, _
         dwType, 0) = False) Then
             ErrorOut Err.LastDllError, "FtpPutFile"
             Exit Sub
        Else
         MsgBox "File transfered!", , "Simple Ftp"
        End If

End Sub

Private Sub Command2_Click()
If (FtpDeleteFile(hConnection, Text5.Text) = False) Then
    MsgBox "FtpDeleteFile error: " & Err.LastDllError
            Exit Sub
Else
     MsgBox "File deleted!"
End If
End Sub

Private Sub Command3_Click()
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = InternetConnect(hOpen, Text1.Text, INTERNET_INVALID_PORT_NUMBER, _
    Text2.Text, Text3.Text, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        Exit Sub
    Else
        MsgBox "Connected!", , "SimpleFtp"
        Option3.Enabled = False
        Option4.Enabled = False
    End If
        

End Sub

Private Sub Command4_Click()
'&H40000000 == GENERIC_WRITE
Dim Data(99) As Byte ' array of 100 elements 0 to 99
Dim Written As Long
Dim Size As Long
Dim Sum As Long
Dim j As Long

Sum = 0
j = 0
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
hFile = FtpOpenFile(hConnection, Text5.Text, &H40000000, dwType, 0)
If hFile = 0 Then
    ErrorOut Err.LastDllError, "FtpOpenFile"
    Exit Sub
End If
Open Text4.Text For Binary Access Read As #1
Size = LOF(1)
For j = 1 To Size \ 100
    Get #1, , Data
    If (InternetWriteFile(hFile, Data(0), 100, Written) = 0) Then
        ErrorOut Err.LastDllError, "InternetWriteFile"
        Exit Sub
    End If
    DoEvents
    Sum = Sum + 100
    Label7.Caption = Str(Sum)
Next j
Get #1, , Data
 If (InternetWriteFile(hFile, Data(0), Size Mod 100, Written) = 0) Then
        ErrorOut Err.LastDllError, "InternetWriteFile"
        Exit Sub
End If
Sum = Sum + (Size Mod 100)
Label7.Caption = Str(Sum)
Close #1
InternetCloseHandle (hFile)
End Sub

Private Sub Command5_Click()
' for ASCII files use FTP_TRANSFER_TYPE_ASCII
' add INTERNET_FLAG_NO_CACHE_WRITE to avoid local caching 0x04000000 (hex)
 If (FtpGetFile(hConnection, Text5.Text, Text4.Text, False, _
         FILE_ATTRIBUTE_NORMAL, dwType Or INTERNET_FLAG_RELOAD, 0) = False) Then
    ErrorOut Err.LastDllError, "FtpPutFile"
    Exit Sub
 Else
   MsgBox "File transfered!", , "SimpleFtp"
   End If
End Sub

Private Sub Command6_Click()
Dim szDir As String

szDir = String(1024, Chr$(0))

If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
    ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
    Exit Sub
 Else
   MsgBox "Current directory is: " & szDir, , "SimpleFtp"
   End If
End Sub

Private Sub Command7_Click()
If (FtpSetCurrentDirectory(hConnection, Text5.Text) = False) Then
   ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
   Exit Sub
Else
  MsgBox "Directory is changed to " & Text5.Text, , "SimpleFtp"
End If

End Sub

Private Sub Command8_Click()
Dim szDir As String
Dim hFind As Long
Dim nLastError As Long
Dim dError As Long
Dim ptr As Long
Dim pData As WIN32_FIND_DATA
    

hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
nLastError = Err.LastDllError
If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            MsgBox "This directory is empty!", , "SimpleFtp"
        Else
            ErrorOut Err.LastDllError, "FtpFindFirstFile"
        End If
        Exit Sub
End If

dError = NO_ERROR
     Dim bRet As Boolean
 
szDir = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1) & " " & Win32ToVbTime(pData.ftLastWriteTime)
szDir = szDir & vbCrLf
Do
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        If Not bRet Then
            dError = Err.LastDllError
            If dError = ERROR_NO_MORE_FILES Then
                Exit Do
            Else
                ErrorOut Err.LastDllError, "InternetFindNextFile"
                InternetCloseHandle (hFind)
                Exit Sub
            End If
        Else
            
            szDir = szDir & Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1) & " " & Win32ToVbTime(pData.ftLastWriteTime) & vbCrLf
        End If
Loop
   
Dim szTemp As String
szTemp = String(1024, Chr$(0))
If (FtpGetCurrentDirectory(hConnection, szTemp, 1024) = False) Then
    ErrorOut Err.LastDllError, "FtpGetCurrentDirectory"
    Exit Sub
End If
MsgBox szDir, , "Directory Listing of: " & szTemp
InternetCloseHandle (hFind)
End Sub

Private Sub Form_Load()
  hOpen = InternetOpen("My VB Test", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If hOpen = 0 Then
    ErrorOut Err.LastDllError, "InternetOpen"
    Unload Form1
  End If
  dwType = FTP_TRANSFER_TYPE_ASCII
  dwSeman = 0
  hConnection = 0
End Sub
Private Sub Form_Uload()
    InternetCloseHandle hOpen
Unload Form2
End Sub

Private Sub Option1_Click()
dwType = FTP_TRANSFER_TYPE_ASCII
End Sub

Private Sub Option2_Click()
dwType = FTP_TRANSFER_TYPE_BINARY
End Sub

Private Sub Option3_Click()
dwSeman = 0
End Sub

Private Sub Option4_Click()
dwSeman = INTERNET_FLAG_PASSIVE
End Sub
