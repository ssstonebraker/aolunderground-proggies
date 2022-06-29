VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PC Development Forum (keyword: PDV)"
   ClientHeight    =   2580
   ClientLeft      =   3660
   ClientTop       =   3105
   ClientWidth     =   4425
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   1605
      TabIndex        =   3
      Top             =   2175
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0442
      Height          =   1050
      Left            =   60
      TabIndex        =   4
      Top             =   1080
      Width           =   4320
   End
   Begin VB.Label lblQuestions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Questions can be sent to pccmiked@aol.com"
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   795
      Width           =   4320
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Written by PCC MikeD"
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   585
      Width           =   4320
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subclassing Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   4320
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown the computer"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart the computer"
      End
      Begin VB.Menu mnuRestartDOS 
         Caption         =   "Restart in MS-DOS mode"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OsVersionInfo
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type

Private OsVer As OsVersionInfo

'Platform constants
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetVersionEx Lib "kernel32.dll" _
Alias "GetVersionExA" (lpStruct As OsVersionInfo) As Long

'Functions for getting the Windows and Windows\System directories
Private Declare Function GetWindowsDirectory Lib "kernel32" _
Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long

Const TASKBARICONID = 1

Dim msWinDir As String
Dim msSysDir As String

Private Sub cmdOK_Click()

Me.Hide

End Sub


Private Sub Form_Load()

Dim lBytes As Long
Dim lRet As Long
Dim iVersion As Integer
Dim sTip As String
Dim sBuffer As String
Dim lBufferSize As Long

'Determine the version of Windows that is running
OsVer.dwVersionInfoSize = 148&
lRet = GetVersionEx(OsVer)
Select Case OsVer.dwPlatform
    Case VER_PLATFORM_WIN32_NT
        iVersion = (OsVer.dwMajorVersion * 100) + OsVer.dwMinorVersion
        If iVersion < 351 Then
            MsgBox "This demo is not compatible with Windows NT 3.50 and earlier."
            Unload Me
            Exit Sub
        End If
    Case VER_PLATFORM_WIN32_WINDOWS
        'Running Windows 95 so no problem
    Case Else
        'Should only occur if Win32s is installed on Win 3.1, but the
        'program still won't work because there will be no task bar.
        MsgBox "This demo can only run under Windows 95 or Windows NT 3.51."
        Unload Me
        Exit Sub
End Select
 
'Check for presence of Taskbar.  The user may have a different
'shell that doesn't support one.
If TaskBar.Exists = False Then
    MsgBox "There is no taskbar currently available."
    Unload Me
    Exit Sub
End If

'Enable subclassing of the form
If Hook(Me) = False Then
    MsgBox "Error initializing program."
    Unload Me
    Exit Sub
End If

'Assign the form's hWnd to the taskbar notify structure
TaskBar.hwnd = Me

'Add icon to system tray
TaskBar.AddIcon Me.Icon, TASKBARICONID

'Assign a tooltip for the icon
sTip = "Double-click exits Windows; Right-click displays menu"
TaskBar.ChangeTip sTip, TASKBARICONID

'Get the Windows and Windows\System directories
sBuffer = Space(145)
lBufferSize = Len(sBuffer)
lBytes = GetWindowsDirectory(sBuffer, lBufferSize)
msWinDir = Left$(sBuffer, lBytes)

sBuffer = Space(145)
lBytes = GetSystemDirectory(sBuffer, lBufferSize)
msSysDir = Left$(sBuffer, lBytes)

End Sub


Private Sub Form_Unload(Cancel As Integer)

'Remove the icon from the task bar
TaskBar.DeleteIcon (TASKBARICONID)

Call Unhook(Me)

Set frmMain = Nothing

End Sub


Private Sub mnuAbout_Click()

Me.Show

End Sub


Private Sub mnuClose_Click()

Unload Me

End Sub


Private Sub mnuRestart_Click()

Dim lResult As Long

lResult = ExitWindowsEx(EWX_REBOOT, 0&)

End Sub

Private Sub mnuRestartDOS_Click()

'Shut down the computer to MS-DOS using the "Exit To Dos.Pif" file

Dim lResult As Long
Dim sExitToDosPif As String
Dim sMsg As String

'First, let's make sure the pif file exists
sExitToDosPif = msWinDir & "\exit to dos.pif"
If Not FileExists(sExitToDosPif) Then
    'Not in Windows directory.  Check Windows\Pif
    sExitToDosPif = msWinDir & "\pif\exit to dos.pif"
    If Not FileExists(sExitToDosPif) Then
        'Still can't find it.  One last check in the System directory
        sExitToDosPif = msSysDir & "\exit to dos.pif"
        If Not FileExists(sExitToDosPif) Then
            'Time to give up
            sMsg = "Can't find the ""Exit To DOS.pif"" file." & vbCrLf & vbCrLf _
                 & "Aborting procedure." & vbCrLf & vbCrLf _
                & "Windows can automatically create this file for you.  " _
                & "Simply click the Start button, select Shut Down, and " _
                & "choose the ""Restart the computer in MS-DOS mode"" option."
            MsgBox sMsg, vbCritical
            Exit Sub
        End If
    End If
End If

lResult = Shell(sExitToDosPif, 1)

End Sub


Private Sub mnuShutdown_Click()

Dim lResult As Long

lResult = StandardShutdown

End Sub


