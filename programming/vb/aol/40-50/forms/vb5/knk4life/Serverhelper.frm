VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KnK 4 Life Server Helper                       By: KnK"
   ClientHeight    =   6510
   ClientLeft      =   2670
   ClientTop       =   945
   ClientWidth     =   9480
   Icon            =   "Serverhelper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command26 
      Caption         =   "Find"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Send Thanx"
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Send Status"
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Send List"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   480
      Width           =   1335
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   4320
      TabIndex        =   53
      Top             =   4800
      Width           =   1575
      Begin VB.CommandButton Command21 
         Caption         =   "<-----------"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Stop"
         Height          =   255
         Left            =   600
         TabIndex        =   56
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   495
      End
      Begin VB.ListBox List5 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Reset #5"
      Height          =   255
      Left            =   2160
      TabIndex        =   50
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Reset #4"
      Height          =   255
      Left            =   1320
      TabIndex        =   49
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   48
      Top             =   4200
      Width           =   375
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Number5"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   6000
      TabIndex        =   35
      Top             =   3600
      Width           =   3375
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   840
         TabIndex        =   41
         Text            =   "Empty"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Text            =   "Empty"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Text            =   "Empty"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Text            =   "Empty"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Text            =   "Send "
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Send Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Send Thanx"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Find X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   44
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Send"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   43
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Send List"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Apply/Save "
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Number4"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   6000
      TabIndex        =   21
      Top             =   1560
      Width           =   3375
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Text            =   "Send "
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Text            =   "Empty"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   "Empty"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "Empty"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "Empty"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Send List"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Send"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Find X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Send Thanx"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Send Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "#5"
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "#4"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2175
      Left            =   6000
      TabIndex        =   16
      Top             =   -120
      Width           =   3480
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   468
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   285
         TabIndex        =   52
         Top             =   600
         Width           =   285
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2760
         Top             =   600
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Emergency Stop!!"
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Currently Running"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1815
      Left            =   3000
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
      Begin VB.CommandButton Command11 
         Caption         =   "<-----------"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Save"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "App Name"
         Top             =   960
         Width           =   1095
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   840
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1560
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
      Begin VB.CommandButton Command4 
         Caption         =   "<---------"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Room"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ListBox List4 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   4680
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   1230
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "KnK"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Bin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Request"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.Image imOn 
      Height          =   465
      Left            =   5400
      Picture         =   "Serverhelper.frx":030A
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imOff 
      Height          =   465
      Left            =   5640
      Picture         =   "Serverhelper.frx":0500
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu info 
         Caption         =   "Info"
      End
      Begin VB.Menu help 
         Caption         =   "Help"
      End
      Begin VB.Menu mailme 
         Caption         =   "&Mail me"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu min 
         Caption         =   "Minimize"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Toolz"
      Begin VB.Menu im 
         Caption         =   "IM's"
         Begin VB.Menu on 
            Caption         =   "&On"
         End
         Begin VB.Menu off 
            Caption         =   "&Off"
         End
      End
      Begin VB.Menu bust 
         Caption         =   "&Room Buster"
      End
      Begin VB.Menu fb 
         Caption         =   "Find Bin"
         Begin VB.Menu sm 
            Caption         =   "Setup/Modify"
         End
         Begin VB.Menu activate 
            Caption         =   "Activate"
         End
      End
      Begin VB.Menu auto 
         Caption         =   "&Auto Find"
      End
   End
   Begin VB.Menu option 
      Caption         =   "&Options"
      Begin VB.Menu servermethod 
         Caption         =   "&Server M&ethods"
         Begin VB.Menu sandm 
            Caption         =   "Setup/Modify"
         End
         Begin VB.Menu line 
            Caption         =   "-"
         End
         Begin VB.Menu number1 
            Caption         =   "/SN Send x"
            Checked         =   -1  'True
         End
         Begin VB.Menu number2 
            Caption         =   "!SN send x"
         End
         Begin VB.Menu number3 
            Caption         =   "-sn send x"
         End
         Begin VB.Menu number4 
            Caption         =   "Empty"
         End
         Begin VB.Menu number5 
            Caption         =   "Empty"
         End
      End
      Begin VB.Menu clearlisrt 
         Caption         =   "&Clear L&ist"
         Begin VB.Menu cl 
            Caption         =   "Clear List "
         End
         Begin VB.Menu dcl 
            Caption         =   "Don't Clear List"
         End
      End
      Begin VB.Menu imonoff 
         Caption         =   "IM's on/off"
         Begin VB.Menu on2 
            Caption         =   "on"
         End
         Begin VB.Menu off2 
            Caption         =   "off"
         End
      End
   End
   Begin VB.Menu namess 
      Caption         =   "&Get N&ames"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu help2 
         Caption         =   "Help"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu max 
         Caption         =   "Maxamize"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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


Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim bOn As Boolean

Private Sub activate_Click()
If Text3 = "" Then
MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
Exit Sub
End If

If List3.ListCount = 0 Then
MsgBox "error: Nothing to request.", vbCritical, "error"
Exit Sub
End If


If off2.Checked = True Then
 Call InstantMessage("$IM_OFF", "=)")
Pause (2)
End If
Form1.Height = 2430
Call Setupform

'<-----------
ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>KnK 4 Life v2.o" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
Pause (0.8)
ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>Server Helper" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
Pause (0.8)

Timer3.Enabled = True
End Sub

Private Sub auto_Click()


Frame1.Top = 4800
Frame1.Left = 1920
Frame2.Top = 4800
Frame2.Left = 4320
Frame6.Top = 0
Frame6.Left = 3480
Form1.Width = 4875

End Sub

Private Sub bust_Click()
'On Local Error Resume Next
'Call ShellExecute(hwnd, "Open", "/C·RB.EXE", "", App.Path, 1)
'If Err Then
'MsgBox "Error: File Not Found!", vbExclamation, "Error"
'End If
Form2.Show
End Sub

Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If What_Said Like "/*" Then
    X$ = What_Said
    T$ = Left$(X$, 2)
    S$ = Right$(X$, 10)
    Y$ = Mid$(X$, 2, 10)

If List5.ListCount = 0 Then List5.AddItem Y$
    For i = 0 To List5.ListCount - 1
    num = List5.List(i)
If num = Y$ Then Exit Sub
    Next i
    List5.AddItem Y$
    ' List5.AddItem Y$
End If
End Sub

Private Sub cl_Click()
dcl.Checked = False
cl.Checked = True
SaveSetting "KnK4Life", "KnK_Clearlist", "yesorno", "yes"

End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
If Text3.Text = "" Then
    MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
    Exit Sub
End If

If List2.ListCount = 0 Then
    MsgBox "error: Nothing to request.", vbCritical, "error"
    Exit Sub
End If

If off2.Checked = True Then
    Call InstantMessage("$IM_OFF", "=)")
    Pause (2)
End If

Call Setupform
ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>KnK 4 Life v2.o" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
Pause (0.8)
ChatSend ("<FONT COLOR=#000000>«-×´¯`° <B><FONT COLOR=#FF0000>Server Helper" & aa$ & "</B><FONT COLOR=#000000>  °´¯`×-»")
Pause (0.8)
Timer2.Enabled = True

End Sub

Private Sub Command10_Click()
If List3.ListCount = 0 Then Exit Sub
    Dim a As Integer
    Dim b As Variant
    On Error GoTo C
    a = 2
    Open CStr(App.Path + "\find.lst") For Output As a
    b = 0
    Do While b < List3.ListCount
    Print #a, List3.List(b)
    b = b + 1
    Loop
    Close a
    'End If
C:
    Exit Sub
End Sub

Private Sub Command11_Click()
Form1.Width = 3480
End Sub

Private Sub Command12_Click()

If Text8 = "" Then
    number4.Caption = "Empty"
End If
If Text8 <> "" Then
    number4.Caption = Text8.Text + "SN" + " " + Text4.Text
End If
If Text11 = "" Then
    number5.Caption = "Empty"
End If
If Text11 <> "" Then
    number5.Caption = Text11 + "SN" + " " + Text14
End If
        R% = WritePrivateProfileString("Number4", "ascii", Text8.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendlist", Text4.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendstatus", Text5.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendthanx", Text6.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "find", Text7.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "send", Text9.Text, App.Path + "\KnK4Life.ini")
'<!-------------
        R% = WritePrivateProfileString("Number5", "ascii", Text11.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number5", "sendlist", Text15.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number5", "sendstatus", Text14.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number5", "sendthanx", Text13.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number5", "find", Text12.Text, App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number5", "send", Text10.Text, App.Path + "\KnK4Life.ini")
End Sub

Private Sub Command13_Click()
Select Case MsgBox("Are you sure you wanna reset method #4?", vbYesNo + vbQuestion + vbDefaultButton2, "Confermation")
Case vbYes
    Text8.Text = ""
    Text4.Text = "Empty"
    Text5.Text = "Empty"
    Text6.Text = "Empty"
    Text7.Text = "Empty"
    Text9.Text = "Empty"
Case vbNo
    Exit Sub
End Select

End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command15_Click()

End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command17_Click()
Select Case MsgBox("Are you sure you wanna reset method #5?", vbYesNo + vbQuestion + vbDefaultButton2, "Confermation")
Case vbYes
    Text11.Text = ""
    Text15.Text = "Empty"
    Text15.Text = "Empty"
    Text13.Text = "Empty"
    Text12.Text = "Empty"
    Text10.Text = "Empty"
Case vbNo
    Exit Sub
End Select
End Sub

Private Sub Command18_Click()
Timer2.Enabled = False
Timer3.Enabled = False
Call Endsetups
End Sub

Private Sub Command19_Click()
Chat1.ScanOff
End Sub

Private Sub Command2_Click()
List2.Clear
End Sub

Private Sub Command20_Click()
Chat1.ScanOn
End Sub

Private Sub Command21_Click()
Form1.Width = 3480
End Sub

Private Sub Command22_Click()
List5.Clear
End Sub

Private Sub Command23_Click()
If Text3.Text = "" Then
MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
Exit Sub
End If
Select Case True
Case number1.Checked
ChatSend ("/" + Text3 + " Send List")
Case number2.Checked
ChatSend ("!" + Text3 + " Send List")
Case number3.Checked
ChatSend ("-" + Text3 + " send list")
Case number4.Checked
If Text4.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text8 + Text3 + " " + Text4)
Case number5.Checked
If Text15.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text11 + Text3 + " " + Text15)
End Select
End Sub

Private Sub Command24_Click()
If Text3.Text = "" Then
MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
Exit Sub
End If
Select Case True
Case number1.Checked
ChatSend ("/" + Text3 + " Send Status")
Case number2.Checked
ChatSend ("!" + Text3 + " Send Status")
Case number3.Checked
ChatSend ("-" + Text3 + " send status")
Case number4.Checked
If Text5.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text8 + Text3 + " " + Text5)
Case number5.Checked
If Text14.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text11 + Text3 + " " + Text14)

End Select
End Sub

Private Sub Command25_Click()
If Text3.Text = "" Then
MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
Exit Sub
End If
Select Case True
Case number1.Checked
ChatSend ("/" + Text3 + " Send Thanks")
Case number2.Checked
ChatSend ("!" + Text3 + " Send Thanks")
Case number3.Checked
ChatSend ("-" + Text3 + " send thanks")
Case number4.Checked
If Text6.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text8 + Text3 + " " + Text6)
Case number5.Checked
If Text13.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text11 + Text3 + " " + Text13)
End Select
End Sub

Private Sub Command26_Click()
If Text3.Text = "" Then
MsgBox "error: No Server Screen Name Specified.", vbCritical, "error"
Exit Sub
End If
Select Case True
Case number1.Checked
ChatSend ("/" + Text3 + " Find " + Text1)
Case number2.Checked
ChatSend ("!" + Text3 + " Find " + Text1)
Case number3.Checked
ChatSend ("-" + Text3 + " find " + Text1)
Case number4.Checked
If Text7.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text8 + Text3 + " " + Text7 + " " + Text1)
Case number5.Checked
If Text12.Text = "Empty" Then
MsgBox "error: This Option hasn't been setup for this method.", vbExclamation, "error"
Exit Sub
End If
ChatSend (Text11 + Text3 + " " + Text12 + " " + Text1)
End Select

End Sub

Private Sub Command3_Click()
List4.Clear
Call AddRoomToListbox(List4, False)
End Sub

Private Sub Command4_Click()
Form1.Width = 3480
End Sub

Private Sub Command5_Click()
Command6.Enabled = True
Command5.Enabled = False
Frame5.Left = 5640
Frame5.Top = 4440
Frame4.Left = 0
Frame4.Top = 2160


End Sub

Private Sub Command6_Click()
Command5.Enabled = True

Command6.Enabled = False
Frame4.Left = 5640
Frame4.Top = 2280
Frame5.Left = 0
Frame5.Top = 2160


End Sub

Private Sub Command7_Click()
Form1.Height = 2430
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()
If List3.ListCount = 0 Then List3.AddItem Text2.Text: Exit Sub
    For i = 0 To List3.ListCount - 1
    num = LCase$(List3.List(i))
If num = LCase$(Text2.Text) Then Exit Sub
    Next i
    List3.AddItem Text2.Text
    Text2.Text = ""
End Sub

Private Sub dcl_Click()
dcl.Checked = True
cl.Checked = False
SaveSetting "KnK4Life", "KnK_Clearlist", "yesorno", "no"
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

End Sub

Private Sub exit_Click()
SetWindowLong Me.hWnd, GWL_WNDPROC, lProcOld
End
End Sub



Private Sub exit2_Click()
End
End Sub

Private Sub Form_Load()
FormOnTop Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong Me.hWnd, GWL_WNDPROC, lProcOld
End Sub

Private Sub help_Click()
On Local Error Resume Next
Call ShellExecute(hWnd, "Open", "/KnK4Life.HLP", "", App.Path, 1)
If Err Then
MsgBox "Error: File Not Found!", vbExclamation, "Error"
End If
End Sub

Private Sub help2_Click()
On Local Error Resume Next
Call ShellExecute(hWnd, "Open", "/KnK4Life.HLP", "", App.Path, 1)
If Err Then
MsgBox "Error: File Not Found!", vbExclamation, "Error"
End If
End Sub

Private Sub info_Click()
MsgBox "Making the first Server Helper was fun.  But I lost the goal I was aming for.  Instead of making a good user friendly program, I made one that took up tons of ram, and was glitchy.  I was trying to out do my good friend PooK.  Ware his Server Helper was awsome because of great programing, I made mine with more options, wich made it slower, and the programming was glitchy. It had 20 timers on 1 form.   This is a toned down version,  what I wanted in the first place.  Like always, This is dedicated to a good Friend, Sam, heh J/K You know its PooK, he's Awsome.  Thanx man", vbInformation, "Info"""
End Sub



Private Sub List1_DblClick()
If List2.ListCount = 0 Then List2.AddItem List1
For i = 0 To List2.ListCount - 1
num = List2.List(i)
If num = List1 Then Exit Sub
Next i
List2.AddItem List1
End Sub

Private Sub List2_DblClick()
List2.RemoveItem List2.ListIndex
End Sub

Private Sub List3_DblClick()
List3.RemoveItem List3.ListIndex
End Sub

Private Sub List3_KeyUp(KeyCode As Integer, Shift As Integer)
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

End Sub

Private Sub List4_DblClick()
Text3.Text = List4
End Sub

Private Sub List5_Click()
Text3.Text = List5
End Sub

Private Sub mailme_Click()

MsgBox "If you have any questons or comments, Please E-MAIL me at Bill@knk.tierranet.com" & vbCrLf, vbInformation, "Mail Me"

End Sub

Private Sub max_Click()
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

TaskBar.DeleteIcon (TASKBARICONID)

Call Unhook(Me)

Form1.Show
End Sub

Private Sub min_Click()

Form1.Hide

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
            MsgBox "This is not compatible with Windows NT 3.50 and earlier."
            Unload Me
            Exit Sub
        End If
    Case VER_PLATFORM_WIN32_WINDOWS
        'Running Windows 95 so no problem
    Case Else
        'Should only occur if Win32s is installed on Win 3.1, but the
        'program still won't work because there will be no task bar.
        MsgBox "This can only run under Windows 95 or Windows NT 3.51."
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
TaskBar.hWnd = Me

'Add icon to system tray
TaskBar.AddIcon Me.Icon, TASKBARICONID

'Assign a tooltip for the icon
'sTip = "Double-click exits Windows; Right-click displays menu"
sTip = "Right-click displays menu"

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



Private Sub namess_Click()
Frame6.Left = 4320
Frame6.Top = 4800
Frame2.Top = 4800
Frame2.Left = 3480
Frame1.Top = 0
Frame1.Left = 3480
Form1.Width = 4875
End Sub

Private Sub number1_Click()
number1.Checked = True
number2.Checked = False
number3.Checked = False
number4.Checked = False
number5.Checked = False

End Sub

Private Sub number2_Click()
number1.Checked = False
number2.Checked = True
number3.Checked = False
number4.Checked = False
number5.Checked = False

End Sub










'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o











Private Sub number3_Click()
number1.Checked = False
number2.Checked = False
number3.Checked = True
number4.Checked = False
number5.Checked = False

End Sub

Private Sub number4_Click()
If number4.Caption = "Empty" Then
MsgBox "error: Current method hasnt been setup.", vbExclamation, "error"

Exit Sub
End If
number1.Checked = False
number2.Checked = False
number3.Checked = False
number4.Checked = True
number5.Checked = False

End Sub

Private Sub number5_Click()
If number5.Caption = "Empty" Then
MsgBox "error: Current method hasnt been setup.", vbExclamation, "error"

Exit Sub
End If
number1.Checked = False
number2.Checked = False
number3.Checked = False
number4.Checked = False
number5.Checked = True


End Sub

Private Sub number6_Click()

End Sub

Private Sub off_Click()
 Call InstantMessage("$IM_OFF", "=)")
End Sub

Private Sub off2_Click()
off2.Checked = True
on2.Checked = False
SaveSetting "KnK4Life", "KnK_IM", "onoroff", "off"
End Sub

Private Sub on_Click()
 Call InstantMessage("$IM_ON", "=)")
End Sub

Private Sub on2_Click()
off2.Checked = False
on2.Checked = True
SaveSetting "KnK4Life", "KnK_IM", "onoroff", "on"

End Sub

Private Sub Option1_Click()
Text1.Enabled = False
End Sub

Private Sub Option2_Click()
Text1.Enabled = False
End Sub

Private Sub Option3_Click()
Text1.Enabled = False
End Sub

Private Sub Option4_Click()
Text1.Enabled = True
End Sub

Private Sub rb_Click()

End Sub

Private Sub sandm_Click()
Command5.Enabled = False
Frame4.Top = 2160
Frame4.Left = 0
Form1.Height = 5175
End Sub

Private Sub sm_Click()
Frame6.Left = 4320
Frame6.Top = 4800
Frame1.Top = 4800
Frame1.Left = 1920
Frame2.Top = 0
Frame2.Left = 3480

Form1.Width = 4875
'SSTab1.Tab = 0
End Sub

Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)

End Sub

Private Sub Timer1_Timer()
DoEvents
        Picture1.Picture = imOn.Picture
        SetCapture Picture1.hWnd
Pause (2)
        Picture1.Picture = imOff.Picture
        ReleaseCapture
Pause (2)
End Sub

Private Sub Timer2_Timer()
'<!---- Request Bin -->

Select Case True

Case number1.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend ("/" + Text3 + " Send " + List2.List(i))
    Pause (2.2)
    Next i
    Timer2.Enabled = False

Case number2.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend ("!" + Text3 + " Send " + List2.List(i))
    Pause (2.2)
    Next i
    Timer2.Enabled = False

Case number3.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend ("-" + Text3 + " send " + List2.List(i))
    Pause (2.2)
    Next i
    Timer2.Enabled = False

Case number4.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend (Text8 + Text3 + " " + Text4 + " " + List2.List(i))
    Pause (2.9)
    Next i
    Timer2.Enabled = False

Case number4.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend (Text11 + Text3 + " " + Text13 + " " + List2.List(i))
    Pause (2.9)
    Next i
    Timer2.Enabled = False


End Select
If cl.Checked = True Then
    List2.Clear
End If

If dcl.Checked = True Then
    End If
If off2.Checked = True Then
    Call InstantMessage("$IM_ON", "=)")
End If
Call Endsetups
End Sub

Private Sub Timer3_Timer()
'<!---- Find Bin -->
Select Case True
Case number1.Checked
    For i = 0 To List3.ListCount - 1
    ChatSend ("/" + Text3 + " Find " + List3.List(i))
    Pause (2.9)
    Next i
    Timer3.Enabled = False

Case number2.Checked
    For i = 0 To List3.ListCount - 1
    ChatSend ("!" + Text3 + " Find " + List3.List(i))
    Pause (2.9)
    Next i
    Timer3.Enabled = False

Case number3.Checked
    For i = 0 To List3.ListCount - 1
    ChatSend ("-" + Text3 + " Find " + List3.List(i))
    Pause (2.9)
    Next i
    Timer3.Enabled = False

Case number4.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend (Text8 + Text3 + " " + Text7 + " " + List3.List(i))
    Pause (2.9)
    Next i
    Timer3.Enabled = False

Case number4.Checked
    For i = 0 To List2.ListCount - 1
    ChatSend (Text11 + Text3 + " " + Text12 + " " + List3.List(i))
    Pause (2.9)
    Next i
    Timer3.Enabled = False

End Select
    
If off2.Checked = True Then
    Call InstantMessage("$IM_ON", "=)")
End If

Call Endsetups
End Sub
