VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type APPBARDATA_TYPE
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long '  message specific
End Type

Private Type NOTIFYICONDATA_TYPE
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
lpData As NOTIFYICONDATA_TYPE) As Long

Private Declare Function SHAppBarMessage Lib "shell32.dll" _
(ByVal dwMessage As Long, pData As APPBARDATA_TYPE) As Long

'Constants used with Shell_NotifyIcon
Private Const NIM_ADD = 0&
Private Const NIM_DELETE = 2&
Private Const NIM_MODIFY = 1&
Private Const NIF_ICON = 2&
Private Const NIF_MESSAGE = 1&
Private Const NIF_TIP = 4&

'Constants used with SHAppBarMessage
Const ABM_GETTASKBARPOS = &H5
Const ABM_WINDOWPOSCHANGED = &H9

Private NotifyIconData As NOTIFYICONDATA_TYPE
Private AppBarData As APPBARDATA_TYPE

Private Const WM_USER = &H400

'This is the message we'll assign to the uCallbackMessage
'member of the NotifyIconData structure.  This message is sent
'by Windows when a mouse event occurs within the bounding
'rectangle of the icon added to the taskbar status area.
Private Const MYWM_NOTIFYICON = WM_USER + &H301


Public Sub ChangeIcon(hIcon As Long, ID As Long)

'Changes the icon identified by ID

Dim lResult As Long

NotifyIconData.uID = ID
NotifyIconData.hIcon = hIcon
NotifyIconData.uFlags = NIF_ICON
lResult = Shell_NotifyIcon(NIM_MODIFY, NotifyIconData)

End Sub

Public Property Let hwnd(Frm As Form)
    
'Assigns the app's hwnd to the notification structure
NotifyIconData.hwnd = Frm.hwnd

End Property


Public Property Get CallbackMessage%()

'Returns the message number
CallbackMessage = NotifyIconData.uCallbackMessage

End Property


Public Sub AddIcon(hIcon As Long, ID As Long)

'Adds a new icon to the system tray
'hIcon is a handle to the icon we want to add
'ID is the unique identifier for the icon

Dim lResult As Long

NotifyIconData.uID = ID
NotifyIconData.hIcon = hIcon
NotifyIconData.uFlags = NIF_ICON Or NIF_MESSAGE
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o
NotifyIconData.uCallbackMessage = MYWM_NOTIFYICON
lResult = Shell_NotifyIcon(NIM_ADD, NotifyIconData)

End Sub

Public Function Exists() As Boolean

'Verify that the task bar exists.  If the user is using a
'3rd party shell, it might not.  Additionally returns the
'dimensions of the task bar.  This is useful for positioning
'a form relative to the taskbar.

AppBarData.cbSize = 36&
Exists = SHAppBarMessage(ABM_GETTASKBARPOS, AppBarData)

End Function


Public Sub ChangeTip(sNewTip As String, ID As Long)
    
'Changes the text displayed when the user pauses
'the mouse over the icon.

Dim lResult As Long

NotifyIconData.szTip = sNewTip & vbNullChar
NotifyIconData.uID = ID
NotifyIconData.uFlags = NIF_TIP
lResult = Shell_NotifyIcon(NIM_MODIFY, NotifyIconData)

End Sub

Public Sub DeleteIcon(ID As Long)
    
'Deletes the specified icon from the system tray
Dim lResult As Long

NotifyIconData.uID = ID
lResult = Shell_NotifyIcon(NIM_DELETE, NotifyIconData)

End Sub

Public Sub DeleteAllIcons(nIcons As Integer)

'Removes ALL icons added to the taskbar by our app
'For this to work, two things are assumed:
'1.  You are storing (by whatever means, a global variable, etc)
'the number of icons that have been added.  This value should
'be passed to the procedure.
'2.  The IDs for those icons are sequential (1, 2, 3, etc.)


Dim iCounter As Integer
Dim lResult As Long

For iCounter = 1 To nIcons
    NotifyIconData.uID = iCounter
    lResult = Shell_NotifyIcon(NIM_DELETE, NotifyIconData)
Next iCounter

End Sub

Private Sub Class_Initialize()

'Initialize elements of the data structure
NotifyIconData.cbSize = 88&
NotifyIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP

End Sub

