VERSION 5.00
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SYSTRAY.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "System Tray Example  By KnK"
   ClientHeight    =   3210
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   3210
   Icon            =   "systrayex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "systrayex.frx":030A
   ScaleHeight     =   3210
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Minimize"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   3255
   End
   Begin SysTray.SystemTray SystemTray1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu info 
         Caption         =   "I&nfo"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu minimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Constants for the ACTION PROPERTY
Private Const sys_Add = 0       'Specifies that an icon is being add
Private Const sys_Modify = 1    'Specifies that an icon is being modified
Private Const sys_Delete = 2    'Specifies that an icon is being deleted

'Constants for the ERROR EVENT
Private Const errUnableToAddIcon = 1    'Icon can not be added to system tray
Private Const errUnableToModifyIcon = 2 'System tray icon can not be modified
Private Const errUnableToDeleteIcon = 3 'System tray icon can not be deleted
Private Const errUnableToLoadIcon = 4   'Icon could not be loaded (occurs while using icon property)

'Constants for MOUSE RELATETED EVENTS
Private Const vbLeftButton = 1     'Left button is pressed
Private Const vbRightButton = 2    'Right button is pressed
Private Const vbMiddleButton = 4   'Middle button is pressed

Private Const cdlOpen = 1
Private Const cdlSave = 2
Private Const cdlColor = 3
Private Const cdlPrint = 4
Private Const cdlOFNReadOnly = 1             'Checks Read-Only check box for Open and Save As dialog boxes.
Private Const cdlOFNOverwritePrompt = 2      'Causes the Save As dialog box to generate a message box if the selected file already exists.
Private Const cdlOFNHideReadOnly = 4         'Hides the Read-Only check box.
Private Const cdlOFNNoChangeDir = 8          'Sets the current directory to what it was when the dialog box was invoked.
Private Const cdlOFNHelpButton = 10          'Causes the dialog box to display the Help button.
Private Const cdlOFNNoValidate = 100         'Allows invalid characters in the returned filename.
Private Const cdlOFNAllowMultiselect = 200   'Allows the File Name list box to have multiple selections.
Private Const cdlOFNExtensionDifferent = 400 'The extension of the returned filename is different from the extension set by the DefaultExt property.
Private Const cdlOFNPathMustExist = 800      'User can enter only valid path names.
Private Const cdlOFNFileMustExist = 1000     'User can enter only names of existing files.
Private Const cdlOFNCreatePrompt = 2000      'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
Private Const cdlOFNShareAware = 4000        'Sharing violation errors will be ignored.
Private Const cdlOFNNoReadOnlyReturn = 8000  'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
Private Const cdlOFNExplorer = 8000          'Use the Explorer-like Open A File dialog box template.  (Windows 95 only.)
Private Const cdlOFNNoDereferenceLinks = 100000
Private Const cdlOFNLongNames = 200000

Private Sub Command1_Click()
SystemTray1.Action = sys_Add
Form1.Hide
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
SystemTray1.Icon = Val(Form1.Icon)
' System Tray Example By KnK
'Goto http://knk.tierranet.com/knk to get
'all your vb needs.
'Any questons E-MAIL me at Bill@knk.tierranet.com
End Sub

Private Sub info_Click()
MsgBox "This is just a simple example of how to use Systemtray.ocx, Since there example is a bit difficult.  Feel free to use this for your prog or whatever.  " & vbCrLf & "Site: http://knk.tierranet.com/knk4o" & vbCrLf & "E-MAIL: Bill@knk.tierranet.com" & vbCrLf & vbCrLf & "Alot more stuff at my site, hope to se ya.", vbInformation, "info"
End Sub

Private Sub minimize_Click()
SystemTray1.Action = sys_Add
Form1.Hide
End Sub

Private Sub SystemTray1_Error(ByVal ErrorNumber As Integer)
Beep
End Sub

Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
SystemTray1.Action = sys_Delete
Form1.Show
End Sub
