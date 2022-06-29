VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SysTray.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Tray Control Demo"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SysTray.SystemTray SystemTray1 
      Left            =   3120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   ""
      IconFile        =   0
   End
   Begin VB.CommandButton cmd_IconLoaded 
      Caption         =   "Is Icon Loaded?"
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   3240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cdl_Main 
      Left            =   2520
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame fra_Appearance 
      Caption         =   "Appearance:"
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   5535
      Begin VB.CommandButton cmd_browse 
         Caption         =   "Browse..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton opt_OtherIcon 
         Caption         =   "Use Other Icon"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton opt_FormIcon 
         Caption         =   "Use Form Icon"
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txt_TipText 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "System Tray Control Test"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lbl_file 
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lbl_SysTrayTipText 
         Caption         =   "System Tray Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame fra_Messages 
      Caption         =   "Messages To Be Shown:"
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      Begin VB.OptionButton opt_MouseDblClk 
         Caption         =   "Show message when mouse is double clicked."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   4695
      End
      Begin VB.OptionButton opt_MouseMove 
         Caption         =   "Show message when mouse moves over system tray icon."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   4695
      End
      Begin VB.OptionButton opt_MouseUp 
         Caption         =   "Show message when mouse is released on system tray icon."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4815
      End
      Begin VB.OptionButton opt_MouseDown 
         Caption         =   "Show message when mouse is down on system tray icon."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmd_DeleteIcon 
      Caption         =   "Delete Icon"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmd_ModifyIcon 
      Caption         =   "Modify Icon"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmd_ShowIcon 
      Caption         =   "Show Icon"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lbl_notes 
      Caption         =   "IMPORTANT: After making changes click on modify icon if icon already exists."
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3240
      Width           =   3015
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------
'PROGRAM: ActiveX System Tray Control Demo
'AUTHOR: Nathan Blecharczyk
'
'This program's purpose is to demonstate the
'ActiveX System Tray Control. This program is
'freeware and may be modified by anyone as long
'as the original source and author are noted.
'This program demonstates the uses of all the
'properties and events of the ActiveX System
'Tray Control. All constants used by this
'control are listed below. You may contact me at
'b_nathan@juno.com if you have any questions or
'comments.
'-----------------------------------------------
'PRODUCT NAME: ActiveX System Tray Control
'VERSION: 1.0
'PROGRAMMER: Nathan Blecharczyk
'
'INFORMATION:
'Action - Performs an operation according to the passed value
'Error - Event that occurs after an error has taken place
'Icon - Icon to display in system tray. Can be a numeric value or path and file name
'IconAdded - Event that occurs after an icon has been added to system tray
'IconDeleted - Event that occurs after an icon has been modified on system tray
'IconModified - Event that occurs after an icon has been deleted from system tray
'IsIconLoaded - Function that returns true if an icon is being displayed in the system tray. Otherwise false.
'MouseDblClk - Event that occurs after mouse button has been double clicked while over the icon located in the system tray.
'MouseDown - Event that occurs after mouse button has been pressed down while over the icon located in the system tray.
'MouseMove - Event that occurs after mouse has been moved over the icon located in the system tray.
'MouseUp - Event that occurs after mouse button has been released while over the icon located in the system tray.
'SysTrayText - Text to be displayed after mouse has been idle over the icon in the system tray for a while. Displays up to 64 characters.
'
'EVENTS:        PROPERTIES:     FUNCTIONS:      Subroutines:
'Error          Action          IsIconLoaded
'IconAdded      Icon
'IconDeleted    SysTrayText
'IconModified
'MouseDblClk
'MouseDown
'MouseMove
'MouseUp
'-----------------------------------------------

Option Explicit

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

'-------------OTHER CONTROL CONSTANTS------------
'Common Dialogue Control
Private Const cdlOpen = 1
Private Const cdlSave = 2
Private Const cdlColor = 3
Private Const cdlPrint = 4
'File stuff
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
Private Const cdlOFNNoDereferenceLinks = 100000 'Do not dereference shortcuts (shell links).  By default, choosing a shortcut causes it to be dereferenced by the shell.  (Windows 95 only.)
Private Const cdlOFNLongNames = 200000       'Use Long filenames.  (Windows 95 only.)

Private Sub cmd_ModifyIcon1_Click()
SystemTray1.Action = sys_Add
End Sub

Private Sub cmd_browse_Click()
On Error GoTo BrowseError

'setup open dialogue box
cdl_Main.CancelError = True
cdl_Main.Filter = "Icon Files (*.ico)|*.ico|"
cdl_Main.InitDir = App.Path
cdl_Main.Flags = cdlOFNHideReadOnly
cdl_Main.DialogTitle = "Open Icon"
cdl_Main.Action = cdlOpen

lbl_file.Caption = cdl_Main.filename
SystemTray1.Icon = cdl_Main.filename
Exit Sub

BrowseError:
Exit Sub
End Sub

Private Sub cmd_DeleteIcon_Click()
SystemTray1.Action = sys_Delete
End Sub

Private Sub cmd_IconLoaded_Click()

If SystemTray1.IsIconLoaded = True Then
   MsgBox "The icon is currently being displayed. (IsIconLoaded = True)"
Else
    MsgBox "The icon is currently not being displayed. (IsIconLoaded = false)"
End If

End Sub

Private Sub cmd_ModifyIcon_Click()
SystemTray1.Action = sys_Modify
End Sub

Private Sub cmd_ShowIcon_Click()
SystemTray1.Action = sys_Add
End Sub

Private Sub Form_Load()
SystemTray1.Icon = Val(frm_main.Icon)
SystemTray1.SysTrayText = txt_TipText
lbl_file.Caption = "Form Icon"
End Sub

Private Sub opt_FormIcon_Click()
SystemTray1.Icon = Val(frm_main.Icon)
cmd_browse.Enabled = False
lbl_file.Caption = "Form Icon"

End Sub

Private Sub opt_OtherIcon_Click()
cmd_browse.Enabled = True
If lbl_file.Caption = "" Then lbl_file.Caption = "Form Icon"

End Sub

Private Sub SystemTray1_Error(ByVal ErrorNumber As Integer)
Beep
MsgBox "Error numer " + Format(ErrorNumber) + " has occured!"
End Sub

Private Sub SystemTray1_IconAdded()
MsgBox "This program's icon has been added to the system tray!"
End Sub

Private Sub SystemTray1_IconDeleted()
MsgBox "This program's icon has been deleted from the system tray!"
End Sub

Private Sub SystemTray1_IconModified()
MsgBox "This program's icon in the system tray has been modified!"
End Sub

Private Sub SystemTray1_MouseDblClk(ByVal Button As Integer)
If opt_MouseDblClk.Value = True Then
    MsgBox "Button number " + Format(Button) + " has been double clicked."
End If
End Sub

Private Sub SystemTray1_MouseDown(ByVal Button As Integer)
If opt_MouseDown.Value = True Then
    MsgBox "Button number " + Format(Button) + " has been pushed down."
End If
End Sub

Private Sub SystemTray1_MouseMove()
If opt_MouseMove.Value = True Then
    MsgBox "The mouse has been moved over the icon in the system tray."
End If
End Sub

Private Sub SystemTray1_MouseUp(ByVal Button As Integer)
If opt_MouseUp.Value = True Then
    MsgBox "Button number " + Format(Button) + " has been released."
End If
End Sub

Private Sub SystemTray2_Error()

End Sub

Private Sub txt_TipText_Change()

SystemTray1.SysTrayText = txt_TipText
End Sub
