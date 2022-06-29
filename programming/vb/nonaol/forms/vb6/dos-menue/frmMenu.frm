VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dos's menu example - 2.12.99"
   ClientHeight    =   1425
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "file"
      Begin VB.Menu mnuFileDummy 
         Caption         =   "dummy"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "edit"
      Begin VB.Menu mnuEditDummy 
         Caption         =   "dummy"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "search"
      Begin VB.Menu mnuSearchDummy 
         Caption         =   "dummy"
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "colors"
      Begin VB.Menu mnuColorDummy 
         Caption         =   "dummy"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************
'* 2.12.99                                                *
'* this example shows how to modify menus and insert menu *
'* items through the use of subclassing with no added     *
'* controls. by modifying existing menus and creating     *
'* additional owner draw menus, you are given complete    *
'* control over several menu attributes which are not     *
'* normally available in visual basic. with this extra    *
'* control comes added responsibility as you'll see. keep *
'* in mind that since we are subclassing, all windows must*
'* be unhooked. so do not use the end button on your      *
'* visual basic window. instead, close the form with the  *
'* control box. failing to do so will crash your visual   *
'* basic.                                                 *
'*                                                        *
'* dos                                                    *
'*                                                        *
'* email:     xdosx@hotmail.com                           *
'* aim:       xdosx                                       *
'* web site:  http://www.hider.com/dos                    *
'**********************************************************

Private Sub Form_Load()
    Dim lngMainMenu As Long
    
    'get our menu's handle
    lngMainMenu& = GetMenu(Me.hwnd)
    
    'get the submenu handles
    lngFile& = GetSubMenu(lngMainMenu&, 0&)
    lngEdit& = GetSubMenu(lngMainMenu&, 1&)
    lngSearch& = GetSubMenu(lngMainMenu&, 2&)
    lngColors& = GetSubMenu(lngMainMenu&, 3&)
    
    'modify our main submenus
    Call ModifyMenu(lngMainMenu&, 0&, MF_BYPOSITION Or MF_OWNERDRAW, lngFile&, "")
    Call ModifyMenu(lngMainMenu&, 1&, MF_BYPOSITION Or MF_OWNERDRAW, lngEdit&, "")
    Call ModifyMenu(lngMainMenu&, 2&, MF_BYPOSITION Or MF_OWNERDRAW, lngSearch&, "")
    Call ModifyMenu(lngMainMenu&, 3&, MF_BYPOSITION Or MF_OWNERDRAW, lngColors&, "")
    
    'set file menu's submenu ids
    lngNew& = 71
    lngOpen& = 72&
    lngSave& = 73&
    lngSaveAs& = 74&
    lngExit& = 75&
    
    'insert the file menu's sub menus
    Call InsertMenuByNum(lngFile&, 0&, MF_BYPOSITION Or MF_OWNERDRAW, lngNew&, 0)
    Call InsertMenuByNum(lngFile&, 1&, MF_BYPOSITION Or MF_OWNERDRAW, lngOpen&, 0)
    Call InsertMenuByNum(lngFile&, 2&, MF_BYPOSITION Or MF_OWNERDRAW, lngSave&, 0)
    Call InsertMenuByNum(lngFile&, 3&, MF_BYPOSITION Or MF_OWNERDRAW, lngSaveAs&, 0)
    Call InsertMenuByNum(lngFile&, 4&, MF_BYPOSITION Or MF_OWNERDRAW, lngExit&, 0)
    
    'set edit menu's submenu ids
    lngUndo& = 81&
    lngCut& = 82&
    lngCopy& = 83&
    lngPaste& = 84&
    lngDelete& = 85&
    
    'insert the edit menu's sub menus
    Call InsertMenuByNum(lngEdit&, 0&, MF_BYPOSITION Or MF_OWNERDRAW, lngUndo&, 0)
    Call InsertMenuByNum(lngEdit&, 1&, MF_BYPOSITION Or MF_OWNERDRAW, lngCut&, 0)
    Call InsertMenuByNum(lngEdit&, 2&, MF_BYPOSITION Or MF_OWNERDRAW, lngCopy&, 0)
    Call InsertMenuByNum(lngEdit&, 3&, MF_BYPOSITION Or MF_OWNERDRAW, lngPaste&, 0)
    Call InsertMenuByNum(lngEdit&, 4&, MF_BYPOSITION Or MF_OWNERDRAW, lngDelete&, 0)
    
    'set search menu's submenu ids
    lngFind& = 91&
    lngFindNext& = 92&
    
    'insert the edit menu's sub menus
    Call InsertMenuByNum(lngSearch&, 0&, MF_BYPOSITION Or MF_OWNERDRAW, lngFind&, 0)
    Call InsertMenuByNum(lngSearch&, 1&, MF_BYPOSITION Or MF_OWNERDRAW, lngFindNext&, 0)
    
    'set colors menu's submenu ids
    lngBlack& = 101
    lngGreen& = 102
    lngPurple& = 103
    lngRed& = 104
    lngYellow& = 105
    lngBlue& = 106
    lngWhite& = 107
    
    'insert the colors menu's sub menus
    Call InsertMenuByNum(lngColors&, 0&, MF_BYPOSITION Or MF_OWNERDRAW, lngBlack&, 0)
    Call InsertMenuByNum(lngColors&, 1&, MF_BYPOSITION Or MF_OWNERDRAW, lngGreen&, 0)
    Call InsertMenuByNum(lngColors&, 2&, MF_BYPOSITION Or MF_OWNERDRAW, lngPurple&, 0)
    Call InsertMenuByNum(lngColors&, 3&, MF_BYPOSITION Or MF_OWNERDRAW, lngRed&, 0)
    Call InsertMenuByNum(lngColors&, 4&, MF_BYPOSITION Or MF_OWNERDRAW, lngYellow&, 0)
    Call InsertMenuByNum(lngColors&, 5&, MF_BYPOSITION Or MF_OWNERDRAW, lngBlue&, 0)
    Call InsertMenuByNum(lngColors&, 6&, MF_BYPOSITION Or MF_OWNERDRAW, lngWhite&, 0)
    
    'delete our dummy menus
    Call DeleteMenu(lngFile&, 5&, MF_BYPOSITION)
    Call DeleteMenu(lngEdit&, 5&, MF_BYPOSITION)
    Call DeleteMenu(lngSearch&, 2&, MF_BYPOSITION)
    Call DeleteMenu(lngColors&, 7&, MF_BYPOSITION)
    
    'set our hook
    gHW = Me.hwnd
    Call Hook
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'important! you must unhook your subclassing before you
    'close your program. hitting the end button in vb will
    'crash visual basic.
    Call Unhook
End Sub
