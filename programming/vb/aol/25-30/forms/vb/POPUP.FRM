VERSION 2.00
Begin Form frmPopups 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Pop up menus"
   ClientHeight    =   2535
   ClientLeft      =   1425
   ClientTop       =   2790
   ClientWidth     =   5925
   Height          =   3225
   Left            =   1365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5925
   Top             =   2160
   Width           =   6045
   Begin CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   4455
      TabIndex        =   1
      Top             =   420
      Width           =   1215
   End
   Begin CommandButton cmdFile 
      Caption         =   "&File"
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   405
      Width           =   1215
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click a button to pop up menu."
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   210
      TabIndex        =   2
      Top             =   1110
      Width           =   5475
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFOpen 
         Caption         =   "&Open"
      End
      Begin Menu mnuFSave 
         Caption         =   "&Save"
      End
      Begin Menu mnuSaveAS 
         Caption         =   "Save &As"
      End
      Begin Menu mnuFDash1 
         Caption         =   "-"
      End
      Begin Menu mnuFExit 
         Caption         =   "E&xit"
      End
   End
   Begin Menu mnuEdit 
      Caption         =   "&Edit"
      Begin Menu mnuEUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin Menu mnuEDash1 
         Caption         =   "-"
      End
      Begin Menu mnuECut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin Menu mnuECopy 
         Caption         =   "C&opy "
         Shortcut        =   ^C
      End
      Begin Menu mnuEPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Option Explicit

'Popup Menu Constants
Const POPUPMENU_LEFTALIGN = 0
Const POPUPMENU_CENTERALIGN = 4
Const POPUPMENU_RIGHTALIGN = 8

Sub cmdEdit_Click ()
    
    ' Declare two variables to hold the top and left coordinates for the
    ' pop up menu.
    Dim nMenuTop As Integer
    Dim nMenuLeft As Integer

    ' The menu's left hand edge is lined up on the same line as the command button, so the
    ' left property of the command button is used to set the nMenuLeft variable
    nMenuLeft = cmdEdit.Left

    ' We want the menu's top edge to appear directly beneath the command button. A little
    ' arithmetic (adding the command button's Top to its Height) gives us the menu's top.
    nMenuTop = cmdEdit.Top + cmdFile.Height

    frmPopups.PopupMenu mnuEdit, POPUPMENU_LEFTALIGN, nMenuLeft, nMenuTop

End Sub

Sub cmdFile_Click ()

    ' Declare two variables to hold the top and left coordinates for the
    ' pop up menu.
    Dim nMenuTop As Integer
    Dim nMenuLeft As Integer

    ' The menu's left hand edge is lined up on the same line as the command button, so the
    ' Left property of the command button is used to set the nMenuLeft variable.
    nMenuLeft = cmdFile.Left

    ' We want the menu's top edge to appear directly beneath the command button. A little
    ' arithmetic (adding the command buttons Top to its height) gives us the menus top.
    nMenuTop = cmdFile.Top + cmdFile.Height

    ' Finally, the required menu, in this case the file menu, is "popped up".
    frmPopups.PopupMenu mnuFile, POPUPMENU_LEFTALIGN, nMenuLeft, nMenuTop

End Sub

