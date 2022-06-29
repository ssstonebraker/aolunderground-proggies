VERSION 2.00
Begin Form frmDropDown 
   BackColor       =   &H00E0E0E0&
   Caption         =   "The Beginners Guide To Visual Basic"
   ClientHeight    =   1560
   ClientLeft      =   1275
   ClientTop       =   2610
   ClientWidth     =   6675
   Height          =   2250
   Left            =   1215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   6675
   Top             =   1980
   Width           =   6795
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drop Down Menus"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   30
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   510
      TabIndex        =   0
      Top             =   300
      Width           =   5655
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFOpen 
         Caption         =   "&Open"
      End
      Begin Menu mnuFSave 
         Caption         =   "&Save"
      End
      Begin Menu mnuFSaveAs 
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
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin Menu mnuEPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin Menu mnuOptions 
      Caption         =   "&Options"
      Begin Menu mnuOSystem 
         Caption         =   "&System"
         Begin Menu mnuOSPrinter 
            Caption         =   "&Printer"
         End
         Begin Menu mnuOSDisplay 
            Caption         =   "&Display"
         End
         Begin Menu mnuOSMouse 
            Caption         =   "&Mouse"
         End
         Begin Menu mnuOSKeyboard 
            Caption         =   "&Keyboard"
         End
      End
      Begin Menu mnuOPrefs 
         Caption         =   "&Preferences"
         Begin Menu mnuOPColours 
            Caption         =   "&Colours"
         End
         Begin Menu mnuOPFont 
            Caption         =   "&Font"
         End
         Begin Menu mnuOPSound 
            Caption         =   "&Sound"
         End
      End
   End
End
Option Explicit

Sub mnuFExit_Click ()

    Unload Me

End Sub

