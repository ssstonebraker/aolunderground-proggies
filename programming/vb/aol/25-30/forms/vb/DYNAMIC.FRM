VERSION 2.00
Begin Form frmDynamic 
   BackColor       =   &H00E0E0E0&
   Caption         =   "The Beginners Guide To Visual Basic"
   ClientHeight    =   2925
   ClientLeft      =   1725
   ClientTop       =   2730
   ClientWidth     =   4680
   Height          =   3615
   Left            =   1665
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   Top             =   2100
   Width           =   4800
   Begin CommonDialog dlgOpen 
      Left            =   4065
      Top             =   2430
   End
   Begin Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dynamic Menus"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   22.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   390
      TabIndex        =   0
      Top             =   960
      Width           =   3945
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFOpen 
         Caption         =   "&Open"
      End
      Begin Menu mnuFSave 
         Caption         =   "&Save"
      End
      Begin Menu mnuFSaveAS 
         Caption         =   "Save &As"
      End
      Begin Menu mnuFDelete 
         Caption         =   "&Delete"
      End
      Begin Menu mnuFdash1 
         Caption         =   "-"
      End
      Begin Menu mnuFExit 
         Caption         =   "E&xit"
      End
      Begin Menu mnuFileList 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Option Explicit


Dim nIndex As Integer

Sub mnuFExit_Click ()

    Unload Me

End Sub

Sub mnuFile_Click ()

    nIndex = 1

End Sub

Sub mnuFOpen_Click ()

    On Error GoTo OpenError

    dlgOpen.CancelError = True

    dlgOpen.DialogTitle = "Select the file to open"
    dlgOpen.Filter = "All Files (*.*)|*.*)"
    dlgOpen.FilterIndex = 1
    dlgOpen.Action = 1

    '--------------------------------------------------------------------------------------
    ' Insert menu code here
    
    
    Exit Sub

OpenError:
    On Error GoTo 0
    Exit Sub
End Sub

