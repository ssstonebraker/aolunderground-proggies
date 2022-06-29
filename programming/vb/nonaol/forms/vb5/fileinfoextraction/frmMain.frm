VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Information Extraction Program"
   ClientHeight    =   2700
   ClientLeft      =   3765
   ClientTop       =   3105
   ClientWidth     =   4575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4575
   Begin VB.CheckBox File_System 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3360
      TabIndex        =   13
      Top             =   1920
      Width           =   200
   End
   Begin VB.CheckBox File_Hidden 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   200
   End
   Begin VB.CheckBox File_ReadOnly 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   1920
      Width           =   200
   End
   Begin VB.CheckBox File_Archive 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   200
   End
   Begin VB.CheckBox File_Normal 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   200
   End
   Begin VB.Frame FileInfo_Frame 
      Caption         =   "File Information"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      Begin VB.TextBox FileName 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton picSmall_save 
         Caption         =   "Save..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton picLarge_Save 
         Caption         =   "Save..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox picLarge 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1200
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   480
      End
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1200
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   960
         Width           =   480
      End
      Begin VB.CommandButton cmd_Open 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   230
         Width           =   975
      End
      Begin VB.Label lblNormal 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblSystem 
         BackStyle       =   0  'Transparent
         Caption         =   "System"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHidden 
         BackStyle       =   0  'Transparent
         Caption         =   "Hidden"
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblReadOnly 
         BackStyle       =   0  'Transparent
         Caption         =   "Read Only"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblArchive 
         BackStyle       =   0  'Transparent
         Caption         =   "Archive"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label FileType 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblFileType 
         BackStyle       =   0  'Transparent
         Caption         =   "File Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLargeIcon 
         BackStyle       =   0  'Transparent
         Caption         =   "Large Icon:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblSmallIcon 
         BackStyle       =   0  'Transparent
         Caption         =   "Small Icon:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label curFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Current File:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Label lblAbout2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: PSXKid3@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Info Extraction Example by: Crash ZER0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LoadInfo(File As String)
'This sets the attributes to the file using check boxes
  'fill in check box if its an Archive
File_Archive.Value = ReturnValue(CBool(GetAttr(File$) And vbArchive))
  'fill in check box if its Hidden
File_Hidden.Value = ReturnValue(CBool(GetAttr(File$) And vbHidden))
  'fill in check box if its Read Only
File_ReadOnly.Value = ReturnValue(CBool(GetAttr(File$) And vbReadOnly))
  'fill in check box if its System
File_System.Value = ReturnValue(CBool(GetAttr(File$) And vbSystem))
  'fill in check box if its Normal
File_Normal.Value = ReturnValue(CBool(GetAttr(File$) And vbNormal))
'--------------------------------------------------
'The perimets for ReturnValue() are:
'Integer% = ReturnValue(BooleanAttribute
'[returnd by CBool(), which converts a number to boolean])
'******************************************
'for a description (not that you'll need)
'see the function
'--------------------------------------------------
'--------------------------------------------------
'The perimeters for GetAttr() are:
'VarVBFileAttribute = GetAttr(FilePath$)
'***************************************
'GetAttr() returns a file's attributes
'--------------------------------------------------
End Sub
Public Function ReturnValue(Attr As Boolean) As Integer
'This function gets a boolean value and returns
'a value for a checkbox.
'(I would have used CInt(), except when it passes
'True as an argument it returns -1 and therefore, creates an error.
 Select Case (Attr) 'select what we're looking at
  Case True:  'if its value is True then
   ReturnValue = 1 'return 1
  Case Else:  'however, if its value is false
   ReturnValue = 0 'return 0
 End Select
End Function

Private Sub cmd_Open_Click()
 'set and display dialog to open file
With dlgDialog
 .DialogTitle = "Open File" 'set title
 .Filter = "All Files (*.*)|*.*" 'set type to open
 .ShowOpen  'show dialog
End With

 'Put the FilePath into the text box
FileName = dlgDialog.FileName

 'Load the file icons
Dim LargeIcon As Long, SmallIcon As Long
Dim shFile As SHFILEINFO
 'this assigns a negative long which will be used to draw the icon
   'get negative long for large icon
 LargeIcon& = SHGetFileInfo(dlgDialog.FileName, 0&, shFile, Len(shFile), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
   'get negative long for small icon
 SmallIcon& = SHGetFileInfo(dlgDialog.FileName, 0&, shFile, Len(shFile), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'---------------------------------------------
'The perimeters are for SHGetFileInfo() are:
'Long& = SHGetFileInfo(FilePath$, _
'FileAttributes& [0&], _
'varSHFILEINFO [a variable of SHFILEINFO], _
'Len(varSHFILEINFO) [the length of the type], _
'FlagConst [SmallIcon or LargeIcon])
'----------------------------------------------
 'Now, put the file type into the label
   'this gets everything before the null character in the type name and displays it
   'Look up in the help file for an explanation of Left$(), InStr(), and Chr$()
FileType = Left$(shFile.szTypeName, InStr(shFile.szTypeName, Chr$(0)) - 1)

 'Here is where we ready the picture
 'boxes by clearing them
    'clears the small picture box
picLarge.Picture = LoadPicture()
   'clears the large picture box
picSmall.Picture = LoadPicture()

 'Now this draws the icons into the picture boxes
    'draw the large icon
r& = ImageList_Draw(LargeIcon&, shFile.iIcon, picLarge.hDC, 0&, 0&, ILD_TRANSPARENT)
    'draw the small icon
r& = ImageList_Draw(SmallIcon&, shFile.iIcon, picSmall.hDC, 0&, 0&, ILD_TRANSPARENT)
'------------------------------------------------------------------------
'The perimeters for ImageList_Draw() are:
'Long& = ImageList_Draw(NegativeLong& [contains the long of the icon size], _
'varSHFILEINFO.iIcon [iIcon, which is filled by SHGetFileInfo], _
'PictureBox.hDC [hDC of picturebox to be drawn in], _
'X& [where the icon should be drawn in the picture box, width-wise], _
'Y& [where the icon should be draw in the picture box, height-wise], _
'Flags [transparent, use ILD_TRANSPARENT])
'-------------------------------------------------------------------------
  'Fill in the file information
  '(See description in subroutine LoadInfo()
LoadInfo dlgDialog.FileName 'pass FilePath as argument

'Enable "Save" command buttons
picSmall_save.Enabled = True
picLarge_Save.Enabled = True
End Sub

Private Sub FileName_DblClick()
'MsgBox the current file (see visual basic help file for information on this function)
MsgBox FileName, vbOKOnly, "Current File:"
End Sub

Private Sub picLarge_Save_Click()
'Saves the icon to a file
'~~~~~~~~~~~~~~~~~~~~~~~~~

  'Display dialog to find out where to save it
With dlgDialog
 .DialogTitle = "Save As" 'set dialog title
 .Filter = "Icon File (*.ico)|*.ico"  'set type to save
 'Set the default filename (gets whatever is left of the file extension)
 .FileName = Left$(.FileName, Len(.FileName) - 4)
 .ShowSave 'display dialog
End With

 'save the icon
SavePicture picLarge.Image, dlgDialog.FileName
'------------------------------------------------------------------------------------
'The perimeters for SavePicture are:
'SavePicture PictureBox.Image [contains the info needed to save the picture data], _
'FilePath$ [where to save the icon, including filename and extension])
'------------------------------------------------------------------------------------
End Sub

Private Sub picSmall_save_Click()
'Saves the icon to a file
'~~~~~~~~~~~~~~~~~~~~~~~~~

  'Display dialog to find out where to save it
With dlgDialog
 .DialogTitle = "Save As" 'set dialog title
 .Filter = "Icon File (*.ico)|*.ico"  'set type to save
 'Set the default filename (gets whatever is left of the file extension)
 .FileName = Left$(.FileName, Len(.FileName) - 4)
 .ShowSave 'display dialog
End With

 'save the icon
SavePicture picSmall.Image, dlgDialog.FileName
'------------------------------------------------------------------------------------
'The perimeters for SavePicture are:
'SavePicture PictureBox.Image [contains the info needed to save the picture data], _
'FilePath$ [where to save the icon, including filename and extension])
'------------------------------------------------------------------------------------
End Sub

