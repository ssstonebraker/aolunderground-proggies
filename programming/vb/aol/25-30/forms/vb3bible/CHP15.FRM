VERSION 2.00
Begin Form formDrag 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon View"
   ClientHeight    =   2685
   ClientLeft      =   870
   ClientTop       =   1395
   ClientWidth     =   5040
   Height          =   3120
   Icon            =   CHP15.FRX:0000
   Left            =   795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5040
   Top             =   1035
   Width           =   5190
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   2880
      TabIndex        =   6
      Top             =   2160
      Width           =   1932
   End
   Begin CommandButton cmndMode 
      Caption         =   "&Manual"
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   2532
   End
   Begin PictureBox pictClear 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1560
      Width           =   492
   End
   Begin PictureBox pictIcon 
      Height          =   492
      Left            =   600
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   1560
      Width           =   492
   End
   Begin DriveListBox drivList 
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   1932
   End
   Begin DirListBox dirList 
      Height          =   1104
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   1932
   End
   Begin FileListBox fileList 
      Height          =   1200
      Left            =   240
      Pattern         =   "*.ico"
      TabIndex        =   0
      Tag             =   "fileList"
      Top             =   240
      Width           =   2535
   End
End
Const BEGIN_DRAG = 1     'drag
Const END_DRAG = 2

Const MANUAL = 0         'DragMode
Const AUTOMATIC = 1

Sub cmndExit_Click ()
    End                       'End the program
End Sub

Sub cmndMode_Click ()
    If cmndMode.Caption = "&Manual" Then     'Switch modes
        cmndMode.Caption = "&Automatic"      'relabel the command button
        pictIcon.DragMode = AUTOMATIC        'and switch the drag mode
    Else
        cmndMode.Caption = "&Manual"         'relabel the command button
        pictIcon.DragMode = MANUAL           'and switch the drag mode
    End If
End Sub

Sub dirList_Change ()
    fileList.Path = dirList.Path        'change the file list's directory
    ChDir fileList.Path                 'change to that directory
    fileList.SetFocus                   'and go to the file list box
End Sub

Sub drivList_Change ()
    dirList.Path = drivList.Drive            'change the dir list's path
End Sub

Sub fileList_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    fileList.Drag BEGIN_DRAG            'starting to drag a file from the file list box
End Sub

Sub pictClear_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
    Source.Drag END_DRAG                     'stop dragging now that we've dumped the file
    pictIcon.Picture = LoadPicture()         'blank out the picture (to simulate deletion)
End Sub

Sub pictIcon_DragDrop (Source As Control, X As Single, Y As Single)
    If Source.Tag = "fileList" Then          'check to make sure it's a new file being dragged
        pictIcon.Picture = LoadPicture(Source.List(Source.ListIndex)) 'show the icon file now
    End If
    pictIcon.DragIcon = pictIcon.Picture     'change to the picture (for the discard drag)
End Sub

