VERSION 5.00
Begin VB.Form frmWinMine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinMine 99"
   ClientHeight    =   6165
   ClientLeft      =   1890
   ClientTop       =   2670
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "winmine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "winmine.frx":030A
   ScaleHeight     =   6165
   ScaleWidth      =   8670
   Begin VB.ListBox lstSortedX 
      Height          =   1020
      ItemData        =   "winmine.frx":3674C
      Left            =   9240
      List            =   "winmine.frx":3674E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblMinesLeft 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mines Left : 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   5790
      Width           =   8655
   End
   Begin VB.Image imgOpenBlocks 
      Height          =   240
      Left            =   2520
      Picture         =   "winmine.frx":36750
      Top             =   6000
      Width           =   8640
   End
   Begin VB.Image imgWrongMine 
      Height          =   240
      Left            =   10200
      Picture         =   "winmine.frx":38F92
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image imgQsPressed 
      Height          =   240
      Left            =   10200
      Picture         =   "winmine.frx":394D4
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image imgQuestion 
      Height          =   240
      Left            =   9960
      Picture         =   "winmine.frx":39A16
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image imgPressed 
      Height          =   240
      Left            =   10200
      Picture         =   "winmine.frx":39F58
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image imgFlag 
      Height          =   240
      Left            =   9960
      Picture         =   "winmine.frx":3A49A
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image imgBlown 
      Height          =   240
      Left            =   9960
      Picture         =   "winmine.frx":3A9DC
      Top             =   3000
      Width           =   240
   End
   Begin VB.Image imgMine 
      Height          =   240
      Left            =   10200
      Picture         =   "winmine.frx":3AF1E
      Top             =   3000
      Width           =   240
   End
   Begin VB.Image imgButton 
      Height          =   240
      Left            =   9960
      Picture         =   "winmine.frx":3B460
      Top             =   2760
      Width           =   240
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeginner 
         Caption         =   "&Beginner"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIntermediate 
         Caption         =   "&Intermediate"
      End
      Begin VB.Menu mnuExpert 
         Caption         =   "&Expert"
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "&Custom ..."
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuPlayingInstructions 
         Caption         =   "&Playing Instructions ..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAboutWinMine 
         Caption         =   "&About WinMine ..."
      End
   End
End
Attribute VB_Name = "frmWinMine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Instantiate an object of type class clsWinMine
' The Initialize Event of clsWinMine is called
Private objMine As New clsWinMine
Private Sub Form_Load()
    ' Supply the Main display form to the class object
    ' so that it can know which form to draw on etc.
    ' Property Set Procedure for clsWinMine is called
    Set objMine.frmDisplay = Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Determine which square was clicked and with
    ' which mouse button, and take action accordingly
    objMine.BeginHitTest Button, x, y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Determine over which square the mouse curser is
    ' at present while the left mouse button is pressed
    ' and take action accordingly
    objMine.TrackHitTest Button, x, y
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Determine over which square the mouse curser is
    ' when the left mouse button is released
    ' and take action accordingly
    objMine.EndHitTest Button, x, y
End Sub
Private Sub mnuAboutWinMine_Click()
    ' Display the About Box
    frmAboutBox.Show 1
End Sub
Private Sub mnuBeginner_Click()
    
    mnuBeginner.Checked = True
    mnuIntermediate.Checked = False
    mnuExpert.Checked = False
    mnuCustom.Checked = False

    ' Set the mine field dimensions for the Beginner
    ' level and get ready to start new game
    objMine.SetMineFieldDimension 8, 8, 10, False
    objMine.mblnNewGame = True
    
End Sub
Private Sub mnuCustom_Click()

    mnuBeginner.Checked = False
    mnuIntermediate.Checked = False
    mnuExpert.Checked = False
    mnuCustom.Checked = True

    ' Get the mine field dimensions for the Previous
    ' level to display as the default values in the dialog box
    objMine.GetMineFieldDimensions frmCustomDlg
    frmCustomDlg.Show 1
    
    ' Abort, if ESC key was pressed
    If frmCustomDlg.mblnEscape Then Exit Sub
    
    ' Set the mine field dimensions for the Desired level
    objMine.SetMineFieldDimension Val(frmCustomDlg.txtRows), Val(frmCustomDlg.txtColumns), Val(frmCustomDlg.txtMines), True
    
    ' Unload the hidden dialog, now that all values have been accessed
    Unload frmCustomDlg
    
    ' Get ready to start new game
    objMine.mblnNewGame = True

End Sub
Private Sub mnuExit_Click()
    ' Calls the terminate event for clsWinMine
    Set objMine = Nothing
    
    ' Exit the program
    End
End Sub
Private Sub mnuExpert_Click()

    mnuBeginner.Checked = False
    mnuIntermediate.Checked = False
    mnuExpert.Checked = True
    mnuCustom.Checked = False

    ' Set the mine field dimensions for the Expert
    ' level and get ready to start new game
    objMine.SetMineFieldDimension 16, 30, 100, False
    objMine.mblnNewGame = True

End Sub
Private Sub mnuIntermediate_Click()

    mnuBeginner.Checked = False
    mnuIntermediate.Checked = True
    mnuExpert.Checked = False
    mnuCustom.Checked = False

    ' Set the mine field dimensions for the Intermediate
    ' level and get ready to start new game
    objMine.SetMineFieldDimension 16, 16, 40, False
    objMine.mblnNewGame = True

End Sub
Private Sub mnuNew_Click()
    ' Prepare for starting a new game
    objMine.NewGame
End Sub
Private Sub mnuPlayingInstructions_Click()
    ' Display the Playing Instructions
    frmInstructBox.Show 1
End Sub
