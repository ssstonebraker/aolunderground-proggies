VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Scrambler List Editor"
   ClientHeight    =   2535
   ClientLeft      =   900
   ClientTop       =   1050
   ClientWidth     =   3600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   2535
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "** For use with the ""*.scr"" list format only."
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   1320
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright© 1998"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   3135
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "By (V)agic"
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   3135
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Scrambler List Editor "
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   2175
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Text            =   "2"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove All"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Word"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   1740
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Word To List"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1510
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Create New List"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open Existing List"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save Current List"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4474
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":49B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5440
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5984
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":640C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6950
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":73D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":791C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu fileitem 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu fileitem 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu fileitem 
         Caption         =   "&Save As"
         Index           =   2
      End
      Begin VB.Menu fileitem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu fileitem 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu helpitem 
         Caption         =   "S&crambler List Help"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu helpitem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu helpitem 
         Caption         =   "&About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = "1"
If Text1.Text <> "" Then
List1.AddItem Text1.Text
Text1.Text = ""
Else
Msg = MsgBox("Please enter a word to add.", , "Error")
End If
End Sub

Private Sub Command2_Click()
Text2.Text = "1"
If List1.ListCount > 0 Then
List1.RemoveItem List1.ListIndex
Else
Msg = MsgBox("There are no words to remove.", , "Error")
End If
End Sub

Private Sub Command3_Click()
Text2.Text = "1"
Msg = MsgBox("Are you sure you want to clear the list?", vbYesNo, "Clear List")
If Msg = 6 Then List1.Clear
End Sub

Private Sub Command4_Click()
Label1.Visible = False
Label2.Visible = False
Label3(1).Visible = False
Label3(0).Visible = False
Command4.Visible = False
Frame1.Visible = False

End Sub

Private Sub fileitem_Click(index As Integer)
Select Case index
Case 0

Msg = MsgBox("All unsaved work will be lost.  Continue?", vbYesNo, "Create New List")

If Msg = 6 Then Text1.Text = "": List1.Clear

Case 1

Msg = MsgBox("All unsaved work will be lost.  Continue?", vbYesNo, "Open List")

If Msg = 6 Then
CommonDialog1.FLAGS = cdlOFNFileMustExist
CommonDialog1.Filter = "Scrambler Lists (*.scr)|*.scr"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
List1.Clear
Call Loadlistbox(CommonDialog1.FileName, List1)
End If
End If
Case 2
CommonDialog1.FLAGS = cdlOFNOverwritePrompt
CommonDialog1.Filter = "Scrambler Lists (*.scr)|*.scr"
CommonDialog1.ShowSave
Call SaveListBox(CommonDialog1.FileName, List1)
Text2.Text = "2"

Case 3

Case 4
If Text2.Text = "2" Then
Goodbye:
Unload Me
End
Else
Msg = MsgBox("All unsaved changes will be lost.  Do you wish to continue?", vbYesNo, "Warning")
If Msg = 6 Then GoTo Goodbye:
End If
End Select
End Sub

Private Sub Frame1_DragDrop(source As Control, X As Single, Y As Single)
Label1.Visible = False
Label2.Visible = False
Label3(1).Visible = False
Label3(0).Visible = False
Command4.Visible = False
Frame1.Visible = False

End Sub

Private Sub helpitem_Click(index As Integer)
Select Case index
Case 0
'CommonDialog1.HelpFile
Case 1
Case 2
Label1.Visible = True
Label2.Visible = True
Label3(1).Visible = True
Label3(0).Visible = True
Command4.Visible = True
Frame1.Visible = True
End Select
End Sub

Private Sub Label1_Click()
Label1.Visible = False
Label2.Visible = False
Label3(1).Visible = False
Label3(0).Visible = False
Command4.Visible = False
Frame1.Visible = False

End Sub

Private Sub Label2_Click()
Label1.Visible = False
Label2.Visible = False
Label3(1).Visible = False
Label3(0).Visible = False
Command4.Visible = False
Frame1.Visible = False

End Sub

Private Sub Label3_Click(index As Integer)
Select Case index
Case 0
Label1.Visible = False
Label2.Visible = False
Label3(1).Visible = False
Label3(0).Visible = False
Command4.Visible = False
Frame1.Visible = False
Case 1
Label1.Visible = False
Label2.Visible = False
Label3(1).Visible = False
Label3(0).Visible = False
Command4.Visible = False
Frame1.Visible = False
End Select
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then List1.AddItem Text1.Text: Text1.Text = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then List1.AddItem Text1.Text: Text1.Text = ""

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.index = 1 Then

Msg = MsgBox("All unsaved work will be lost.  Continue?", vbYesNo, "Create New List")

If Msg = 6 Then Text1.Text = "": List1.Clear
End If

If Button.index = 2 Then

Msg = MsgBox("All unsaved work will be lost.  Continue?", vbYesNo, "Open List")

If Msg = 6 Then
CommonDialog1.FLAGS = cdlOFNFileMustExist
CommonDialog1.Filter = "Scrambler Lists (*.scr)|*.scr"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
List1.Clear
Call Loadlistbox(CommonDialog1.FileName, List1)
End If
End If
End If

If Button.index = 3 Then
CommonDialog1.Filter = "Scrambler Lists (*.scr)|*.scr"
CommonDialog1.ShowSave
Call SaveListBox(CommonDialog1.FileName, List1)
Text2.Text = "2"
End If
End Sub

