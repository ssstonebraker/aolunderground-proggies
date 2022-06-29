VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2700
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   2990
      _Version        =   327681
      Tabs            =   6
      Tab             =   5
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Intro"
      TabPicture(0)   =   "KnK-Options.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Exit"
      TabPicture(1)   =   "KnK-Options.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ascii"
      TabPicture(2)   =   "KnK-Options.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Theme"
      TabPicture(3)   =   "KnK-Options.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "AOL"
      TabPicture(4)   =   "KnK-Options.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Scroll"
      TabPicture(5)   =   "KnK-Options.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Scroll Advertisor"
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2295
         Begin VB.OptionButton Option12 
            Caption         =   "No"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option11 
            Caption         =   "YES"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "AOL Version"
         Height          =   855
         Left            =   -74880
         TabIndex        =   15
         Top             =   720
         Width           =   2295
         Begin VB.OptionButton Option10 
            Caption         =   "AOL4.o"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option9 
            Caption         =   "AOL95"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Themes"
         Height          =   855
         Left            =   -74880
         TabIndex        =   10
         Top             =   660
         Width           =   2295
         Begin VB.OptionButton Option8 
            Caption         =   "Theme 2"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Theme 1"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ascii Color"
         Height          =   855
         Left            =   -74880
         TabIndex        =   7
         Top             =   660
         Width           =   2295
         Begin VB.OptionButton Option6 
            Caption         =   "BlackGreenBlack"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option5 
            Caption         =   "BlueBlackBlue"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show Exit"
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   -74880
         TabIndex        =   4
         Top             =   660
         Width           =   2295
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Yes"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Load Intro"
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   -74880
         TabIndex        =   1
         Top             =   660
         Width           =   2295
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Yes"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   615
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "When you click on an option, It is automaticly saved"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form10.PopupMenu Form10.help, 5

End Sub

Private Sub Form_Load()
StayOnTop Me



'Exit ini option
Loads2$ = GetFromINI("Exit", "Loads2", App.Path + "\KnK.ini")
If Loads2$ = "no" Then
Option3 = True
End If
If Loads2$ = "yes" Then
Option4 = True
End If

'Ascii ini option
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
If Color$ = "bbb" Then
Option5 = True
End If
If Color$ = "bgb" Then
Option6 = True
End If

'Theme ini option
KnKload$ = GetFromINI("KnKTheme", "KnKload", App.Path + "\KnK.ini")
If KnKload$ = "knk1" Then
Option7 = True
End If
If KnKload = "knk2" Then
Option8 = True
End If

Loads$ = GetFromINI("Intro", "Loads", App.Path + "\KnK.ini")
'Intro ini option
If Loads$ = "no" Then
Option2 = True
End If
If Loads$ = "yes" Then
Option1 = True
End If

'AOL ini option
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
Option9 = True
End If
If aversion$ = "aol4" Then
Option10 = True
End If
'Scroll ini option
adver$ = GetFromINI("Scroll", "adver", App.Path + "\KnK.ini")
If adver$ = "yes" Then
Option11 = True
End If
If adver$ = "no" Then
Option12 = True
End If


End Sub

Private Sub Option1_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("Intro", "Loads", "yes", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option10_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("AOL", "aversion", "aol4", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option11_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("Scroll", "adver", "yes", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option12_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("Scroll", "adver", "no", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option2_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("Intro", "Loads", "no", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option3_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("Exit", "Loads2", "no", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option4_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("Exit", "Loads2", "yes", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option5_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option6_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("ascii", "Color", "bgb", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option7_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("KnKTheme", "KnKload", "knk1", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub

Private Sub Option8_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("KnKTheme", "KnKload", "knk2", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If

End Sub

Private Sub Option9_Click()
On Local Error Resume Next
R% = WritePrivateProfileString("AOL", "aversion", "aol95", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
End Sub
