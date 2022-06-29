VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows TAB"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   8421504
      TabCaption(0)   =   "Home"
      TabPicture(0)   =   "tabfrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Picture1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Iexplorer"
      TabPicture(1)   =   "tabfrm.frx":22B2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "WebBrowser1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "tabfrm.frx":22CE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Image3"
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(3)=   "Text2"
      Tab(2).ControlCount=   4
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   0
         Picture         =   "tabfrm.frx":22EA
         ScaleHeight     =   435
         ScaleWidth      =   6435
         TabIndex        =   10
         Top             =   5040
         Width           =   6495
         Begin VB.Line Line1 
            X1              =   1080
            X2              =   1080
            Y1              =   0
            Y2              =   480
         End
         Begin VB.Label Label6 
            Caption         =   "MENU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.TextBox Text2 
         Height          =   3495
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "tabfrm.frx":458C
         Top             =   1800
         Width           =   6015
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   4
         Top             =   840
         Width           =   6255
         ExtentX         =   11033
         ExtentY         =   8070
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74880
         TabIndex        =   2
         Text            =   "Internet URL here"
         Top             =   480
         Width           =   5415
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   5640
         Picture         =   "tabfrm.frx":4728
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mp3 Stylist"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   5640
         Picture         =   "tabfrm.frx":69CA
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "by maqic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   -71760
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74640
         Picture         =   "tabfrm.frx":8C6C
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Windows Tabbed Shell for Windows"
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
         Height          =   375
         Left            =   -74160
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MY PC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   840
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5640
         Picture         =   "tabfrm.frx":AF0E
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   11520
         Left            =   0
         Picture         =   "tabfrm.frx":D1B0
         Top             =   480
         Width           =   15360
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
WebBrowser1.Navigate (Text1.Text)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlack
End Sub

Private Sub Image2_Click()
CommonDialog1.Filter = "Excutable Files | *.exe"
CommonDialog1.ShowOpen
FileName = CommonDialog1.FileName
If CommonDialog1.FileName = "" Then
Exit Sub
End If
Shell (FileName)
End Sub

Private Sub Image5_Click()
Shell (App.Path & "\mp3stylist.exe")
End Sub

Private Sub Label6_Click()
Form1.PopupMenu Form2.menu, 0
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbWhite
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlack
End Sub
