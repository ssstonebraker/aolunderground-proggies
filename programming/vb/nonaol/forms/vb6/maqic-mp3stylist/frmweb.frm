VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "mp3 stylist : search the web"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   Icon            =   "frmweb.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3855
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   5318
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   720
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Text            =   "http://www.lycos.com/music"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "  mp3 sylist : search the web"
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   4200
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Go"
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
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmweb.frx":22A2
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   255
      Left            =   720
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dos32.FormNotOnTop Form1
dos32.FormOnTop Form3
Call click32.FadeBy2(Form3.Picture1, vbBabyBlue, vbNavyBlue)
Form3.WebBrowser1.Navigate ("http://www.lycos.com/music")
End Sub

Private Sub Image1_Click()
Form3.PopupMenu Form2.webrowser, 0, 4800, 20
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form3.MousePointer = 15

       Call ReleaseCapture
    Call SendMessage(Form3.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
  
Form3.MousePointer = 1
End Sub

Private Sub Label2_Click()
FormOnTop Form1
Unload Form3
End Sub

Private Sub Label3_Click()
Form3.WindowState = 1
End Sub

Private Sub Label4_Click()
Form3.WebBrowser1.Navigate (Text1.Text)
End Sub
