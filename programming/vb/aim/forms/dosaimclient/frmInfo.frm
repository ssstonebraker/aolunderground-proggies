VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buddy Info: "
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox WhoInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Info"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser InfoBrow 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   4048
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
   Begin VB.Label Label2 
      Caption         =   "Personal Profile:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label UsrStatus 
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label OnTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label WarnLevel 
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label StatusLbl 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label TimeLbl 
      Caption         =   "Online Time:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label WarnLbl 
      Caption         =   "Warning Level:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "User you wish to get info on:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Not WhoInfo.Text = "" Then
    frmInfo.Caption = "Buddy Info: " & WhoInfo.Text
    Call SendProc(2, "toc_get_info " & Chr(34) & WhoInfo.Text & Chr(34) & Chr(0))
  End If
End Sub

Private Sub Form_Load()
InfoBrow.Navigate ("about:blank")
End Sub

Private Sub whoinfo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    Command1_Click
  End If
End Sub
