VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      Height          =   2190
      Left            =   120
      Picture         =   "frmAbout.frx":1272
      ScaleHeight     =   2130
      ScaleWidth      =   2730
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2790
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "dosfx or dosrox"
      Height          =   195
      Index           =   8
      Left            =   960
      TabIndex        =   9
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "www.dosfx.com"
      Height          =   195
      Index           =   7
      Left            =   960
      TabIndex        =   8
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "chad@dosfx.com"
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   7
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Chad J. Cox"
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "AIM:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   360
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "WebSite:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  lblInfo(0).Caption = "This project was developed by Chad J. Cox of www.dosfx.com. This is not a full client as some of the protocol  and a few features were left out. The reason being that this is meant to be an example only and is in no way what I would consider to be a full client ready for release." & vbCrLf & "Special thanks goes out to Pre (pre@dosfx.com). He has worked just as hard on this protocol and without a little teamwork, this project might not have come about." & vbCrLf & "If you have any questions or comments, please feel free to contact me (with the understanding that I do not have the time to teach you the protocol)." & vbCrLf & vbCrLf & "Also be sure to visit us in #visualbasic on irc.otherside.com."
End Sub

