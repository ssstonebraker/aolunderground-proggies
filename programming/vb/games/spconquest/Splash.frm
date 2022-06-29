VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2370
   ClientLeft      =   255
   ClientTop       =   45
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Task1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1412.884
   ScaleMode       =   0  'User
   ScaleWidth      =   18146.12
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "This Product is CheezeWare, and Free to Distribute"
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   2520
      Picture         =   "Splash.frx":030A
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label5 
      Caption         =   "Email: Pb2012@mad.scientist.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "1999"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "By: Paul Bryan"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SouthQuest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu mnumenu 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' "SouthQuest v1.8" By Paul Bryan ^1999, Southpark Command & Conquer Engine

Private Sub Label1_Click()
Shell "start mailto:pb2012@mad.scientist.com"
End Sub
