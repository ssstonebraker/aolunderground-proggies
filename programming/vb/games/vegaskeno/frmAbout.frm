VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Keno"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8400
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4740
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1059"
   Begin VB.Image Image2 
      Height          =   4620
      Left            =   15
      Top             =   60
      Width           =   8340
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   255
      TabIndex        =   3
      Top             =   3030
      Width           =   8115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   465
      X2              =   8100
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmAbout.frx":00FB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   420
      TabIndex        =   2
      Top             =   1965
      Width           =   7845
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Version 3.03  02/16/1999 Mike Altmanshofer"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   5745
      TabIndex        =   1
      Top             =   75
      Width           =   2685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please send BUG reports and game requests to: Masher2@ibm.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   315
      TabIndex        =   0
      Top             =   3930
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   0
      Picture         =   "frmAbout.frx":01DA
      Top             =   0
      Width           =   5550
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Unload Me
End Sub
