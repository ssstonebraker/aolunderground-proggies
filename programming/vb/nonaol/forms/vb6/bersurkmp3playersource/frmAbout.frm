VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Bersurk Mp3 Player"
   ClientHeight    =   3390
   ClientLeft      =   5340
   ClientTop       =   3465
   ClientWidth     =   5715
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2001 Renegade's World."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label6 
      Caption         =   "Visit Renegade's World!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      ToolTipText     =   "Visit Renegade's World!"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   $"frmAbout.frx":0442
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "AIM Contacts: ProSk8er120, MyNameIsRenegade"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Language: Visual Basic 6.0"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Author: Renegade(CHoushyani@msn.com)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bersurk Mp3 Player v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmExplore.Show
frmExplore.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
frmExplore.Enabled = False
frmMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmExplore.Show
frmExplore.Enabled = True
Unload Me
End Sub

Private Sub Label6_Click()
Shell ("start http://renegadesworld.cjb.net")
End Sub
