VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Caption         =   "       ÇLëár                 ÇøLø)2§¹            By:                      PLaya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   10995
      Left            =   -1320
      TabIndex        =   20
      Top             =   0
      Width           =   10995
   End
   Begin VB.Label Label20 
      BackColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   19
      Top             =   5880
      Width           =   11000
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   5040
      Width           =   11000
   End
   Begin VB.Label Label18 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   11000
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   3360
      Width           =   11000
   End
   Begin VB.Label Label16 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   2520
      Width           =   11000
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   1680
      Width           =   11000
   End
   Begin VB.Label Label14 
      BackColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   11000
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11000
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Height          =   7800
      Left            =   9120
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Height          =   7800
      Left            =   8280
      TabIndex        =   10
      Top             =   -120
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Height          =   7800
      Left            =   7440
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Height          =   7800
      Left            =   6600
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   7800
      Left            =   5760
      TabIndex        =   7
      Top             =   -120
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   7800
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   7905
      Left            =   4080
      TabIndex        =   5
      Top             =   -120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   7800
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "        "
      Height          =   7800
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7800
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "              "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7005
      Left            =   720
      TabIndex        =   1
      Top             =   -120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7800
      Left            =   0
      TabIndex        =   0
      Top             =   -960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Do
Label1.Visible = True
Pause (0.01)
Label1.Visible = False
Label2.Visible = True
Pause (0.01)
Label2.Visible = False
Label3.Visible = True
Pause (0.01)
Label3.Visible = False
Label4.Visible = True
Pause (0.01)
Label4.Visible = False
Label5.Visible = True
Pause (0.01)
Label5.Visible = False
Label6.Visible = True
Pause (0.01)
Label6.Visible = False
Label7.Visible = True
Pause (0.01)
Label7.Visible = False
Label8.Visible = True
Pause (0.01)
Label8.Visible = False
Label9.Visible = True
Pause (0.01)
Label9.Visible = False
Label10.Visible = True
Pause (0.01)
Label10.Visible = False
Label11.Visible = True
Pause (0.01)
Label11.Visible = False
Label12.Visible = True
Pause (0.01)
Label12.Visible = False
Label13.Visible = True
Pause (0.01)
Label13.Visible = False
Label14.Visible = True
Pause (0.01)
Label14.Visible = False
Label15.Visible = True
Pause (0.01)
Label15.Visible = False
Label16.Visible = True
Pause (0.01)
Label16.Visible = False
Label17.Visible = True
Pause (0.01)
Label17.Visible = False
Label18.Visible = True
Pause (0.01)
Label18.Visible = False
Label19.Visible = True
Pause (0.01)
Label19.Visible = False
Label20.Visible = True
Pause (0.01)
Label20.Visible = False
Label21.Visible = True
Pause (0.01)
Label21.Visible = False
Label21.Visible = True
Pause (0.3)
Label21.Visible = False
Label21.Visible = True
Pause (0.3)
Label21.Visible = False
Label21.Visible = True
Pause (0.3)
Label21.Visible = False
Loop







End Sub

Private Sub Form_Click()
End


End Sub

Private Sub Form_Load()
Form1.Visible = True

Label1.Visible = False
Label2.Visible = False

Label3.Visible = False

Label4.Visible = False

Label5.Visible = False

Label6.Visible = False

Label7.Visible = False

Label8.Visible = False

Label9.Visible = False

Label10.Visible = False

Label11.Visible = False

Label12.Visible = False

Label13.Visible = False

Label14.Visible = False

Label15.Visible = False

Label16.Visible = False

Label17.Visible = False

Label18.Visible = False

Label19.Visible = False

Label20.Visible = False

Label21.Visible = False


End Sub

