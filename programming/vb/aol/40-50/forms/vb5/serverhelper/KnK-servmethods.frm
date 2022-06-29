VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   4590
   ClientLeft      =   1830
   ClientTop       =   0
   ClientWidth     =   4770
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form6"
   ScaleHeight     =   4590
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label26 
      BackColor       =   &H00000000&
      Caption         =   $"KnK-servmethods.frx":0000
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   4440
      TabIndex        =   25
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   4440
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Caption         =   "!SN find x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackColor       =   &H00000000&
      Caption         =   "!SN send x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackColor       =   &H00000000&
      Caption         =   "!SN send list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Or any Server using;"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Caption         =   "èmbràcè sèrvèr"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   4560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   $"KnK-servmethods.frx":0118
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   4440
      TabIndex        =   19
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   0
      TabIndex        =   18
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "-SN status"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "-SN send x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "-SN find x"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "-SN send list"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Or any Server using;"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "válkyrie máil server"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "/SN Send Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Send Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Find X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Send X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Send List"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Or Any IM server using;"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "IM Server (AO-NiN Style)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "/SN Find X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "/SN Send X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "/SN Send List"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Any Normal server using;"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Chatroom Server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
Text1.text = Label26 + Label20
End Sub

Private Sub Label19_Click()
Unload Me
End Sub

Private Sub Text1_Change()

End Sub
