VERSION 5.00
Begin VB.Form Form26 
   BorderStyle     =   0  'None
   Caption         =   "Mass Mailer"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   LinkTopic       =   "Form26"
   Picture         =   "Form26.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4560
      Top             =   840
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Moon Mass Mailer"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "People:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Messager for mail:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Start MM"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MM Bot"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Keep New"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AddRoom"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label13_Click()
Form26.WindowState = 1
End Sub

Private Sub Label14_Click()
Unload Me
End Sub

Private Sub Label4_Click()
List1.Clear
Text2 = ""
End Sub

Private Sub Label8_Click()
Unload Me
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Timer1_Timer()
Label11.Caption = Str(List1.ListCount)

End Sub
