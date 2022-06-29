VERSION 5.00
Begin VB.Form Form32 
   BorderStyle     =   0  'None
   Caption         =   "45 Min Killer"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   LinkTopic       =   "Form32"
   Picture         =   "Form32.frx":0000
   ScaleHeight     =   1680
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "  _"
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
      TabIndex        =   4
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label4 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "45 Min Killer"
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
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Stop"
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
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Start"
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
Form32.WindowState = 1
End Sub
