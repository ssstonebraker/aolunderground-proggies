VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   LinkTopic       =   "Form3"
   ScaleHeight     =   1245
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   360
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Copyright © 1998"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "By (V)agic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Outer Limits Scrambler V 1.O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
FormOnTop Me
FormNotOnTop Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormOnTop Form1
End Sub

Private Sub Label1_Click(Index As Integer)
Unload Me
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub
