VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Cartman's Eating-Range  - Instructions"
   ClientHeight    =   4032
   ClientLeft      =   36
   ClientTop       =   228
   ClientWidth     =   7428
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4032
   ScaleWidth      =   7428
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   612
      Left            =   2160
      TabIndex        =   2
      Top             =   3360
      Width           =   3252
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3132
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   6132
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Instructions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   492
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   4452
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Unload Me
    
End Sub
