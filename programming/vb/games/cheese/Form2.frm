VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Cartman's Eating-Range - Info"
   ClientHeight    =   1812
   ClientLeft      =   36
   ClientTop       =   228
   ClientWidth     =   3696
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1812
   ScaleWidth      =   3696
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   492
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "by Richard Nicol  && Oliver Twardowski"
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   2892
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00F7935E&
      Caption         =   "Cartman's Eating-Range"
      BeginProperty Font 
         Name            =   "Simpson"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3012
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    
    Unload Me
    
End Sub

Private Sub Image1_Click()
    
    Unload Me
    
End Sub
